<?php
/**
 * Sellsy API v2 proxy for SuiviPDP
 * Hosted on OVH alongside index.html
 *
 * Endpoints:
 *   GET ?action=healthcheck          -> { ok, has_credentials, token_ok }
 *   GET ?action=search&q=...         -> Sellsy /companies/search
 *   GET ?action=company&id=...       -> Sellsy /companies/{id}
 *   GET ?action=contacts&id=...      -> Sellsy /companies/{id}/contacts
 *   GET ?action=addresses&id=...     -> Sellsy /companies/{id}/addresses
 *   GET ?action=contact&id=...       -> Sellsy /contacts/{id}
 *   GET ?action=list&limit=100&offset=0 -> Sellsy /companies (paginated)
 *   GET ?action=opp_search&q=...     -> Sellsy /opportunities/search
 *   GET ?action=opportunity&id=...   -> Sellsy /opportunities/{id}
 *   GET ?action=opp_all[&refresh=1]  -> liste complete agregee (champs reduits),
 *                                       cache disque 6h : 1 requete cote mobile
 *                                       au lieu de ~50 paginees
 *   GET ?action=co_all[&refresh=1]   -> annuaire compact des societes (id, nom, type,
 *                                       SIREN, SIRET, NAF, archivee), agrege cote
 *                                       serveur, cache disque 6h (utilise par SuiviMarche
 *                                       pour rapprocher les prospects par SIREN)
 *
 * Credentials are read from sellsy-config.php (NOT committed).
 *
 * Authentification Microsoft (2026-09) : chaque requete (sauf healthcheck et
 * auth_check) doit porter un jeton d'identite Azure AD valide de l'application
 * Suivi* (en-tete Authorization: Bearer <id_token>) : signature RS256 verifiee
 * avec les cles publiques du tenant (JWKS, cache 24h), emetteur, audience,
 * tenant et expiration controles. Mode pilote par le fichier sellsy-auth.txt
 * (a cote de ce script, commite) : "microsoft" = obligatoire, "log" = verifie
 * et journalise sans bloquer (observation), "off" = desactive (retour arriere).
 * Une cle proxy_access_key valide reste acceptee a la place (scripts).
 *   GET ?action=auth_check           -> { mode, ok, reason, user } pour tester un jeton
 */

// ---- Config (chargee tot pour la liste blanche CORS) ----
$configFile = __DIR__ . '/sellsy-config.php';
$cfg = file_exists($configFile) ? (require $configFile) : null;
$hasCreds = is_array($cfg) && !empty($cfg['client_id']) && !empty($cfg['client_secret']);

// ---- CORS ----
// L'app SuiviPDP appelle ce proxy en same-origin (sellsy-proxy.php relatif) :
// aucune en-tete CORS n'est alors necessaire. On n'autorise explicitement que les
// origines listees dans sellsy-config.php (allowed_origins) ; sinon aucun en-tete
// Access-Control-Allow-Origin n'est emis -> un site tiers ne peut PAS lire la reponse.
$allowedOrigins = (is_array($cfg) && !empty($cfg['allowed_origins']) && is_array($cfg['allowed_origins'])) ? $cfg['allowed_origins'] : [];
$origin = $_SERVER['HTTP_ORIGIN'] ?? '';
if ($origin !== '' && in_array($origin, $allowedOrigins, true)) {
    header('Access-Control-Allow-Origin: ' . $origin);
    header('Vary: Origin');
}
header('Access-Control-Allow-Methods: GET, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type, X-Api-Key, Authorization, X-Ms-Token');
header('Content-Type: application/json; charset=utf-8');
if (($_SERVER['REQUEST_METHOD'] ?? 'GET') === 'OPTIONS') { http_response_code(204); exit; }

// ---- Cle d'acces optionnelle (defense-in-depth) ----
// Si 'proxy_access_key' est defini dans sellsy-config.php, toute requete (sauf
// healthcheck) doit fournir la cle via l'en-tete X-Api-Key (ou ?key= en repli).
$accessKey = is_array($cfg) ? ($cfg['proxy_access_key'] ?? null) : null;

$action = $_GET['action'] ?? '';

if ($action === 'healthcheck') {
    $tokenOk = false;
    $err = null;
    if ($hasCreds) {
        try { sellsy_get_token($cfg); $tokenOk = true; }
        catch (Exception $e) { $err = $e->getMessage(); }
    }
    // #21 : ne PAS divulguer access_key_required (revele un proxy ouvert a un scan)
    // ni le message d'exception OAuth (detail interne) a un appelant non authentifie.
    if ($err) error_log('sellsy-proxy healthcheck OAuth: ' . $err);
    echo json_encode(['ok' => true, 'has_credentials' => $hasCreds, 'token_ok' => $tokenOk]);
    exit;
}

// ---- Authentification : jeton Microsoft (Azure AD) ou cle d'acces ----
$authMode = ms_auth_mode();
if ($action === 'auth_check') {
    $r = ms_auth_verify($cfg);
    echo json_encode(['mode' => $authMode, 'ok' => $r['ok'], 'reason' => $r['reason'], 'user' => $r['ok'] ? ms_mask($r['claims']['preferred_username'] ?? ($r['claims']['upn'] ?? '')) : null, 'exp' => $r['ok'] ? ($r['claims']['exp'] ?? null) : null]);
    exit;
}
$keyOk = false;
if (!empty($accessKey)) {
    $provided = $_SERVER['HTTP_X_API_KEY'] ?? ($_GET['key'] ?? '');
    $keyOk = hash_equals((string)$accessKey, (string)$provided);
}
if (!$keyOk && $authMode !== 'off') {
    $r = ms_auth_verify($cfg);
    if (!$r['ok']) {
        error_log('sellsy-proxy auth ' . $authMode . ' : ' . $r['reason'] . ' [' . $action . '] ' . ($_SERVER['REMOTE_ADDR'] ?? ''));
        if ($authMode === 'microsoft') {
            http_response_code(401);
            echo json_encode(['error' => 'connexion Microsoft requise', 'auth' => 'microsoft', 'reason' => $r['reason']]);
            exit;
        }
    }
} elseif (!$keyOk && !empty($accessKey) && $authMode === 'off') {
    // mode off avec une cle configuree : comportement historique (cle obligatoire)
    http_response_code(401);
    echo json_encode(['error' => 'cle d acces invalide ou manquante']);
    exit;
}

if (!$hasCreds) {
    http_response_code(503);
    echo json_encode(['error' => 'sellsy-config.php manquant ou incomplet sur le serveur']);
    exit;
}

try {
    $token = sellsy_get_token($cfg);
    switch ($action) {
        case 'search':
            $q = trim($_GET['q'] ?? '');
            if ($q === '') { http_response_code(400); echo json_encode(['error' => 'q manquant']); exit; }
            // Sellsy v2 search by name/SIRET via POST /companies/search
            $body = ['filters' => ['name' => $q]];
            $res = sellsy_call('POST', '/companies/search?limit=20', $token, $body);
            break;
        case 'list':
            $limit = max(1, min(100, (int)($_GET['limit'] ?? 50)));
            $offset = max(0, (int)($_GET['offset'] ?? 0));
            // Sellsy v2 exige le champ filters (même vide). new stdClass() force {} en JSON.
            $body = ['filters' => new stdClass()];
            $res = sellsy_call('POST', "/companies/search?limit=$limit&offset=$offset", $token, $body);
            break;
        case 'company':
            $id = (int)($_GET['id'] ?? 0);
            if (!$id) { http_response_code(400); echo json_encode(['error' => 'id manquant']); exit; }
            $res = sellsy_call('GET', "/companies/$id", $token);
            break;
        case 'contacts':
            $id = (int)($_GET['id'] ?? 0);
            if (!$id) { http_response_code(400); echo json_encode(['error' => 'id manquant']); exit; }
            $res = sellsy_call('GET', "/companies/$id/contacts?limit=20", $token);
            break;
        case 'addresses':
            $id = (int)($_GET['id'] ?? 0);
            if (!$id) { http_response_code(400); echo json_encode(['error' => 'id manquant']); exit; }
            $res = sellsy_call('GET', "/companies/$id/addresses?limit=20", $token);
            break;
        case 'contact':
            $id = (int)($_GET['id'] ?? 0);
            if (!$id) { http_response_code(400); echo json_encode(['error' => 'id manquant']); exit; }
            $res = sellsy_call('GET', "/contacts/$id", $token);
            break;
        case 'opp_search':
            $q = trim($_GET['q'] ?? '');
            if ($q === '') { http_response_code(400); echo json_encode(['error' => 'q manquant']); exit; }
            // Recherche par nom (subject). La recherche par numéro d'OPP se fait
            // côté client par accès direct à l'opportunité (action 'opportunity').
            // NB : embed[] est refusé par Sellsy sur /opportunities/search (400) ;
            // step/pipeline/status/related sont de toute façon inclus d'office.
            $body = ['filters' => ['subject' => $q]];
            $res = sellsy_call('POST', '/opportunities/search?limit=20', $token, $body);
            break;
        case 'opp_list':
            $limit = max(1, min(100, (int)($_GET['limit'] ?? 50)));
            $offset = max(0, (int)($_GET['offset'] ?? 0));
            // Liste de toutes les opportunités (filters obligatoire mais vide)
            $body = ['filters' => new stdClass()];
            $res = sellsy_call('POST', "/opportunities/search?limit=$limit&offset=$offset", $token, $body);
            break;
        case 'opportunity':
            $id = (int)($_GET['id'] ?? 0);
            if (!$id) { http_response_code(400); echo json_encode(['error' => 'id manquant']); exit; }
            // Sans embed[] (refusés sur cet endpoint en v2)
            $res = sellsy_call('GET', "/opportunities/$id", $token);
            break;
        case 'opp_all':
            // Liste complète des opportunités, agrégée côté serveur (bande passante
            // datacenter) et réduite aux champs utilisés par l'autocomplete client.
            // Cache disque 6h partagé entre tous les utilisateurs : le terrain fait
            // UNE requête (~300 Ko gzip) au lieu de ~50 paginées sur réseau mobile.
            $oppsCacheFile = __DIR__ . '/.sellsy-opps.cache';
            $maxAge = 6 * 3600;
            $force = !empty($_GET['refresh']);
            if (!$force && file_exists($oppsCacheFile) && (time() - filemtime($oppsCacheFile)) < $maxAge) {
                echo file_get_contents($oppsCacheFile);
                exit;
            }
            @set_time_limit(120); // ~50 appels Sellsy séquentiels côté serveur
            $all = [];
            $offset = 0; $limit = 100; $total = null;
            while (count($all) < 10000) {
                // Sans embed[] (refusé par Sellsy ici) : step/pipeline/status inclus d'office.
                $body = ['filters' => new stdClass()];
                $resp = sellsy_call('POST', "/opportunities/search?limit=$limit&offset=$offset", $token, $body);
                $page = json_decode($resp, true);
                if (!is_array($page) || empty($page['data']) || !is_array($page['data'])) break;
                if ($total === null && isset($page['pagination']['total'])) $total = (int)$page['pagination']['total'];
                foreach ($page['data'] as $o) { $all[] = sellsy_map_opp_compact($o); }
                if (count($page['data']) < $limit) break;
                $offset += $limit;
                if ($total !== null && count($all) >= $total) break;
            }
            $complete = ($total === null || count($all) >= $total);
            $res = json_encode([
                'items' => $all,
                'total' => $total !== null ? $total : count($all),
                'complete' => $complete,
                'cached_at' => time(),
            ]);
            // Ne jamais figer 6h une liste incomplète (rate-limit/timeout en cours de pagination)
            if ($complete && count($all) > 0) @file_put_contents($oppsCacheFile, $res);
            break;
        case 'co_all':
            // Annuaire compact de toutes les societes (~100 pages de 100), cache disque 6h.
            $cosCacheFile = __DIR__ . '/.sellsy-cos.cache';
            $maxAge = 6 * 3600;
            $force = !empty($_GET['refresh']);
            if (!$force && file_exists($cosCacheFile) && (time() - filemtime($cosCacheFile)) < $maxAge) {
                echo file_get_contents($cosCacheFile);
                exit;
            }
            @set_time_limit(280);
            $all = [];
            $offset = 0; $limit = 100; $total = null;
            while (count($all) < 30000) {
                $body = ['filters' => new stdClass()];
                $resp = sellsy_call('POST', "/companies/search?limit=$limit&offset=$offset", $token, $body);
                $page = json_decode($resp, true);
                if (!is_array($page) || empty($page['data']) || !is_array($page['data'])) break;
                if ($total === null && isset($page['pagination']['total'])) $total = (int)$page['pagination']['total'];
                foreach ($page['data'] as $c) {
                    $lf = (isset($c['legal_france']) && is_array($c['legal_france'])) ? $c['legal_france'] : [];
                    $all[] = [
                        'id' => $c['id'] ?? null,
                        'nom' => $c['name'] ?? '',
                        'type' => $c['type'] ?? '',
                        'siren' => $lf['siren'] ?? '',
                        'siret' => $lf['siret'] ?? '',
                        'naf' => $lf['ape_naf_code'] ?? '',
                        'archivee' => !empty($c['is_archived']),
                        'maj' => $c['updated_at'] ?? ($c['created'] ?? ''),
                    ];
                }
                if (count($page['data']) < $limit) break;
                $offset += $limit;
                if ($total !== null && count($all) >= $total) break;
            }
            $complete = ($total === null || count($all) >= $total);
            $res = json_encode([
                'items' => $all,
                'total' => $total !== null ? $total : count($all),
                'complete' => $complete,
                'cached_at' => time(),
            ]);
            if ($complete && count($all) > 0) @file_put_contents($cosCacheFile, $res);
            break;
        default:
            http_response_code(404);
            echo json_encode(['error' => 'action inconnue']);
            exit;
    }
    echo $res;
} catch (Exception $e) {
    // #25 : journaliser le detail cote serveur, message generique au client.
    error_log('sellsy-proxy: ' . $e->getMessage());
    http_response_code(502);
    echo json_encode(['error' => 'erreur proxy Sellsy']);
}

// ===========================================
// Projection compacte d'une opportunité Sellsy v2 — MÊME forme que le
// _mapOpportunity() du client (index.html) : le cache client les utilise tel quel.
function sellsy_map_opp_compact($o) {
    $related = (isset($o['related']) && is_array($o['related'])) ? $o['related'] : [];
    $companyId = $o['company_id'] ?? null;
    $companyName = '';
    foreach ($related as $r) {
        $t = $r['type'] ?? ($r['related_type'] ?? '');
        if ($t === 'company') {
            if (!$companyId) $companyId = $r['id'] ?? ($r['related_id'] ?? null);
            $companyName = $r['name'] ?? '';
            break;
        }
    }
    $step = (isset($o['step']) && is_array($o['step'])) ? $o['step'] : [];
    $estAmount = (isset($o['estimated_amount']) && is_array($o['estimated_amount'])) ? $o['estimated_amount'] : [];
    return [
        'id' => $o['id'] ?? null,
        'ref' => $o['reference'] ?? ($o['number'] ?? ('OPP-' . ($o['id'] ?? ''))),
        'nom' => $o['subject'] ?? ($o['name'] ?? ''),
        'montant' => $o['amount'] ?? ($estAmount['value'] ?? null),
        'devise' => $estAmount['currency'] ?? ($o['currency'] ?? 'EUR'),
        'etape' => $step['label'] ?? ($step['name'] ?? ($o['step_label'] ?? '')),
        'statut' => $o['status'] ?? '',        // open / closed / won / lost (cycle Sellsy)
        'pipeline' => (isset($o['pipeline']) && is_array($o['pipeline'])) ? ($o['pipeline']['name'] ?? '') : '',
        'probabilite' => $o['probability'] ?? ($step['probability'] ?? null),
        'dateCloture' => $o['estimated_closing_date'] ?? ($o['due_date'] ?? ''),
        'companyId' => $companyId,
        'companyName' => $companyName,
        'contactIds' => (isset($o['contact_ids']) && is_array($o['contact_ids'])) ? $o['contact_ids'] : [],
    ];
}

// ===========================================
// Authentification Microsoft (Azure AD, jeton d'identite v2.0 de l'application Suivi*)
function ms_auth_mode() {
    $f = __DIR__ . '/sellsy-auth.txt';
    $m = file_exists($f) ? strtolower(trim((string)file_get_contents($f))) : 'microsoft';
    return in_array($m, ['microsoft', 'log', 'off'], true) ? $m : 'microsoft';
}
function ms_mask($u) { $u = (string)$u; $at = strpos($u, '@'); return $at > 1 ? substr($u, 0, 2) . '…' . substr($u, $at) : ($u ? substr($u, 0, 2) . '…' : ''); }
function ms_b64url_decode($s) { $s = strtr($s, '-_', '+/'); return base64_decode(str_pad($s, strlen($s) % 4 ? strlen($s) + 4 - strlen($s) % 4 : strlen($s), '=', STR_PAD_RIGHT)); }
function ms_jwks($tenant, $forceRefresh = false) {
    $cacheFile = __DIR__ . '/.ms-jwks.cache';
    if (!$forceRefresh && file_exists($cacheFile) && (time() - filemtime($cacheFile)) < 24 * 3600) {
        $j = json_decode((string)file_get_contents($cacheFile), true);
        if (is_array($j) && !empty($j['keys'])) return $j['keys'];
    }
    $ctx = stream_context_create(['http' => ['timeout' => 10, 'header' => "User-Agent: sellsy-proxy/1.0\r\n"]]);
    $raw = @file_get_contents('https://login.microsoftonline.com/' . rawurlencode($tenant) . '/discovery/v2.0/keys', false, $ctx);
    $j = $raw ? json_decode($raw, true) : null;
    if (is_array($j) && !empty($j['keys'])) { @file_put_contents($cacheFile, $raw); return $j['keys']; }
    return [];
}
// cle publique PEM a partir des composantes RSA (n, e) d'une JWK
function ms_jwk_to_pem($jwk) {
    $n = ms_b64url_decode($jwk['n']); $e = ms_b64url_decode($jwk['e']);
    $der_len = function ($len) { if ($len < 128) return chr($len); $b = ltrim(pack('N', $len), "\0"); return chr(0x80 | strlen($b)) . $b; };
    $der_int = function ($x) use ($der_len) { if (ord($x[0]) > 0x7f) $x = "\0" . $x; return "\x02" . $der_len(strlen($x)) . $x; };
    $rsa = $der_int($n) . $der_int($e); $rsaSeq = "\x30" . $der_len(strlen($rsa)) . $rsa;
    $bit = "\x03" . $der_len(strlen($rsaSeq) + 1) . "\0" . $rsaSeq;
    $alg = "\x30\x0d\x06\x09\x2a\x86\x48\x86\xf7\x0d\x01\x01\x01\x05\x00";
    $spki = "\x30" . $der_len(strlen($alg . $bit)) . $alg . $bit;
    return "-----BEGIN PUBLIC KEY-----\n" . chunk_split(base64_encode($spki), 64, "\n") . "-----END PUBLIC KEY-----\n";
}
function ms_auth_verify($cfg) {
    $ms = (is_array($cfg) && !empty($cfg['ms_auth']) && is_array($cfg['ms_auth'])) ? $cfg['ms_auth'] : [];
    $tenant = $ms['tenant_id'] ?? '6487f869-e502-42be-91ab-5c1ecde34be6';
    $clients = !empty($ms['client_ids']) && is_array($ms['client_ids']) ? $ms['client_ids'] : ['ae0f1f6c-aedc-4866-899c-8d128e2af81a'];
    $h = $_SERVER['HTTP_AUTHORIZATION'] ?? ($_SERVER['REDIRECT_HTTP_AUTHORIZATION'] ?? '');
    if (!empty($_SERVER['HTTP_X_MS_TOKEN'])) $h = 'Bearer ' . $_SERVER['HTTP_X_MS_TOKEN'];   // en-tete de repli : Authorization est parfois retire par PHP-CGI (OVH mutualise)
    if ($h === '' && function_exists('apache_request_headers')) { $ah = apache_request_headers(); foreach ($ah as $k => $v) if (strtolower($k) === 'authorization') $h = $v; }
    if (!preg_match('/^Bearer\s+(\S+)$/i', trim($h), $m)) return ['ok' => false, 'reason' => 'jeton absent', 'claims' => null];
    $parts = explode('.', $m[1]); if (count($parts) !== 3) return ['ok' => false, 'reason' => 'jeton mal forme', 'claims' => null];
    $hdr = json_decode(ms_b64url_decode($parts[0]), true); $cl = json_decode(ms_b64url_decode($parts[1]), true); $sig = ms_b64url_decode($parts[2]);
    if (!is_array($hdr) || !is_array($cl)) return ['ok' => false, 'reason' => 'jeton illisible', 'claims' => null];
    if (($hdr['alg'] ?? '') !== 'RS256' || empty($hdr['kid'])) return ['ok' => false, 'reason' => 'algorithme inattendu', 'claims' => null];
    $now = time();
    if (!isset($cl['exp']) || $cl['exp'] < $now - 60) return ['ok' => false, 'reason' => 'jeton expire', 'claims' => null];
    if (isset($cl['nbf']) && $cl['nbf'] > $now + 300) return ['ok' => false, 'reason' => 'jeton pas encore valide', 'claims' => null];
    if (($cl['tid'] ?? '') !== $tenant) return ['ok' => false, 'reason' => 'tenant inattendu', 'claims' => null];
    if (($cl['iss'] ?? '') !== 'https://login.microsoftonline.com/' . $tenant . '/v2.0') return ['ok' => false, 'reason' => 'emetteur inattendu', 'claims' => null];
    if (!in_array((string)($cl['aud'] ?? ''), $clients, true)) return ['ok' => false, 'reason' => 'audience inattendue', 'claims' => null];
    $keys = ms_jwks($tenant); $jwk = null;
    foreach ($keys as $k) if (($k['kid'] ?? '') === $hdr['kid']) { $jwk = $k; break; }
    if (!$jwk) { $keys = ms_jwks($tenant, true); foreach ($keys as $k) if (($k['kid'] ?? '') === $hdr['kid']) { $jwk = $k; break; } }
    if (!$jwk || empty($jwk['n']) || empty($jwk['e'])) return ['ok' => false, 'reason' => 'cle de signature inconnue', 'claims' => null];
    $pub = openssl_pkey_get_public(ms_jwk_to_pem($jwk));
    if (!$pub) return ['ok' => false, 'reason' => 'cle publique illisible', 'claims' => null];
    $ok = openssl_verify($parts[0] . '.' . $parts[1], $sig, $pub, OPENSSL_ALGO_SHA256) === 1;
    return $ok ? ['ok' => true, 'reason' => 'ok', 'claims' => $cl] : ['ok' => false, 'reason' => 'signature invalide', 'claims' => null];
}

function sellsy_get_token($cfg) {
    $cacheFile = __DIR__ . '/.sellsy-token.cache';
    if (file_exists($cacheFile)) {
        $cached = json_decode(@file_get_contents($cacheFile), true);
        if (is_array($cached) && isset($cached['access_token'], $cached['expires_at']) && $cached['expires_at'] > time() + 30) {
            return $cached['access_token'];
        }
    }
    $ch = curl_init('https://login.sellsy.com/oauth2/access-tokens');
    $payload = http_build_query([
        'grant_type' => 'client_credentials',
        'client_id' => $cfg['client_id'],
        'client_secret' => $cfg['client_secret'],
    ]);
    curl_setopt_array($ch, [
        CURLOPT_POST => true,
        CURLOPT_POSTFIELDS => $payload,
        CURLOPT_HTTPHEADER => ['Content-Type: application/x-www-form-urlencoded'],
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_TIMEOUT => 15,
    ]);
    $body = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $err = curl_error($ch);
    curl_close($ch);
    if ($body === false) throw new Exception('OAuth network error: ' . $err);
    $data = json_decode($body, true);
    if ($httpCode >= 400 || empty($data['access_token'])) {
        throw new Exception('OAuth failed (' . $httpCode . '): ' . substr($body, 0, 200));
    }
    $expiresIn = (int)($data['expires_in'] ?? 3600);
    @file_put_contents($cacheFile, json_encode([
        'access_token' => $data['access_token'],
        'expires_at' => time() + $expiresIn,
    ]));
    return $data['access_token'];
}

function sellsy_call($method, $path, $token, $body = null) {
    $url = 'https://api.sellsy.com/v2' . $path;
    $ch = curl_init($url);
    $headers = [
        'Authorization: Bearer ' . $token,
        'Accept: application/json',
    ];
    $opts = [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_TIMEOUT => 20,
        CURLOPT_CUSTOMREQUEST => $method,
    ];
    if ($body !== null) {
        $headers[] = 'Content-Type: application/json';
        $opts[CURLOPT_POSTFIELDS] = json_encode($body);
    }
    $opts[CURLOPT_HTTPHEADER] = $headers;
    curl_setopt_array($ch, $opts);
    $resp = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $err = curl_error($ch);
    curl_close($ch);
    if ($resp === false) throw new Exception('Sellsy network error: ' . $err);
    if ($httpCode >= 400) {
        // #25 : ne pas forwarder le corps d'erreur Sellsy brut (peut contenir des
        // details internes). Journaliser cote serveur, renvoyer un message generique.
        error_log('sellsy-proxy Sellsy ' . $httpCode . ' ' . $method . ' ' . $path . ': ' . substr((string)$resp, 0, 500));
        http_response_code($httpCode >= 500 ? 502 : $httpCode);
        return json_encode(['error' => 'erreur API Sellsy', 'status' => $httpCode]);
    }
    return $resp;
}
