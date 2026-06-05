<?php
/**
 * Sellsy API v2 — passerelle LECTURE SEULE dédiée au Bilan Carbone (R3 / Climatip).
 * Hébergée sur OVH à côté de sellsy-proxy.php.
 *
 * INDÉPENDANT de l'application SuiviPDP :
 *   - ne modifie pas sellsy-proxy.php
 *   - réutilise simplement le même sellsy-config.php (mêmes identifiants)
 *   - réutilise le même cache de token (.sellsy-token.cache)
 * Supprimable sans aucun impact une fois le Bilan Carbone terminé.
 *
 * SÉCURITÉ :
 *   - Lecture seule : seuls GET et POST /{resource}/search sont autorisés.
 *     Aucune création / modification / suppression possible.
 *   - Ressources en liste blanche (voir $ALLOWED_RESOURCES).
 *   - Clé d'accès optionnelle : si 'bc_access_key' est défini dans sellsy-config.php,
 *     toute requête doit fournir ?key=... identique. (recommandé : factures = données sensibles)
 *
 * Endpoints :
 *   GET ?action=healthcheck
 *   GET ?action=get&path=/invoices/123                 -> GET Sellsy /invoices/123
 *   GET ?action=get&path=/invoices/123/rows            -> lignes d'une facture (produits)
 *   GET ?action=get&path=/items/45                     -> une fiche catalogue
 *   GET ?action=search&resource=invoices&limit=20&offset=0[&filters={...}]
 *   GET ?action=search&resource=items&limit=100
 *
 * Le paramètre filters est un JSON (urlencodé). Rappels API v2 :
 *   - le champ "filters" est OBLIGATOIRE sur /search (même vide => {}).
 *   - pas d'embed[] sur les GET /{resource}/{id}.
 */

// ---- Config (chargee tot pour la liste blanche CORS) ----
$configFile = __DIR__ . '/sellsy-config.php';
$cfg = file_exists($configFile) ? (require $configFile) : null;
$hasCreds = is_array($cfg) && !empty($cfg['client_id']) && !empty($cfg['client_secret']);

// ---- CORS ----
// On n'autorise que les origines listees dans sellsy-config.php (allowed_origins).
// Un import BC non-navigateur (script/curl) ignore le CORS ; la protection reelle
// de ce proxy reste la cle d'acces bc_access_key verifiee plus bas.
$allowedOrigins = (is_array($cfg) && !empty($cfg['allowed_origins']) && is_array($cfg['allowed_origins'])) ? $cfg['allowed_origins'] : [];
$origin = $_SERVER['HTTP_ORIGIN'] ?? '';
if ($origin !== '' && in_array($origin, $allowedOrigins, true)) {
    header('Access-Control-Allow-Origin: ' . $origin);
    header('Vary: Origin');
}
header('Access-Control-Allow-Methods: GET, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type, X-Api-Key');
header('Content-Type: application/json; charset=utf-8');
if (($_SERVER['REQUEST_METHOD'] ?? 'GET') === 'OPTIONS') { http_response_code(204); exit; }

// ---- Ressources autorisées (lecture seule) ----
$ALLOWED_RESOURCES = [
    'invoices',     // factures de vente -> CA + lignes produits (modules, onduleurs)
    'estimates',    // devis
    'orders',       // commandes / bons de commande
    'items',        // catalogue produits/services
    'taxes',
    'units',
    'payments',
    'companies',    // utile pour relier une facture à un client/site
    'individuals',
];

// ---- Config déjà chargée plus haut (pour le CORS) ----

// ---- Clé d'accès optionnelle ----
$accessKey = is_array($cfg) ? ($cfg['bc_access_key'] ?? null) : null;
$action = $_GET['action'] ?? '';

if ($action === 'healthcheck') {
    $tokenOk = false; $err = null;
    if ($hasCreds) {
        try { sellsy_get_token($cfg); $tokenOk = true; }
        catch (Exception $e) { $err = $e->getMessage(); }
    }
    echo json_encode([
        'ok' => true,
        'has_credentials' => $hasCreds,
        'token_ok' => $tokenOk,
        'access_key_required' => !empty($accessKey),
        'error' => $err,
    ]);
    exit;
}

// Vérif clé d'accès (si configurée)
if (!empty($accessKey)) {
    $provided = $_SERVER['HTTP_X_API_KEY'] ?? ($_GET['key'] ?? '');
    if (!hash_equals((string)$accessKey, (string)$provided)) {
        http_response_code(401);
        echo json_encode(['error' => 'cle d acces invalide ou manquante (en-tete X-Api-Key)']);
        exit;
    }
}

if (!$hasCreds) {
    http_response_code(503);
    echo json_encode(['error' => 'sellsy-config.php manquant ou incomplet sur le serveur']);
    exit;
}

try {
    $token = sellsy_get_token($cfg);
    switch ($action) {

        case 'get':
            // Passthrough GET en lecture seule, chemin validé.
            $path = $_GET['path'] ?? '';
            $path = is_string($path) ? trim($path) : '';
            if ($path === '' || $path[0] !== '/') {
                http_response_code(400);
                echo json_encode(['error' => 'path manquant ou invalide (doit commencer par /)']);
                exit;
            }
            // Sépare le chemin de l'éventuelle querystring pour valider la ressource.
            $resource = explode('/', ltrim($path, '/'))[0];
            $resource = explode('?', $resource)[0];
            if (!in_array($resource, $ALLOWED_RESOURCES, true)) {
                http_response_code(403);
                echo json_encode(['error' => "ressource non autorisee: $resource"]);
                exit;
            }
            if (preg_match('/\.\.|\s/', $path)) {
                http_response_code(400);
                echo json_encode(['error' => 'path invalide']);
                exit;
            }
            $res = sellsy_call('GET', $path, $token);
            break;

        case 'search':
            $resource = $_GET['resource'] ?? '';
            if (!in_array($resource, $ALLOWED_RESOURCES, true)) {
                http_response_code(403);
                echo json_encode(['error' => "ressource non autorisee: $resource"]);
                exit;
            }
            $limit = max(1, min(100, (int)($_GET['limit'] ?? 50)));
            $offset = max(0, (int)($_GET['offset'] ?? 0));
            // filters : JSON optionnel ; sinon {} (obligatoire en v2).
            $filters = new stdClass();
            if (!empty($_GET['filters'])) {
                $decoded = json_decode($_GET['filters'], true);
                if (is_array($decoded)) { $filters = $decoded; }
            }
            $body = ['filters' => $filters];
            $res = sellsy_call('POST', "/$resource/search?limit=$limit&offset=$offset", $token, $body);
            break;

        default:
            http_response_code(404);
            echo json_encode(['error' => 'action inconnue (utiliser healthcheck | get | search)']);
            exit;
    }
    echo $res;
} catch (Exception $e) {
    http_response_code(502);
    echo json_encode(['error' => $e->getMessage()]);
}

// ===========================================
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
        http_response_code($httpCode);
        return $resp; // remonte le JSON d'erreur Sellsy au client
    }
    return $resp;
}
