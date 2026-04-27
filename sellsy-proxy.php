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
 *   GET ?action=list&limit=100&offset=0 -> Sellsy /companies (paginated)
 *   GET ?action=opp_search&q=...     -> Sellsy /opportunities/search
 *   GET ?action=opportunity&id=...   -> Sellsy /opportunities/{id}
 *
 * Credentials are read from sellsy-config.php (NOT committed).
 */

// ---- CORS ----
$allowedOrigins = ['*']; // tighten to your OVH domain in production
$origin = $_SERVER['HTTP_ORIGIN'] ?? '*';
if (in_array('*', $allowedOrigins, true) || in_array($origin, $allowedOrigins, true)) {
    header('Access-Control-Allow-Origin: ' . ($origin !== '' ? $origin : '*'));
    header('Vary: Origin');
}
header('Access-Control-Allow-Methods: GET, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');
header('Content-Type: application/json; charset=utf-8');
if (($_SERVER['REQUEST_METHOD'] ?? 'GET') === 'OPTIONS') { http_response_code(204); exit; }

// ---- Config ----
$configFile = __DIR__ . '/sellsy-config.php';
$cfg = file_exists($configFile) ? (require $configFile) : null;
$hasCreds = is_array($cfg) && !empty($cfg['client_id']) && !empty($cfg['client_secret']);

$action = $_GET['action'] ?? '';

if ($action === 'healthcheck') {
    $tokenOk = false;
    $err = null;
    if ($hasCreds) {
        try { sellsy_get_token($cfg); $tokenOk = true; }
        catch (Exception $e) { $err = $e->getMessage(); }
    }
    echo json_encode(['ok' => true, 'has_credentials' => $hasCreds, 'token_ok' => $tokenOk, 'error' => $err]);
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
            $res = sellsy_call('GET', "/companies/$id?embed[]=address&embed[]=contact&embed[]=phone_number&embed[]=email", $token);
            break;
        case 'contacts':
            $id = (int)($_GET['id'] ?? 0);
            if (!$id) { http_response_code(400); echo json_encode(['error' => 'id manquant']); exit; }
            $res = sellsy_call('GET', "/companies/$id/contacts?limit=20", $token);
            break;
        case 'opp_search':
            $q = trim($_GET['q'] ?? '');
            if ($q === '') { http_response_code(400); echo json_encode(['error' => 'q manquant']); exit; }
            // Sellsy v2 : POST /opportunities/search — filtre sur le nom (subject)
            $body = ['filters' => ['subject' => $q]];
            $res = sellsy_call('POST', '/opportunities/search?limit=20&embed[]=related&embed[]=step', $token, $body);
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
        default:
            http_response_code(404);
            echo json_encode(['error' => 'action inconnue']);
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
        return $resp; // forward error JSON to client
    }
    return $resp;
}
