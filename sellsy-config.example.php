<?php
/**
 * Sellsy API credentials.
 *
 * 1. Copier ce fichier sur le serveur OVH sous le nom : sellsy-config.php
 *    (chemin : /www/suivipdp/sellsy-config.php)
 * 2. Renseigner client_id et client_secret depuis :
 *    Sellsy → Paramètres → Développeurs → "Créer une application"
 *    (type "Personal" / "Mes applications privées")
 *
 * IMPORTANT : sellsy-config.php est listé dans .gitignore et NE DOIT PAS
 * être commité sur GitHub. Il reste uniquement sur le serveur OVH.
 */
return [
    'client_id' => 'YOUR_SELLSY_CLIENT_ID',
    'client_secret' => 'YOUR_SELLSY_CLIENT_SECRET',

    // ---- Securite (recommande en production) ----
    // Domaine(s) exact(s) du site autorise(s) a lire les reponses via CORS.
    // L'app etant servie en same-origin que les proxys, ceci ne sert qu'a bloquer
    // les sites tiers. Laisser [] = aucun cross-origin autorise (le plus sur).
    // Exemple : 'allowed_origins' => ['https://suivipdp.mondomaine.fr'],
    'allowed_origins' => [],

    // Cle d'acces du proxy SuiviPDP (sellsy-proxy.php). Si renseignee, le front
    // doit l'envoyer (localStorage 'sellsyProxyKey' -> en-tete X-Api-Key).
    // 'proxy_access_key' => 'une-longue-chaine-aleatoire',

    // Cle d'acces du proxy Bilan Carbone (sellsy-bc.php) — donnees financieres.
    // 'bc_access_key' => 'une-autre-longue-chaine-aleatoire',
];
