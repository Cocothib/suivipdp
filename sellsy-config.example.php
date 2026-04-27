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
];
