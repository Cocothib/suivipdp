# Changelog SuiviPDP — historique simplifié

## 2026-09-18 — Glisser-déposer de documents et courriels dans un PDP ; correctifs de synchronisation (#21)

### Documents joints
- Zone de dépôt dans « Documents joints / Annexes » et dépôt possible n'importe où sur la fiche PDP ouverte (surcouche « Déposer pour joindre au PDP », bascule automatique sur la section Documents). Même traitement que le bouton Ajouter : catégorie du sélecteur, compression des images.
- **Courriels Outlook `.msg` et `.eml`** : lecture côté navigateur (`@kenjiuno/msgreader` / `postal-mime`, bundles ESM jsDelivr chargés à la demande) → un PDF lisible « Courriel - <objet>.pdf » (De / À / Cc / Date / Objet / pièces jointes + texte re-flué, jsPDF) annexé automatiquement au dossier PDF exporté, le fichier d'origine conservé (réouvrable dans Outlook, inclus dans le ZIP), et chaque pièce jointe ajoutée comme document à part entière (les images incorporées au corps — signatures, logos — sont ignorées). Nouvelle catégorie « Courrier / Courriel ». Un courriel illisible est joint tel quel.

### Synchronisation (analyse du journal d'activité du 21/08 au 15/09)
- **Création paresseuse et flush à la mise en veille marquent désormais la base « dirty »** (`_materializeNewForm`, `_flushCurrentForm`). Avant : la fiche vivait en IndexedDB sans flag → poussée vide par une sauvegarde déclenchée pour une autre raison (auto-conflits « moi contre moi » sur les nouvelles ICP 4478, 4516, 4523, 4525, 4527), et surtout, sur mobile, la saisie flushée au kill de la PWA pouvait être effacée par le clear+bulkPut du distant ou perdre le merge 3-way (vue comme « déjà synchronisée »).
- Conflits « sans référence » (fiche connue du distant mais absente du snapshot : poussée par un autre onglet/PWA du même utilisateur) : arbitrés par date comme avant, mais plus de toast alarmant ; le journal les distingue (`sansReference`).
- Hors ligne / 412 persistant : backoff exponentiel entre deux relances par le poll (8 s → 16 → 32 … 5 min), un seul toast et une seule ligne de journal par épisode (`conflit_persistant`), puis `sauvegarde_retablie`. Avant : un toast de 12 s et une ligne toutes les 8 s pendant toute la coupure.

### Journal d'activité
- `pdp.create` / `icp.create` de nouveau journalisés (la création paresseuse donnait un id avant le premier Enregistrer → jamais « nouveau »).
- File d'envoi mise en miroir dans localStorage : plus de perte des événements au rechargement / kill de la PWA, événements hors ligne conservés et envoyés à la reconnexion, nouvelle tentative 30 s après un échec d'envoi.
- `conflit_merge` : un événement par fiche (liste des champs) au lieu d'un par champ ; `otherUser` = auteur du fichier de l'entité concernée (avant : celui du fichier PDP pour tout). `numero_dedup` : un événement agrégé par fusion (le 21/08, 3 942 lignes unitaires avaient saturé le journal). Erreurs de sauvegarde 401/403/autres journalisées (`sauvegarde_erreur`). Rétention 5 000 → 8 000 événements.

### Numérotation inter-postes (#22)
- Collision d'id Dexie entre postes (ICP-202609-4520 créée par deux techniciens le 14/09 → renumérotée 4521 après coup, photo uploadée sous deux dossiers Media) : à la création d'un PDP / d'une ICP (saisie, création paresseuse, création croisée PDP↔ICP, duplication), le poste connecté **réserve le prochain numéro** dans `SuiviPDP/counters.json` (écriture conditionnelle If-Match, relecture + nouvel essai sur 412, valeur ≥ max local) et l'utilise comme id local explicite → id et numéro identiques sur tous les postes, plus de renumérotation. Hors ligne ou en cas d'échec : auto-incrément local comme avant (journal `numero_non_reserve`), `_dedupeNumeros` reste en filet de sécurité. Champ `numeroReserve: true` sur les fiches concernées.

## 2026-09-15 — Restauration d'un PDP / d'une ICP depuis son rapport PDF

- Seules les copies PDF **archivées sur SharePoint** (`SuiviPDP/PDP/Exports`, `SuiviPDP/ICP/Exports` : copie déposée à l'export, archivage automatique des ICP, ré-archivage batch — indicateur `PdfRestore._archiveMode`) embarquent une pièce jointe JSON `suivipdp-<pdp|icp>-<numero>.json` contenant la fiche complète : photos et documents joints avec leur binaire (jusqu'à 20 Mo, sinon références seules), signatures, liens PDP↔ICP par uid. Ajoutée via pdf-lib après la génération jsPDF. Le fichier téléchargé, le ZIP et les pièces jointes de mail (destinés aux entreprises extérieures) n'en portent jamais. Poids : environ celui des photos de la fiche (+33 % de base64), soit typiquement +0,1 Mo par photo.
- Paramètres > « Restaurer un PDP / une ICP depuis un rapport PDF » : dépôt du PDF, lecture de la pièce jointe (pdf.js `getAttachments`), aperçu (numéro, titre, dates, entreprises, risques, photos, signatures), détection de la fiche existante par uid (Remplacer / Créer une copie) ou restauration directe. La fiche recréée garde son uid (clé de fusion inter-postes) ; les photos avec binaire perdent leur `ref` pour être réexternalisées par le socle B3/B5 ; les liens PDP↔ICP sont résolus par uid (lien inverse posé si absent ou pointant vers une fiche disparue).
- Interrupteur « Embarquer les données de restauration dans les copies PDF archivées sur SharePoint » (réglage `pdfEmbedRecord`, activé par défaut) : la pièce jointe augmente la taille des PDF et contient l'intégralité de la fiche — à désactiver pour alléger l'archivage.
- Les PDF archivés avant cette version ne sont pas restaurables (pas de relecture du texte imprimé).
- Module `PdfRestore` (avant la section INIT) ; test headless `pptest/test_pdp.js` (export blob/complet, relecture, restauration new/replace/copy, ICP, option désactivée, UI).

## 2026-09-09 — Proxy Sellsy : authentification Microsoft

- `sellsy-proxy.php` vérifie désormais le jeton d'identité Azure AD de l'utilisateur (signature RS256 via les clés du tenant, émetteur, audience = application Suivi*, tenant, expiration). L'en-tête `Authorization` étant retiré par PHP-CGI sur OVH, le jeton est aussi envoyé dans `X-Ms-Token`.
- Mode piloté par `sellsy-auth.txt` (commité, déployé avec l'app) : `log` = vérifie et journalise sans bloquer (mode actuel, période d'observation le temps que les téléphones mettent à jour le service worker), `microsoft` = jeton obligatoire, `off` = retour arrière immédiat.
- Action `co_all` : annuaire compact des sociétés (id, nom, type, SIREN, SIRET, NAF, archivée), cache disque 6 h ; action `auth_check` pour tester un jeton.
- `index.html` : `_getIdToken()` (MSAL, `forceRefresh` sur 401), `SellsyAPI._call` envoie le jeton et réessaie une fois ; message explicite si l'utilisateur n'est pas connecté.
- Utilisé aussi par SuiviMarché (même hébergement, même inscription Azure AD).


Évolutions de l'application SuiviPDP de mars à juillet 2026 (versions v1 à v194). SuiviPDP gère les Plans De Prévention (PDP) et les Inspections Communes Préalables (ICP) réalisés avec les entreprises extérieures. Sigles utilisés : EU = Entreprise Utilisatrice (le client), EE = Entreprise Extérieure (l'intervenant), OPP = numéro d'opportunité commerciale Sellsy, FDS = Fiche de Données de Sécurité, FR-01 = fiche réflexe environnement, CNPP = modèle officiel de permis de feu (assureur AXA/CNPP).

---

## Août 2026 (v210)

### Connexion automatique renforcée
- Quand le cache de connexion Microsoft est vide au démarrage (navigateur qui purge les données du site à la fermeture, profil géré), l'application tente désormais un SSO silencieux via la session Azure AD du navigateur avant de demander un clic sur le bouton nuage ; l'identifiant du dernier compte connecté est mémorisé pour cibler le bon compte (v210).
- Le journal d'activité trace le mode d'autoconnexion (cache / SSO / retour de redirection) pour diagnostiquer les postes où elle échoue (v210).

---

## Août 2026 (v208)

### Correctif critique : duplication massive des fiches
- Corrigé : au rechargement de la page avec des modifications en attente (flag persisté introduit en v207), la fusion tournait avec un snapshot vide et re-clonait TOUTE la base sous de nouveaux identifiants à chaque boot (92 ICP réelles → 1 743 copies, 30 PDP → 566, 2 500 entreprises → 66 660 ; incident du 21/08). Une fiche locale dont l'uid existe déjà côté serveur est désormais fusionnée champ à champ avec sa jumelle distante au lieu d'être dupliquée (v208).
- Auto-réparation : à chaque fusion, les clones stricts (contenu identique hors id/uid/numéro) sont purgés de façon déterministe — un poste encore pollué se nettoie seul au premier merge ; l'exemplaire conservé privilégie les photos locales, l'uid d'origine et le numéro le plus ancien (v208).
- Les liens PDP → ICP (icpId) et ICP → entreprise (entrepriseId du représentant EE) suivent désormais les réattributions d'identifiants lors des fusions, comme le faisaient déjà les liens ICP → PDP et PDP → entreprises (v208).

---

## Août 2026 (v207)

### Anti-perte de données
- Les saisies non synchronisées survivent désormais au rechargement de la page : le marqueur « modifications locales en attente » est persisté. Auparavant, après un échec de sauvegarde (réseau instable, conflit avec un collègue) suivi d'un F5, l'application écrasait silencieusement les saisies locales (dont les signatures) par la version serveur — cas signalé le 28/07 (v207).

### Signatures issues de l'ICP
- Nouvelle section « Signatures issues de l'ICP » (lecture seule) dans la rubrique Signatures du PDP : tous les visas recueillis lors de l'inspection commune préalable liée (représentants et participants EU/EE, entreprises extérieures) y sont affichés avec nom, fonction, société, date et image de signature. Le bloc est repris dans les exports PDF et Word, distinct des signatures du PDP (v207).
- Les signatures posées sur l'ICP ne sont plus recopiées dans les cases de signature du PDP : le PDP se signe en propre, la traçabilité ICP passe par la nouvelle section dédiée (v207).

---

## Juillet 2026 (v167 → v197)

### Intégration inter-apps QSE
- SuiviPDP publie désormais son référentiel (entreprises, chantiers, personnel avec habilitations, agences) pour les autres applications QSE, avec bouton « Publier maintenant » dans les Paramètres (v197).
- Nouvelle carte « Plan d'actions QSE (SuiviNC) » sur le tableau de bord : NC ouvertes, actions en cours et retards du registre central SuiviNC, en lecture seule depuis le référentiel partagé (v196).
- Bouton « Dupliquer » sur les ICP, comme pour les PDP (v195).

### Permis de feu
- Permis de feu rapide : derniers libellés alignés mot pour mot sur le PDF CNPP (« Lieu et emplacement du travail », actions essentielles complètes, « Actions complémentaires (s'aider de la liste au verso) » replacée avant les moyens, « Une ronde de sécurité est nécessaire ? ») (v205).
- Plus aucun texte tronqué sur le PDF CNPP : quand la réduction de police ne suffit plus (texte très long), la case passe automatiquement en multiligne — « Outillage et matériel » s'étend vers le haut sur 2-3 lignes, « Lieu », « Moyens de lutte » et « Moyens d'alerte » se replient dans leur propre cadre (v204).
- Signatures bien visibles sur le PDF CNPP : la signature est recadrée sur le tracé (les marges vides du cadre de saisie l'écrasaient dans la case), le trait est épaissi et l'image est ajustée sans déformation dans une case optimisée (v204).
- Refonte UX de la section : saisie guidée en 4 étapes dépliables (Travail par point chaud, Risques & prévention, Validité & rondes, Signataires) avec badge d'avancement par étape (« Complet ✓ », « 2/3 signatures »...) ; les 4 pavés d'alerte sont condensés en un bandeau « Règles clés » d'une ligne, dépliable ; pastilles « Signé ✓ / En attente » sur chaque signataire ; au clic sur « Générer », contrôle de complétude avec liste des manques et lien direct vers l'étape concernée (« Générer quand même » possible) (v203).
- Horodatage automatique des signatures du permis : au moment où le signataire trace ou valide sa signature, la date et l'heure de l'appareil (téléphone/PC) sont remplies automatiquement si les champs sont vides — une saisie manuelle n'est jamais écrasée ; chaque signataire horodate uniquement sa propre section (v201).
- Effacer une signature du permis efface aussi la date et l'heure de sa section (v202).
- Les trois blocs de signature du permis (Donneur d'ordre, Personne désignée pour la surveillance, Intervenant) affichent désormais le trio « Signature — Date — Heure » comme sur le PDF CNPP, et l'heure de signature est reportée sur le PDF (champs Heure jamais remplis auparavant) — PDP et permis rapide (v200).
- La rubrique « Intervenants » du formulaire est désormais le miroir exact du permis CNPP : type d'intervenant (Entreprise extérieure — raison sociale / Interne — service, case « Interne » enfin reportée avec le service sur le PDF), responsable d'intervention, opérateurs en « Nom/téléphone » regroupés au même endroit, et une seule signature d'intervenant ; blocs réordonnés comme le PDF (Donneur d'ordre → Surveillance → Intervenants) (v199).
- Suppression de la note explicative « Permis de feu officiel AXA / CNPP » en bas de la section permis de feu du PDP (v199).
- Les rubriques du formulaire (PDP et permis rapide) reprennent désormais la nomenclature exacte du permis CNPP : « Description du travail par point chaud », « Nature du travail », « Outillage et matériel », « Risques identifiés », « Actions essentielles / complémentaires », « Donneur d'ordre », « Intervenants » (v198).
- Nouveau champ « Autre — préciser » pour la nature du travail, reporté sur le PDF et dans le mail au conducteur de travaux (v198).
- Le texte saisi tient toujours dans les cases du PDF CNPP : taille de police ajustée automatiquement au cadre, risques identifiés sur une ligne par catégorie, retour à la ligne automatique pour les actions complémentaires — plus aucun texte coupé (v198).
- Notification automatique par mail au conducteur de travaux à chaque permis de feu créé, avec le PDF rempli en pièce jointe et un lien SharePoint direct (v187-v190).
- Les exports de dossier PDP incluent désormais le permis de feu rempli en 1 exemplaire plus 3 permis vierges à remplir sur le chantier, mis en page pour l'impression recto-verso (v191-v194).

### Signatures
- Nouveau mode « Signer en grand » en plein écran, plus confortable sur tablette et smartphone (v181).
- Une signature validée est verrouillée : plus de trace accidentelle possible ; bouton « Valider » sur toutes les cases de signature (v185-v186).
- Distinction claire entre le rôle (Contact EU / Représentant EE) et la fonction réelle de la personne (v184).
- Corrections : la signature du représentant EE de l'ICP ne disparaît plus dans le PDP ; plus de colonnes en double pour une même société à l'export (v179, v183).

### Performance & synchronisation
- Démarrage de l'application nettement plus rapide (chargement non bloquant) et synchronisation plus réactive entre postes (v167-v168).
- Photos stockées en référence : l'application est plus légère et plus fluide (v169).

### Exports
- Nouvelle section « Mesures de prévention générales » dans les exports Word et PDF (v182).
- Documents Word allégés (compression des pièces incrustées) et photos de risque plus jamais coupées en bas de page dans le PDF ICP (v170, v174).
- Fiche réflexe FR-01 auto-réparée si absente et rattrapage automatique de l'envoi SharePoint des FDS importées hors connexion (v176-v177).

### Intégration Sellsy
- Nouvel indicateur au tableau de bord : couverture ICP du parc de contrats actifs Sellsy, réconcilié avec le tableau par technicien (v171-v172).

### Divers
- Bouton « Dupliquer » sur un PDP pour repartir d'un dossier existant (v173).
- L'export DUERP (Document Unique) est remplacé par un lien vers l'application dédiée SuiviDUERP (v175).
- Le bandeau « nouvelle version disponible » rappelle d'enregistrer son travail avant de recharger (v180).

---

## Juin 2026 (v18 → v166)

### Permis de feu
- Refonte complète sur le modèle officiel AXA/CNPP : le permis est généré en remplissant directement le formulaire PDF officiel, signatures incrustées, cases cochées propres (v148-v165).
- Permis de feu « rapide » créé depuis le tableau de bord sans passer par un PDP, avec archivage SharePoint centralisé (v155, v159).
- Un permis vierge CNPP est annexé d'office à chaque PDP ; historique par période avec gestion des avenants (v139, v162).
- Saisie allégée : champs obligatoires CNPP (moyens de lutte et d'alerte), génération automatique du permis rempli à l'enregistrement (v156-v158).

### Synchronisation & fiabilité des données
- Grand chantier anti-perte de données issu d'un audit complet : fusion intelligente des saisies simultanées de plusieurs postes, protections contre les écrasements concurrents, fichiers volumineux sécurisés (v53-v66).
- Protection de la saisie en cours : plus de rechargement automatique pendant qu'un formulaire est ouvert, enregistrement automatique renforcé (v88, v111).
- Numéros ICP/PDP garantis uniques même quand plusieurs postes créent des dossiers en même temps (v166).
- Synchronisation multi-postes durcie : fusion des FDS distantes, journal d'activités préservé en cas d'erreur de lecture (v90, v165).
- Mode « Travail en cours » : bandeau d'avertissement affiché sur tous les postes pendant une maintenance de l'application (v46-v51).
- Messages d'erreur SharePoint explicites (session expirée, accès en lecture seule…) pour faciliter le diagnostic (v31-v32).

### Sécurité
- Correction d'une faille d'injection via les noms d'utilisateurs et durcissement des accès au connecteur Sellsy (v64).
- Onglet Paramètres réservé aux administrateurs (v78).

### Photos & documents joints
- Prise de photos guidée par étapes avec légendes, et exports Word/PDF regroupés par étape (v67-v68).
- Photos et documents joints stockés en fichiers séparés sur SharePoint au lieu du fichier central : application beaucoup plus légère et rapide (v122-v132).
- Bibliothèque de documents par défaut, fiches réflexes incrustées automatiquement dans les exports, ajout rétroactif sur les PDP existants (v140-v147).

### Intégration Sellsy
- Recherche d'opportunité (OPP) directement dans le champ N° d'opération, avec autocomplétion sur ICP et PDP (v21-v24).
- Recherche accélérée : interrogation à la frappe, cache local des opportunités conservé 7 jours (moins d'attente, moins d'appels) (v112-v118).

### Tableau de bord & suivi maintenance
- Bloc « Avancement ICP » par technicien de maintenance, avec vue par mois et sélecteur d'année (v70-v73).
- Trophées maintenance (qualité et régularité), configurables par les administrateurs, avec notification de déblocage (v80-v87).
- Filtres de la liste ICP par technicien et par type Maintenance / Travaux (v86, v91).
- Numéro de version de l'application visible dans l'en-tête (v41-v43).

### Gestion des risques
- Éditeur de modèles de risques pour les administrateurs : édition, recherche, cotation par défaut, classement par catégorie (v95-v110).
- Proposition de risque possible directement depuis le tableau de bord, avec photos ; l'auteur est notifié par mail de la validation ou du rejet (v33, v95, v102).
- Nouveaux risques au catalogue : ligne électrique aérienne, chaleur intense / canicule (v97, v103).

### Exports
- Feuille de relevé des risques et observations en fin de dossier PDP (format paysage, colonne date) pour les annotations sur chantier (v18-v19).
- Encart Environnement (absorbants) et fiche réflexe FR-01 annexés aux exports (v19-v20).
- FDS intégrées au PDF et en annexe Word ; en cas d'export ZIP incomplet, confirmation explicite et fichier renommé « INCOMPLET_ » (v66, v89).
- Archivage SharePoint automatique des ICP dès le statut « Effectuée », re-archivage à chaque modification (v100).

### Répertoire & agences
- Pré-remplissage automatique de l'entreprise intervenante avec l'entité de l'agence (AGRIWATT/ENERSOLYS) (v25-v28).
- Fusion Agences + Sites en une source unique dans les Paramètres (v34).
- Enrichissement automatique du répertoire (entreprises, contacts, téléphones saisis) avec détection des doublons (v74-v76).

### Corrections & retours terrain
- 9 correctifs issus des retours terrain d'un conducteur de travaux (saisie, signatures, exports) (v120).
- Optimisations de performance et de connexion (démarrage, lectures conditionnelles) (v121-v137).
- Section ICP « délimitation, circulation et consignes » conforme au Code du travail (R4512-3/4).

---

## Mai 2026

### Ergonomie terrain
- Bouton flottant « Sauvegarder » toujours accessible, mode Express pour les ICP, vocabulaire simplifié et guide de prise en main.
- Fusion « champ par champ » des modifications simultanées de deux utilisateurs sur un même dossier.

### Signatures & émargement
- Bloc d'émargement de 20 lignes pour les visiteurs imprévus, tenant sur une seule page à l'export.
- Correction : les signatures EE étaient parfois perdues à l'enregistrement du PDP.

### Habilitations
- Gestion des habilitations des salariés avec génération des autorisations internes (visa repris de l'émargement), synchronisation automatique depuis le fichier SharePoint AGRIWATT.

### Risques & exports
- Photos intégrées par risque (ICP et PDP), affichées dans les exports ; cases EU/EE cochables mesure par mesure.
- Protection contre la coupure des chapitres en bas de page dans les exports.

### Agences & filtres
- 5 agences détaillées avec héritage automatique PDP ↔ ICP ; filtres Site et Conducteur de travaux sur les listes.

---

## Avril 2026 (v8 → v17)

### Synchronisation SharePoint
- Visualisation de l'état de la connexion SharePoint, journal d'activité (connexions, sauvegardes, modifications) et fusion anti-conflit des saisies concurrentes (v8-v12).
- Indication de présence : bandeau lorsqu'un autre utilisateur modifie le même dossier ; déploiement automatique de l'application sur le serveur OVH.

### Intégration Sellsy
- Connexion au CRM Sellsy : recherche d'opportunités depuis l'ICP et le PDP, récupération automatique du client, de l'adresse, des contacts et des coordonnées GPS.
- Création/mise à jour automatique de l'entreprise cliente dans le répertoire à l'enregistrement.

### Répertoire & contacts
- Contacts multiples par entreprise, recherche avec autocomplétion, participants additionnels EU/EE avec signatures individuelles pré-remplies.
- Propagation automatique des modifications vers le répertoire à chaque enregistrement.

### Exports & FDS
- Bibliothèque de FDS partagée avec rattachement automatique aux PDP ; exports ZIP complets (document + FDS + annexes).
- Noms de fichiers normalisés (OPP + numéro + intitulé) et archivage automatique Word + PDF à la clôture d'un PDP ou la validation d'une ICP.
- Recherche multi-champs et filtres (client, année, statut) sur les listes PDP et ICP.

### Gestion des risques
- Sélecteur de risques en arborescence, workflow de validation des nouveaux risques avec cotation F×G×M et export vers le Document Unique (v13-v17).

### Divers
- Géolocalisation automatique des chantiers (adresse, GPS, ouverture dans Google Maps).
- Émargement par signature à chaque ronde de surveillance du permis de feu.
- Distinction chantier interne / externe avec terminologie EU/EE adaptée ; améliorations d'accessibilité.

---

## Mars 2026 (v1 → v7)

### Naissance de l'application
- Première version : gestion des PDP et des visites d'inspection (devenues ICP) (v1).
- Refonte majeure : permis de feu complet, exports ZIP, documents joints, signatures tactiles (v5).
- Catalogue de risques pré-rempli (situations dangereuses et mesures de prévention), mesures cochables, pictogrammes INRS dans l'interface et les exports.
- ICP intégrées aux exports Word et PDF du PDP ; export « application autonome » avec données embarquées pour consultation hors ligne.
- Base des sites AGRIWATT/ENERSOLYS, compression des photos et des PDF, affichage adapté tablette et smartphone, guide d'utilisation complet (v6-v7).
