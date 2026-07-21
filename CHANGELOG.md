# Changelog SuiviPDP — historique simplifié

Évolutions de l'application SuiviPDP de mars à juillet 2026 (versions v1 à v194). SuiviPDP gère les Plans De Prévention (PDP) et les Inspections Communes Préalables (ICP) réalisés avec les entreprises extérieures. Sigles utilisés : EU = Entreprise Utilisatrice (le client), EE = Entreprise Extérieure (l'intervenant), OPP = numéro d'opportunité commerciale Sellsy, FDS = Fiche de Données de Sécurité, FR-01 = fiche réflexe environnement, CNPP = modèle officiel de permis de feu (assureur AXA/CNPP).

---

## Juillet 2026 (v167 → v197)

### Intégration inter-apps QSE
- SuiviPDP publie désormais son référentiel (entreprises, chantiers, personnel avec habilitations, agences) pour les autres applications QSE, avec bouton « Publier maintenant » dans les Paramètres (v197).
- Nouvelle carte « Plan d'actions QSE (SuiviNC) » sur le tableau de bord : NC ouvertes, actions en cours et retards du registre central SuiviNC, en lecture seule depuis le référentiel partagé (v196).
- Bouton « Dupliquer » sur les ICP, comme pour les PDP (v195).

### Permis de feu
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
