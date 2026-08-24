# SuiviPDP — Carte du projet

Application mono-fichier HTML (PWA) de gestion des Plans De Prévention (PDP) et Inspections Communes
Préalables (ICP) pour Agriwatt. Déployée sur GitHub Pages (Cocothib/suivipdp) ; stockage SharePoint via
Microsoft Graph (MSAL) + IndexedDB local (Dexie).

## MÉTHODE OBLIGATOIRE
`index.html` fait **1,8 Mo / 21 952 lignes** : ne JAMAIS le lire en entier. Grep pour localiser, puis
Read avec offset/limit. Les `index.backup-*.html` sont des sauvegardes historiques : ne pas les auditer
(seulement pour comparer un comportement avec la version actuelle).

## Fichiers principaux
| Fichier | Taille | Rôle |
|---|---|---|
| `index.html` | 1,8 Mo / 21 952 lignes | Toute l'application |
| `sw.js` | 119 lignes | Service worker — cache app `suivi-pdp-vNNN` (actuellement v210) + cache libs CDN séparé `suivi-pdp-libs-1` |
| `sellsy-proxy.php` / `sellsy-bc.php` | 14/10 Ko | Proxy PHP serveur vers API Sellsy |
| `manifest.json` | — | Manifest PWA |
| `CHANGELOG.md` | — | Historique des versions |
| `import-personnel.js`, `generate_*.py`, `scripts/` | — | Outillage annexe (non déployé) |

## Carte de index.html (numéros de ligne vérifiés, version v210 — 21/08/2026)
- **13–1270** : `<style>` CSS (thème Agriwatt)
- **1272–3841** : corps HTML (vues, modales, formulaires)
- **3842–3855** : CDN externes (voir Dépendances)
- **3857–21950** : `<script>` principal
  - 3858–3939 : DATABASE — Dexie `SuiviPDP`, schéma **v35** (voir plus bas)
  - 3941–3982 : cache court des lectures de tables
  - 3983–4132 : SHAREPOINT CONFIG, logo base64
  - 4133–4600 : MSAL + GRAPH API — `_uid()` (l.4283), externalisation binaires (strip/ré-hydratation
    base64 photos & documents), upload chunké Graph (contrôles #7)
  - 4601–4728 : `SharedRef` — référentiel partagé inter-apps QSE (`Shared/referentiels.json` sur site QHSE)
  - 4729–4928 : Lot 1 — publication du référentiel partagé (owner : suivipdp)
  - 4929–6720 : **DATASYNC — CŒUR SYNCHRO SharePoint** : etags (#8, If-None-Match), clamp d'horloge (#5),
    `_mergeStore` (l.5425), `_collapseClones` (l.5581), `_mergeByKey` (l.5706), fusion FDS dédiée (#14, l.5738),
    `_mergeAllWithRemote` (l.5789), `_loadedSnapshot` (l.5060/6192), auth MSAL, polling conditionnel,
    présence multi-postes (#15), sauvegardes séparées pdp/icp/shared
  - 6721–6912 : ACTIVITY LOG (journal connexions/sauvegardes/modifs)
  - 6913–7091 : SELLSY API (via proxy PHP)
  - 7092–8506 : AUTORISATION INTERNE / HABILITATIONS — template PDF, salariés, signataires, import Excel
    (parsers long l.7778 & wide Agriwatt l.7826), sync SharePoint (l.8000), génération PDF/depuis PDP
  - 8507–8894 : APP — navigation, routing hash, onboarding, toasts, confirm
  - 8895–9011 : DASHBOARD / indicateurs (couverture ICP parc Sellsy)
  - 9012–9377 : PDP LIST
  - 9378–14920 : PDP CRUD — formulaire, `RISK_CATALOG` (l.10297), risques (l.11130), EE, mesures, exigences,
    signatures, émargements, permis de feu, géocodage, photos, documents, cache média, entreprises (l.14068,
    import CSV l.14902), API Sellsy
  - 14921–18003 : ICP — participants, représentant EE, risques (l.15711), photos, signatures,
    liaison ICP↔PDP (`createPDPFromICP` l.15901, `createICPFromPDP` l.16117), risques personnalisés avec
    workflow de validation (l.16164), notifications mail, trophées, équipe maintenance
  - 18004–19189 : paramètres — sites (DEFAULT_SITES l.18005), bibliothèque FDS, archivage batch,
    gestion des risques unifiée (l.18489), référentiel partagé (l.18885), agences (l.18914), export/import
  - 19190–21865 : EXPORT DOCX / PDF / ZIP (ICP, permis de feu, PDP simplifié, données embarquées l.19130)
  - 21866–21950 : INIT, modales déplaçables, tutoriel, bootstrap de l'app

## Schéma Dexie (v35, l.3929)
```
pdps:  ++id, numero, titre, statut, dateDebut, dateFin, dateCreation, dateModification
entreprises: ++id, nom, siret
icps:  ++id, numero, site, dateVisite, pdpId, dateCreation
sites: ++id, nom, entite
fds:   ++id, nom, standard            (bibliothèque FDS, référencée par ID dans pdp.documents)
salaries_habilitations: ++id, &[nom+prenom], ...
signataires: ++id, &[nom+fonction], ...
media: &ref, ts                       (cache Blob des photos externalisées — APPEND-ONLY, JAMAIS vidé par la sync)
settings: key
```

## Structures de données clés (vérifiées dans le code)
**PDP** (l.15907) : `{ id, uid, numero ('PDP-AAAAMM-0000'), titre, statut, version, typePdp,
entite, agenceId, dateDebut/Fin, site, lieu, gps, description, materiel, nbSalaries, horaires,
eu:{nom,adresse,siret,responsable,tel,email,fonction}, entreprisesExt:[], visite:{date,heure,participants,
observations,effectuee,photos}, risques:[], mesures:[], instructions, exigences:[], urgence:{...},
environnement:{...}, pictogrammes:{interdictions,epiObligatoires,remarques}, signatureEU:{nom,date,data},
signaturesEE:[], modifications:[], historique:[], documents:[], icpId, refOperation, dateCreation, dateModification }`

**ICP** (l.16122) : `{ id, uid, numero ('ICP-AAAAMM-0000'), refOperation, statut, dateVisite, heure, site,
lieu, objet, entite, agenceId, euInfo, euParticipant:{nom,fonction,tel}, euParticipants:[], eeRepresentant
(porte entrepriseId), eeParticipants:[], entreprisesExt:[], risques:[], observations, photos:[],
signatureEU, signaturesEE:[], pdpId, dateCreation, dateModification }`

**Entreprise** (l.14068/14902) : `{ id, uid, nom, siret, adresse, tel, email, responsable, fonction,
activite, contacts:[] }` — les champs racine responsable/fonction/tel/email sont un miroir du contact
principal (rétro-compat v32).

**Risque** (l.11163) : `{ danger, situation, prevention, preventionItems:[], niveau:1-4, photos:[] }` —
catalogue standard `RISK_CATALOG` (l.10297 : `{cat, icon, color, risks:[]}`) ; risques personnalisés dans
`settings.customRisks` : `{ id, nom, categorie, situation, prevention, niveau, statut:'pending'|'approved'|'rejected',
proposedBy/At, reviewedBy/At, reviewComment }` (l.16164).

**Photos/documents** : binaires externalisés sur SharePoint (référence `ref`), Blob caché dans le store
`media`, base64 strippé avant sync et ré-hydraté à la demande.

## Technologies & CDN (l.3842–3855)
bootstrap 5.3.3 · dexie 3.2.4 · msal-browser 2.28.1 · jspdf 2.5.1 · jspdf-autotable 3.8.2 ·
html2canvas 1.4.1 · xlsx 0.18.5 · docx 8.5.0 · FileSaver 2.0.5 · jszip 3.10.1 · pdf-lib 1.17.1 ·
pdf.js 3.11.174. PHP côté serveur uniquement pour le proxy Sellsy.

## Conventions
- Tout en français (UI, commentaires, toasts). Objet global `App`, modules en objets littéraux
  (`DataSync`, `SharedRef`, `ActivityLog`...), méthodes privées préfixées `_`.
- Sections délimitées par des bannières `// ========` ; correctifs numérotés `// #N :` (#5 clamp horloge,
  #7 upload chunké, #8 etags, #10 modale bloquante, #14 fusion FDS, #15 échappement/présence).
- HTML injecté via template literals + `this.esc()` pour l'échappement.
- **Hook pre-commit** (`.git/hooks/pre-commit`) : bump auto de `suivi-pdp-vNNN` dans `sw.js` quand
  index.html ou manifest.json est modifié — ne pas bumper manuellement.
- `.gitignore` exclut les données RH/RGPD (`import-personnel.js`, `*-personnel.json`) et `sellsy-config.php`.

## PIÈGES CONNUS — invariants du merge à ne JAMAIS casser
**Incident duplication 21/08/2026 (v207 → correctif v208, commit 3cfb7af)** : un flag "dirty" persistant en
localStorage faisait tourner `_mergeAllWithRemote` au boot avec `_loadedSnapshot` vide → toutes les fiches
locales classées "ajouts locaux" → `_mergeStore` régénérait un uid (branche keyClash) → duplication totale
à chaque rechargement (92 ICP → 1743, 30 PDP → 566, entreprises → 66660).

Invariants :
1. Si un `uid` existe déjà côté distant, `_mergeStore` doit fusionner champ à champ, JAMAIS régénérer un uid.
2. `_loadedSnapshot` doit refléter l'état réellement chargé avant tout merge ; ne jamais merger avec un
   snapshot vide alors que des données locales existent.
3. Après merge : `_collapseClones` + propagation de `pdp.icpId` et `eeRepresentant.entrepriseId`
   (les id auto-incrémentés locaux divergent entre postes ; l'uid est la clé fiable).
4. Le store `media` est APPEND-ONLY : jamais vidé par le cycle clear+bulkPut de la sync.
5. Clamp d'horloge (#5) : `dateModification` vient de postes clients à l'horloge non fiable — ne pas
   tie-breaker sur des dates non clampées.
6. Restes connus post-incident : ~6 groupes d'ICP en double à vraies divergences, 3 liens `icpId=21` cassés,
   ~252 entreprises en doublon de nom à contenu divergent.
