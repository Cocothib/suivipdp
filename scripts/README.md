# Captures formation Suivi PDP/ICP

Script Playwright qui automatise la prise des 10 captures pour la présentation du 01/06/2026.

## Prérequis

- Node.js installé (vérifier : `node --version`)
- Accès internet (téléchargement Chromium au premier lancement)

## Installation (1 seule fois)

```bash
cd C:\Users\ThibaultHOCEDEZ\Documents\suivipdp\scripts
npm install
```

Cette commande installe Playwright + télécharge le navigateur Chromium (~150 Mo). Compter 2-3 min.

## Lancement

```bash
npm run capture
```

Le script :
1. Ouvre Chromium sur l'URL OVH de l'app
2. Demande de te connecter à SharePoint (une fois — le profil est conservé entre runs)
3. Pour chacune des 10 slides :
   - Navigue automatiquement vers la bonne vue
   - Affiche une instruction dans le terminal ("Mettre en scène...")
   - Attend que tu appuies sur **Entrée** pour capturer
4. Sauvegarde les PNG dans `../captures_formation/`

## Captures générées

| Slide | Fichier |
|-------|---------|
| 7 | slide07_icp_nouvelle_recherche_sellsy.png |
| 8 | slide08_risques_mesures.png |
| 9 | slide09_icp_signatures_canvas.png |
| 10 | slide10_generer_pdp_modale.png |
| 11 | slide11_pdp_barre_onglets.png |
| 12 | slide12_permis_feu.png |
| 14 | slide14_sellsy_resultats.png |
| 15 | slide15_mobile_pwa.png |
| 16 | slide16_export_pdp.png |
| 18 | slide18_aide_tutoriels.png |

Glisser-déposer chaque PNG dans le rectangle gris correspondant du PPTX.

## Options

URL personnalisée :
```bash
SUIVIPDP_URL=http://localhost:8000/suivipdp/ npm run capture
```

## Dépannage

- **Captures vides ou sélecteur introuvable** : vérifier que tu as bien des PDP/ICP réels en base. Le script utilise le premier PDP/ICP trouvé.
- **L'auth SharePoint se redemande à chaque fois** : ne pas supprimer le dossier `.playwright-profile`.
- **Bug d'affichage sur mobile (slide 15)** : si le rendu n'est pas convaincant, tu peux faire la capture sur ton propre smartphone et la remplacer ensuite.
