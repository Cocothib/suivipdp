// Script de capture automatique pour la formation Suivi PDP/ICP
// Génère les 10 captures listées dans Liste_screenshots_a_prendre.txt
//
// Usage :
//   cd scripts
//   npm install
//   npm run capture
//
// Approche hybride : automation maximale, pauses interactives pour les vues
// qui nécessitent une mise en scène (signer un canvas, ouvrir une modale, etc.)

import { chromium } from 'playwright';
import { mkdir } from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import readline from 'readline';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUTPUT_DIR = path.join(__dirname, '..', 'captures_formation');
const USER_DATA_DIR = path.join(__dirname, '.playwright-profile');
const URL = process.env.SUIVIPDP_URL || 'https://zbpbasv.cluster121.hosting.ovh.net/suivipdp/';

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
const ask = (q) => new Promise(res => rl.question(q, res));

async function main() {
  await mkdir(OUTPUT_DIR, { recursive: true });

  console.log('\n=== Capture formation Suivi PDP/ICP ===');
  console.log('URL :', URL);
  console.log('Captures :', OUTPUT_DIR);
  console.log('Profil Chromium :', USER_DATA_DIR, '(conserve l\'auth SharePoint entre runs)\n');

  const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless: false,
    viewport: { width: 1600, height: 950 },
    deviceScaleFactor: 2,
    ignoreHTTPSErrors: true,
    args: ['--start-maximized', '--ignore-certificate-errors']
  });

  const page = context.pages()[0] || await context.newPage();
  await page.goto(URL, { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(2000);

  console.log('--- ÉTAPE 1/2 : Connexion ---');
  console.log('1) Connectez-vous à SharePoint dans la fenêtre Chromium si nécessaire (icône en haut).');
  console.log('2) Attendez que vos PDP/ICP soient synchronisés et visibles dans l\'app.');
  console.log('3) Revenez ici et appuyez sur Entrée pour démarrer les captures.\n');
  await ask('▶ Entrée pour continuer... ');

  const capture = async (filename, opts = {}) => {
    const filepath = path.join(OUTPUT_DIR, filename);
    await page.screenshot({ path: filepath, fullPage: !!opts.fullPage });
    console.log(`  ✓ ${filename}`);
  };

  const run = async (jsCode) => {
    try { return await page.evaluate(jsCode); } catch (e) { console.warn('  ⚠ JS:', e.message); }
  };

  const wait = (ms) => page.waitForTimeout(ms);

  const interactive = async (label) => {
    console.log(`\n--- ${label} ---`);
    await ask('▶ Mettez en scène puis Entrée pour capturer... ');
  };

  console.log('\n--- ÉTAPE 2/2 : Captures ---\n');

  // ============ SLIDE 7 : ICP - recherche Sellsy ============
  console.log('SLIDE 7 — Vue ICP, nouveau ICP, recherche Sellsy ouverte sur OPP');
  await run(`App.showView('icps'); App.newICP();`);
  await wait(1500);
  await interactive('Saisir un n° OPP dans le champ Recherche Sellsy et attendre l\'affichage de l\'opportunité');
  await capture('slide07_icp_nouvelle_recherche_sellsy.png');

  // ============ SLIDE 8 : Risques + mesures ============
  console.log('\nSLIDE 8 — Vue ICP/PDP onglet Risques avec catégories cochées et mesures saisies');
  await run(`App.showView('icps');`);
  await wait(800);
  await run(`(async () => { const list = await db.icps.toArray(); const icp = list.find(i => (i.risques||[]).length > 0) || list[0]; if (icp) await App.openICP(icp.id); })();`);
  await wait(1500);
  await run(`App.showICPSection && App.showICPSection('risques');`);
  await wait(800);
  await interactive('Si nécessaire, basculer manuellement sur l\'onglet Risques d\'un ICP ou d\'un PDP bien rempli');
  await capture('slide08_risques_mesures.png');

  // ============ SLIDE 9 : Signatures ICP ============
  console.log('\nSLIDE 9 — Vue ICP bloc Signatures, canvas avec signature en cours');
  await run(`App.showICPSection && App.showICPSection('signatures');`);
  await wait(800);
  await interactive('Dessiner une signature sur le canvas + remplir nom/fonction du signataire');
  await capture('slide09_icp_signatures_canvas.png');

  // ============ SLIDE 10 : ICP signé - bouton Générer PDP + modale ============
  console.log('\nSLIDE 10 — Vue ICP signé avec bouton "Générer PDP" + modale de confirmation');
  await interactive('Faire défiler jusqu\'au bouton "Créer un PDP depuis cette ICP" et cliquer dessus pour afficher la confirmation');
  await capture('slide10_generer_pdp_modale.png');

  // ============ SLIDE 11 : PDP barre 10 onglets ============
  console.log('\nSLIDE 11 — Vue PDP barre des onglets en haut, focus Général ou Risques');
  await run(`App.showView('list');`);
  await wait(800);
  await run(`(async () => { const list = await db.pdps.toArray(); const pdp = list.find(p => p.statut === 'actif' || p.statut === 'signature') || list[0]; if (pdp) await App.openPDP(pdp.id); })();`);
  await wait(1500);
  await run(`App.showPDPSection && App.showPDPSection('general');`);
  await wait(500);
  await interactive('Vérifier que la barre d\'onglets est bien visible en haut (ascenseur si besoin)');
  await capture('slide11_pdp_barre_onglets.png');

  // ============ SLIDE 12 : Permis de feu ============
  console.log('\nSLIDE 12 — Vue PDP onglet Permis de feu avec permis créé, 3 signatures et rondes');
  await run(`App.showPDPSection && App.showPDPSection('permisfeu');`);
  await wait(1000);
  await interactive('Vérifier qu\'un permis de feu existe avec ses signatures et le tableau d\'émargement des rondes');
  await capture('slide12_permis_feu.png');

  // ============ SLIDE 14 : Recherche Sellsy résultats ============
  console.log('\nSLIDE 14 — Vue ICP avec recherche Sellsy ayant retourné plusieurs opportunités');
  await run(`App.showView('icps'); App.newICP();`);
  await wait(1500);
  await interactive('Saisir un terme générique dans Sellsy (ex: nom client) pour obtenir une LISTE de résultats');
  await capture('slide14_sellsy_resultats.png');

  // ============ SLIDE 15 : Smartphone PWA ============
  console.log('\nSLIDE 15 — Vue mobile PWA avec PDP ouvert + photo chantier');
  await page.setViewportSize({ width: 390, height: 844 });
  await wait(500);
  await run(`(async () => { const list = await db.pdps.toArray(); const pdp = list.find(p => (p.visite?.photos||[]).length > 0 || (p.risques||[]).some(r => (r.photos||[]).length > 0)) || list[0]; if (pdp) await App.openPDP(pdp.id); })();`);
  await wait(1000);
  await interactive('Naviguer dans le PDP pour afficher une photo chantier (onglet Visite, Risques, ou Annexes)');
  await capture('slide15_mobile_pwa.png');
  // Restaurer viewport desktop
  await page.setViewportSize({ width: 1600, height: 950 });
  await wait(300);

  // ============ SLIDE 16 : Export PDP - 3 boutons PDF/DOCX/ZIP ============
  console.log('\nSLIDE 16 — Vue PDP onglet Export avec PDF, DOCX et ZIP visibles');
  await run(`(async () => { const list = await db.pdps.toArray(); if (list[0]) await App.openPDP(list[0].id); })();`);
  await wait(800);
  await run(`App.showPDPSection && App.showPDPSection('documents');`);
  await wait(500);
  await interactive('Vérifier que les 3 boutons Export Word, Export PDF et ZIP complet sont visibles');
  await capture('slide16_export_pdp.png');

  // ============ SLIDE 18 : Vue Aide - liste tutoriels ============
  console.log('\nSLIDE 18 — Vue Aide avec liste des 13 tutoriels et barre de recherche');
  await run(`App.showView('tutoriel');`);
  await wait(1000);
  await interactive('Vérifier que la liste des tutoriels et la barre de recherche sont visibles');
  await capture('slide18_aide_tutoriels.png');

  // ============ Fin ============
  console.log('\n=== Terminé ===');
  console.log('10 captures dans :', OUTPUT_DIR);
  console.log('Glissez-déposez chaque PNG dans le rectangle gris correspondant du PPTX.\n');
  await ask('▶ Entrée pour fermer le navigateur... ');
  await context.close();
  rl.close();
}

main().catch(err => {
  console.error('\n✗ Erreur :', err);
  rl.close();
  process.exit(1);
});
