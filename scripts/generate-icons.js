/**
 * Script pour générer les icônes PNG à partir des fichiers SVG
 * Nécessite sharp: npm install sharp --save-dev
 */

const fs = require('fs');
const path = require('path');

async function generateIcons() {
  try {
    const sharp = require('sharp');

    const sizes = [16, 32, 80];
    const assetsDir = path.join(__dirname, '..', 'assets');

    for (const size of sizes) {
      const svgPath = path.join(assetsDir, `icon-${size}.svg`);
      const pngPath = path.join(assetsDir, `icon-${size}.png`);

      if (fs.existsSync(svgPath)) {
        await sharp(svgPath)
          .png()
          .toFile(pngPath);
        console.log(`Généré: icon-${size}.png`);
      } else {
        console.warn(`SVG non trouvé: ${svgPath}`);
      }
    }

    // Créer logo-filled.png (copie de icon-80)
    const logoPath = path.join(assetsDir, 'logo-filled.png');
    if (fs.existsSync(path.join(assetsDir, 'icon-80.svg'))) {
      await sharp(path.join(assetsDir, 'icon-80.svg'))
        .png()
        .toFile(logoPath);
      console.log('Généré: logo-filled.png');
    }

    console.log('\nIcônes générées avec succès!');
  } catch (error) {
    if (error.code === 'MODULE_NOT_FOUND') {
      console.error('Le module sharp n\'est pas installé.');
      console.error('Exécutez: npm install sharp --save-dev');
      console.error('\nAlternativement, vous pouvez convertir manuellement les fichiers SVG en PNG.');
    } else {
      console.error('Erreur:', error.message);
    }
    process.exit(1);
  }
}

generateIcons();
