const sharp = require('sharp');
const pngToIco = require('png-to-ico').default;
const fs = require('fs');
const path = require('path');

const SRC = path.join(__dirname, 'LOGO.jpg');
const TMP_PNG = path.join(__dirname, '_icon_src_256.png');
const OUT = path.join(__dirname, 'icon.ico');

async function main() {
  await sharp(SRC).resize(256, 256, { fit: 'cover' }).png().toFile(TMP_PNG);
  const icoBuffer = await pngToIco(TMP_PNG);
  fs.writeFileSync(OUT, icoBuffer);
  fs.unlinkSync(TMP_PNG);
  console.log('Wrote', OUT);
}

main().catch(err => { console.error(err); process.exit(1); });
