// Renders public/assets/icon.svg to the PNG sizes the Office manifest needs.
// Edit icon.svg, then run: npm run icons
import sharp from "sharp";
import { readFileSync } from "node:fs";

const SIZES = [16, 32, 64, 80, 128];
const svg = readFileSync("public/assets/icon.svg");

for (const size of SIZES) {
  await sharp(svg, { density: 384 })
    .resize(size, size, { fit: "contain" })
    .png({ compressionLevel: 9 })
    .toFile(`public/assets/icon-${size}.png`);
}
console.log(`rendered ${SIZES.length} icons from public/assets/icon.svg`);
