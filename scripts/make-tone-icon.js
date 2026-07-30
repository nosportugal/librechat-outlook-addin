/**
 * One-off generator for the "Change Tone" ribbon icon.
 *
 * Draws a white sliders/equalizer glyph (matching the line-art style of the
 * other ribbon icons) into assets/icon-tone-<size>.png at 16/32/80 px.
 *
 * Rendered at 8x supersampling then downscaled so thin lines stay crisp and
 * anti-aliased. Run with: node scripts/make-tone-icon.js
 */
const path = require("path");
const Jimp = require("jimp");

const SIZES = [16, 32, 80];
const SS = 8; // supersample factor
const BASE = 80; // master design size
const MASTER = BASE * SS; // 640
const WHITE = Jimp.rgbaToInt(255, 255, 255, 255);

function fillCircle(img, cx, cy, r, color) {
  for (let y = -r; y <= r; y++) {
    for (let x = -r; x <= r; x++) {
      if (x * x + y * y <= r * r) {
        const px = cx + x;
        const py = cy + y;
        if (
          px >= 0 &&
          py >= 0 &&
          px < img.bitmap.width &&
          py < img.bitmap.height
        ) {
          img.setPixelColor(color, px, py);
        }
      }
    }
  }
}

// Thick horizontal line with rounded caps.
function hLine(img, x0, x1, y, thickness, color) {
  const r = Math.round(thickness / 2);
  for (let x = x0; x <= x1; x++) {
    for (let dy = -r; dy <= r; dy++) {
      const py = y + dy;
      if (x >= 0 && py >= 0 && x < img.bitmap.width && py < img.bitmap.height) {
        img.setPixelColor(color, x, py);
      }
    }
  }
  fillCircle(img, x0, y, r, color);
  fillCircle(img, x1, y, r, color);
}

async function main() {
  const img = new Jimp(MASTER, MASTER, 0x00000000);

  const margin = Math.round(MASTER * 0.16);
  const x0 = margin;
  const x1 = MASTER - margin;
  const rail = Math.round(MASTER * 0.05); // rail thickness
  const knobR = Math.round(MASTER * 0.085); // knob radius
  const ringR = Math.round(MASTER * 0.038); // inner hole radius

  // 3 rails evenly spaced vertically.
  const ys = [0.3, 0.5, 0.7].map((f) => Math.round(MASTER * f));
  // Knob position along each rail (fraction from left).
  const knobF = [0.68, 0.35, 0.58];

  ys.forEach((y, i) => {
    hLine(img, x0, x1, y, rail, WHITE);
    const kx = Math.round(x0 + (x1 - x0) * knobF[i]);
    // White filled knob, then punch a transparent hole so it reads as a ring
    // over the rail (line-art look).
    fillCircle(img, kx, y, knobR, WHITE);
    fillCircle(img, kx, y, ringR, 0x00000000);
  });

  for (const size of SIZES) {
    const out = img.clone().resize(size, size, Jimp.RESIZE_BICUBIC);
    const file = path.join(__dirname, "..", "assets", `icon-tone-${size}.png`);
    await out.writeAsync(file);
    console.log(`[make-tone-icon] wrote ${file}`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
