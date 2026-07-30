/**
 * Build-time generation of environment-labeled app icons.
 *
 * The same Docker image is deployed to every environment and the active
 * environment is chosen at runtime (see server.js / ICON_LABEL). Because the
 * runtime image ships no node_modules, we cannot draw labels on the fly there.
 * Instead we pre-generate labeled variants of the main app icons at build time
 * (e.g. icon-64-dev.png, icon-64-stb.png) and let server.js serve the right one
 * based on the runtime ICON_LABEL env var. PROD uses the original, unlabeled
 * icons.
 *
 * This module also exports makeLabeledIconBuffer so the webpack dev server can
 * generate the same badges on the fly during `npm start` / `npm run dev`.
 *
 * As a CLI (npm run build) it writes labeled variants into dist/assets.
 */
const path = require("path");
const fs = require("fs");
const {Jimp, loadFont, measureText, JimpMime} = require("jimp");

// Font files bundled with @jimp/plugin-print (v1 removed the FONT_SANS_* constants).
const FONT_DIR = path.join(
  path.dirname(require.resolve("@jimp/plugin-print")),
  "../..",
  "fonts/open-sans",
);
const FONT_PATH = (size, color) =>
  path.join(
    FONT_DIR,
    `open-sans-${size}-${color}`,
    `open-sans-${size}-${color}.fnt`,
  );

const DIST_ASSETS = path.join(__dirname, "..", "dist", "assets");

// Main app icon sizes (must match assets/icon-<size>.png and manifest.xml).
const SIZES = [16, 32, 64, 80, 128];

// Fixed label set. Keys become the filename/ICON_LABEL suffix (lowercase).
const LABELS = {
  dev: {text: "DEV", color: {r: 21, g: 101, b: 192}}, // blue
  stb: {text: "STB", color: {r: 230, g: 81, b: 0}}, // orange
  loc: {text: "LOC", color: {r: 56, g: 142, b: 60}}, // green
};

// Largest-to-smallest so we can pick the biggest font that fits the band.
const FONTS = [
  [FONT_PATH(64, "white"), 64],
  [FONT_PATH(32, "white"), 32],
  [FONT_PATH(16, "white"), 16],
  [FONT_PATH(8, "white"), 8],
];

const fontCache = new Map();
async function getFont(def) {
  if (!fontCache.has(def)) fontCache.set(def, await loadFont(def));
  return fontCache.get(def);
}

// Pick the largest white font whose rendered "DEV"/"STB" fits the badge.
async function pickFont(maxWidth, maxHeight, text) {
  for (const [def, height] of FONTS) {
    if (height > maxHeight) continue;
    const font = await getFont(def);
    const width = measureText(font, text);
    if (width <= maxWidth) return {font, width, height};
  }
  return null;
}

// Paint an opaque, color-filled rounded rectangle onto the image. Pixels
// outside the pill are left untouched so the icon keeps its transparency.
function fillRoundedRect(img, x0, y0, w, h, radius, color) {
  const {r, g, b} = color;
  const rad = Math.min(radius, Math.floor(w / 2), Math.floor(h / 2));
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      // Round the corners by testing distance from each corner's center.
      let cx = null;
      let cy = null;
      if (x < rad && y < rad) {
        cx = rad;
        cy = rad;
      } else if (x >= w - rad && y < rad) {
        cx = w - rad - 1;
        cy = rad;
      } else if (x < rad && y >= h - rad) {
        cx = rad;
        cy = h - rad - 1;
      } else if (x >= w - rad && y >= h - rad) {
        cx = w - rad - 1;
        cy = h - rad - 1;
      }
      if (cx !== null) {
        const dx = x - cx;
        const dy = y - cy;
        if (dx * dx + dy * dy > rad * rad) continue;
      }
      const px = x0 + x;
      const py = y0 + y;
      if (
        px < 0 ||
        py < 0 ||
        px >= img.bitmap.width ||
        py >= img.bitmap.height
      ) {
        continue;
      }
      const idx = img.getPixelIndex(px, py);
      img.bitmap.data[idx + 0] = r;
      img.bitmap.data[idx + 1] = g;
      img.bitmap.data[idx + 2] = b;
      img.bitmap.data[idx + 3] = 255;
    }
  }
}

async function makeLabeledIcon(srcPath, outPath, label) {
  const buffer = await makeLabeledIconBuffer(srcPath, label);
  await fs.promises.writeFile(outPath, buffer);
}

/**
 * Draw the environment badge on the source icon and return a PNG Buffer.
 * `label` may be a LABELS entry or a label key (e.g. "dev").
 *
 * The badge is a centered rounded "pill" near the bottom edge. Only the pill
 * itself is made opaque; the rest of the icon keeps its original transparency.
 */
async function makeLabeledIconBuffer(srcPath, label) {
  const resolved = typeof label === "string" ? LABELS[label] : label;
  if (!resolved) throw new Error(`Unknown icon label: ${label}`);

  const img = await Jimp.read(srcPath);
  const size = img.bitmap.width;

  // Largest font that fits within the badge footprint.
  const fit = await pickFont(
    Math.round(size * 0.82),
    Math.round(size * 0.5),
    resolved.text,
  );

  let pillW;
  let pillH;
  if (fit) {
    const padX = Math.max(3, Math.round(fit.height * 0.55));
    const padY = Math.max(2, Math.round(fit.height * 0.3));
    pillW = Math.min(size, fit.width + padX * 2);
    pillH = fit.height + padY * 2;
  } else {
    // Too small for legible text: draw a plain indicator bar instead.
    pillW = Math.round(size * 0.72);
    pillH = Math.max(4, Math.round(size * 0.24));
  }

  const margin = Math.max(1, Math.round(size * 0.04));
  const pillX = Math.round((size - pillW) / 2);
  const pillY = size - pillH - margin;
  const radius = Math.round(pillH / 2);

  fillRoundedRect(img, pillX, pillY, pillW, pillH, radius, resolved.color);

  if (fit) {
    const x = pillX + Math.round((pillW - fit.width) / 2);
    const y = pillY + Math.round((pillH - fit.height) / 2);
    img.print({font: fit.font, x, y, text: resolved.text});
  }

  return img.getBuffer(JimpMime.png);
}

async function main() {
  if (!fs.existsSync(DIST_ASSETS)) {
    console.error(
      `[label-icons] ${DIST_ASSETS} not found. Run the webpack build first.`,
    );
    process.exit(1);
  }

  let generated = 0;
  for (const size of SIZES) {
    const srcPath = path.join(DIST_ASSETS, `icon-${size}.png`);
    if (!fs.existsSync(srcPath)) {
      console.warn(`[label-icons] skipping missing ${srcPath}`);
      continue;
    }
    for (const [key, label] of Object.entries(LABELS)) {
      const outPath = path.join(DIST_ASSETS, `icon-${size}-${key}.png`);
      await makeLabeledIcon(srcPath, outPath, label);
      generated += 1;
    }
  }

  console.log(
    `[label-icons] generated ${generated} labeled icon(s) for: ${Object.keys(
      LABELS,
    ).join(", ")}`,
  );
}

// Run the file-writing CLI only when invoked directly (npm run build).
if (require.main === module) {
  main().catch((err) => {
    console.error("[label-icons] failed:", err);
    process.exit(1);
  });
}

module.exports = {LABELS, SIZES, makeLabeledIconBuffer};
