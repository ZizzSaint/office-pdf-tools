/** 生成加载项图标（16/32/64/80/128 PNG）到 addin/assets 与 manifest/assets。 */
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createCanvas, GlobalFonts } from "@napi-rs/canvas";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SIZES = [16, 32, 64, 80, 128];

function roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

function drawIcon(size) {
  const canvas = createCanvas(size, size);
  const ctx = canvas.getContext("2d");
  const s = size / 100; // 归一化坐标

  // 背景
  const gradient = ctx.createLinearGradient(0, 0, size, size);
  gradient.addColorStop(0, "#e8590c");
  gradient.addColorStop(1, "#b02a0d");
  ctx.fillStyle = gradient;
  roundRect(ctx, 0, 0, size, size, 22 * s);
  ctx.fill();

  // 文档
  const pageW = 52 * s;
  const pageH = 66 * s;
  const px = (size - pageW) / 2;
  const py = (size - pageH) / 2 - 4 * s;
  const fold = 15 * s;
  ctx.fillStyle = "rgba(255,255,255,0.96)";
  ctx.beginPath();
  ctx.moveTo(px, py);
  ctx.lineTo(px + pageW - fold, py);
  ctx.lineTo(px + pageW, py + fold);
  ctx.lineTo(px + pageW, py + pageH);
  ctx.lineTo(px, py + pageH);
  ctx.closePath();
  ctx.fill();

  // 折角
  ctx.fillStyle = "rgba(0,0,0,0.14)";
  ctx.beginPath();
  ctx.moveTo(px + pageW - fold, py);
  ctx.lineTo(px + pageW, py + fold);
  ctx.lineTo(px + pageW - fold, py + fold);
  ctx.closePath();
  ctx.fill();

  // 文本线
  ctx.fillStyle = "rgba(43,87,154,0.55)";
  const lineX = px + 9 * s;
  let lineY = py + 30 * s;
  for (const width of [34, 30, 34, 22]) {
    ctx.fillRect(lineX, lineY, width * s, 3.4 * s);
    lineY += 8 * s;
  }

  // PDF 徽标
  if (size >= 32) {
    const bw = 44 * s;
    const bh = 17 * s;
    const bx = (size - bw) / 2;
    const by = size - bh - 9 * s;
    ctx.fillStyle = "#c43e1c";
    roundRect(ctx, bx, by, bw, bh, 5 * s);
    ctx.fill();
    ctx.fillStyle = "#ffffff";
    ctx.font = `bold ${11.5 * s}px "Segoe UI", Arial, sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("PDF", size / 2, by + bh / 2 + 0.5 * s);
  }
  return canvas;
}

async function main() {
  const targets = [
    path.join(ROOT, "addin", "assets"),
    path.join(ROOT, "manifest", "assets"),
  ];
  for (const dir of targets) await fs.mkdir(dir, { recursive: true });
  for (const size of SIZES) {
    const canvas = drawIcon(size);
    const png = await canvas.encode("png");
    for (const dir of targets) {
      await fs.writeFile(path.join(dir, `icon-${size}.png`), png);
    }
    if (size === 32) {
      await fs.writeFile(path.join(ROOT, "addin", "assets", "logo.png"), png);
    }
  }
  console.log(`icons generated: ${SIZES.join(", ")}`);
  void GlobalFonts;
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
