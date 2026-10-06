#!/usr/bin/env node
/** 由 PNG 生成多尺寸 .ico（Vista+ 支持内嵌 PNG）。 */
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createCanvas, loadImage } from "@napi-rs/canvas";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SIZES = [16, 24, 32, 48, 64, 128, 256];

export async function pngToIco(pngPath, icoPath, sizes = SIZES) {
  const source = await loadImage(await fs.readFile(pngPath));
  const images = [];
  for (const size of sizes) {
    const canvas = createCanvas(size, size);
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, size, size);
    ctx.drawImage(source, 0, 0, size, size);
    images.push({ size, data: await canvas.encode("png") });
  }
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(images.length, 4);
  const entries = [];
  let offset = 6 + images.length * 16;
  for (const image of images) {
    const entry = Buffer.alloc(16);
    entry.writeUInt8(image.size >= 256 ? 0 : image.size, 0);
    entry.writeUInt8(image.size >= 256 ? 0 : image.size, 1);
    entry.writeUInt8(0, 2);
    entry.writeUInt8(0, 3);
    entry.writeUInt16LE(1, 4);
    entry.writeUInt16LE(32, 6);
    entry.writeUInt32LE(image.data.length, 8);
    entry.writeUInt32LE(offset, 12);
    offset += image.data.length;
    entries.push(entry);
  }
  await fs.writeFile(icoPath, Buffer.concat([header, ...entries, ...images.map((i) => i.data)]));
  return icoPath;
}

const invokedDirectly = process.argv[1] && process.argv[1].replace(/\\/g, "/").endsWith("scripts/gen-ico.mjs");
if (invokedDirectly) {
  const input = process.argv[2] || path.join(ROOT, "addin/assets/icon-128.png");
  const output = process.argv[3] || path.join(ROOT, "addin/assets/app.ico");
  await pngToIco(input, output);
  const stat = await fs.stat(output);
  console.log(path.relative(ROOT, output) + " · " + Math.round(stat.size / 1024) + " KB");
}
