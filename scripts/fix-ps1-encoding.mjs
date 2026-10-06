import fs from 'node:fs/promises';

/** 给 .ps1 加 UTF-8 BOM（Windows PowerShell 5.1 否则会按 ANSI/GBK 读取中文）。 */
const files = process.argv.slice(2);
for (const file of files) {
  const buf = await fs.readFile(file);
  const hasBom = buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf;
  const text = buf.toString('utf8').replace(/^\uFEFF/, '');
  const withEncoding = text.includes('[Console]::OutputEncoding')
    ? text
    : text.replace(
        /\$ErrorActionPreference = 'Stop'/,
        "$ErrorActionPreference = 'Stop'\ntry { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }",
      );
  await fs.writeFile(file, '\uFEFF' + withEncoding, 'utf8');
  console.log(`${file}: BOM ${hasBom ? 'already present' : 'added'}`);
}
