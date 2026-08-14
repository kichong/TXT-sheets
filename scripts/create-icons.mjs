import fs from 'node:fs/promises';
import { resolve } from 'node:path';
import pngToIco from 'png-to-ico';
import sharp from 'sharp';

const source = resolve('build/icon.svg');
const png = resolve('build/icon.png');
const png256 = resolve('build/icon-256.png');
const ico = resolve('build/icon.ico');
const publicSvg = resolve('src/public/txt-sheets-logo.svg');

await sharp(source).resize(512, 512).png().toFile(png);
await sharp(source).resize(256, 256).png().toFile(png256);
await fs.writeFile(ico, await pngToIco([png256]));
await fs.copyFile(source, publicSvg);
