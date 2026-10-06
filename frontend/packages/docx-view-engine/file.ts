import fs from 'fs';
import { inflateRawSync } from 'zlib';

export function parseZip(filePath: string): Map<string, Buffer> {
  const buffer = fs.readFileSync(filePath);
  const files = new Map<string, Buffer>();
  let offset = 0;

  while (offset < buffer.length - 4) {
    if (buffer.readUInt32LE(offset) !== 0x04034b50) break;

    const compression = buffer.readUInt16LE(offset + 8);
    const compressedSize = buffer.readUInt32LE(offset + 18);
    const fileNameLength = buffer.readUInt16LE(offset + 26);
    const extraLength = buffer.readUInt16LE(offset + 28);

    const fileName = buffer.slice(offset + 30, offset + 30 + fileNameLength).toString('utf8');
    const dataOffset = offset + 30 + fileNameLength + extraLength;
    const compressedData = buffer.slice(dataOffset, dataOffset + compressedSize);

    if (compression === 0) {
      files.set(fileName, compressedData);
    } else if (compression === 8) {
      files.set(fileName, inflateRawSync(compressedData));
    }

    offset = dataOffset + compressedSize;
  }

  return files;
}

//Convert to px
export function ConvertPPI(twip: number): number {
  return Math.round(twip / 1440 * 96);
}