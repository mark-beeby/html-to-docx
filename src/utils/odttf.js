import crypto from 'crypto';
import fs from 'fs';

function formatGuid(buffer) {
  const hex = buffer.toString('hex').toUpperCase();
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(
    16,
    20
  )}-${hex.slice(20)}`;
}

function processGuid(guid) {
  // Remove braces and dashes
  const cleanGuid = guid.replace(/[{}-]/g, '');

  // Split into components
  const timeLow = cleanGuid.substr(0, 8); // 09014A78
  const timeMid = cleanGuid.substr(8, 4); // CABC
  const timeHi = cleanGuid.substr(12, 4); // 4EF0
  const clockSeq = cleanGuid.substr(16, 4); // 12AC
  const node = cleanGuid.substr(20, 12); // 5CD89AEFDE09

  // Convert components to bytes (no reverse)
  const timeLowBytes = Buffer.from(timeLow, 'hex');
  const timeMidBytes = Buffer.from(timeMid, 'hex');
  const timeHiBytes = Buffer.from(timeHi, 'hex');
  const clockSeqBytes = Buffer.from(clockSeq, 'hex');
  const nodeBytes = Buffer.from(node, 'hex');

  // Concatenate in correct order
  const guidBytes = Buffer.concat([
    Buffer.from([timeLowBytes[0]]), // 09
    nodeBytes,
    clockSeqBytes,
    timeHiBytes,
    timeMidBytes,
    timeLowBytes.slice(1),
  ]);

  return guidBytes;
}

function convertToODTTF(inputTTFPath) {
  const fontData = fs.readFileSync(inputTTFPath);
  const guid = `${formatGuid(crypto.randomBytes(16))}`;
  const guidBytes = processGuid(guid);
  const key = Buffer.concat([guidBytes, guidBytes]);

  const obfuscatedFont = Buffer.from(fontData);
  // eslint-disable-next-line no-plusplus
  for (let i = 0; i < 32; i++) {
    // eslint-disable-next-line no-bitwise
    obfuscatedFont[i] = fontData[i] ^ key[i];
  }

  return {
    guid,
    data: obfuscatedFont,
  };
}

// eslint-disable-next-line import/prefer-default-export
export { convertToODTTF };
