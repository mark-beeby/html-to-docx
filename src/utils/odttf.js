import fs from 'fs';

function formatGuid(buffer) {
  const hex = buffer.toString('hex').toUpperCase();
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(
    16,
    20
  )}-${hex.slice(20)}`;
}

function processGuid(guid) {
  // Remove GUID formatting characters
  const cleanGuid = guid.replace(/[{}-]/g, '');

  // Split GUID into its logical components
  const components = {
    fontNumber: cleanGuid.substr(0, 2), // XX
    timeLow: cleanGuid.substr(2, 6), // 014A78
    timeMid: cleanGuid.substr(8, 4), // CABC
    timeHigh: cleanGuid.substr(12, 4), // 4EF0
    clockSeq: cleanGuid.substr(16, 4), // 12AC
    node: cleanGuid.substr(20, 12), // 5CD89AEFDEXX
  };

  // Convert hex strings to byte arrays
  const bytes = {
    fontNumber: [...Buffer.from([parseInt(components.fontNumber, 16)])],
    timeLow: [...Buffer.from(components.timeLow, 'hex')],
    timeMid: [...Buffer.from(components.timeMid, 'hex')],
    timeHigh: [...Buffer.from(components.timeHigh, 'hex')],
    clockSeq: [...Buffer.from(components.clockSeq, 'hex')],
    node: [...Buffer.from(components.node, 'hex')],
  };

  // Create the 16-byte key array using array destructuring
  const key = [
    // Byte 0: Font number
    bytes.fontNumber[0],

    // Bytes 1-5: Node component (reversed order)
    bytes.node[4], // DE -> df
    bytes.node[3], // EF -> ef
    bytes.node[2], // 9A -> 9a
    bytes.node[1], // D8 -> d8
    bytes.node[0], // 5C -> 4c/4f

    // Bytes 6-7: Clock sequence
    bytes.clockSeq[1], // AC -> ad
    bytes.clockSeq[0], // 12

    // Bytes 8-9: Time high
    bytes.timeHigh[1], // f0
    bytes.timeHigh[0], // 4a

    // Bytes 10-11: Time mid
    bytes.timeMid[1], // bc
    bytes.timeMid[0], // ca

    // Bytes 12-14: Time low
    bytes.timeLow[2], // 3f
    bytes.timeLow[1], // 0e
    bytes.timeLow[0], // 44
  ];

  // Byte 15: Last byte pattern based on font number
  const lastByteMap = {
    1: 0x4c,
    2: 0x4f,
    3: 0x4e,
    4: 0x49,
    9: 0x4f,
  };

  // Add the final byte based on font number
  key.push(lastByteMap[bytes.fontNumber[0]] || 0x4f); // Default to 0x4f if font number not in map

  return Buffer.from(key);
}

function convertToODTTF(inputTTFPath) {
  const fontData = fs.readFileSync(inputTTFPath);
  // we've hard coded this GUID as the full algorithm has been difficult to fully reverse engineer.
  const guid = `${formatGuid('09014A78CABC4EF012AC5CD89AEFDE09')}`;
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
