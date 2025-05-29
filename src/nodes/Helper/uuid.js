/**
 * @copyright https://antonz.org/uuidv7/#javascript
 */
export function uuidv7(getRaw = false) {
  // random bytes
  const value = Buffer.alloc(16);
  crypto.getRandomValues(value);

  // current timestamp in ms
  // eslint-disable-next-line no-undef
  const timestamp = BigInt(Date.now());

  // timestamp
  value[0] = Number((timestamp >> 40n) & 0xffn);
  value[1] = Number((timestamp >> 32n) & 0xffn);
  value[2] = Number((timestamp >> 24n) & 0xffn);
  value[3] = Number((timestamp >> 16n) & 0xffn);
  value[4] = Number((timestamp >> 8n) & 0xffn);
  value[5] = Number(timestamp & 0xffn);

  // version and variant
  value[6] = (value[6] & 0x0f) | 0x70;
  value[8] = (value[8] & 0x3f) | 0x80;

  return getRaw ? value : uuidToString(value);
}

// The IFC-specified base-64 character set for GUID compression
const BASE64_CHARS =
  "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$";

/**
 *
 * @param {Uint8Array} uuid
 */
export function uuidToString(uuid) {
  let result = "";

  // First byte -> two base64 chars
  result += BASE64_CHARS.charAt((uuid[0] >> 6) & 0x3f);
  result += BASE64_CHARS.charAt(uuid[0] & 0x3f);

  // Remaining 15 bytes in groups of 3 -> four base64 chars each
  for (let i = 1; i < 16; i += 3) {
    const triplet = (uuid[i] << 16) | (uuid[i + 1] << 8) | uuid[i + 2];
    // Use j instead of i for the inner loop
    for (let j = 18; j >= 0; j -= 6) {
      result += BASE64_CHARS.charAt((triplet >> j) & 0x3f);
    }
  }
  return result;
}

/**
 * Generates a new IFC GUID
 * @returns {string} A 22-character IFC GUID
 */
export function generateIfcGuid() {
  const uuid = uuidv7(true);
  return uuidToString(uuid);
}

//parse String to UUID binary
export function parseUUID(uuid) {
  const value = new Uint8Array(16);
  
  // Parse the first two characters (first byte)
  value[0] = ((BASE64_CHARS.indexOf(uuid[0]) << 6) | BASE64_CHARS.indexOf(uuid[1])) & 0xff;
  
  // Parse the remaining 20 characters in groups of 4 (3 bytes per group)
  let index = 1; // Start from the second byte
  for (let i = 2; i < 22; i += 4) {
    const triplet =
      (BASE64_CHARS.indexOf(uuid[i]) << 18) |
      (BASE64_CHARS.indexOf(uuid[i + 1]) << 12) |
      (BASE64_CHARS.indexOf(uuid[i + 2]) << 6) |
      BASE64_CHARS.indexOf(uuid[i + 3]);
      
    value[index++] = (triplet >> 16) & 0xff;
    value[index++] = (triplet >> 8) & 0xff;
    value[index++] = triplet & 0xff;
  }
  return value;
}
/**
 *
 * @param {Uint8Array} uuidBytes
 * @returns
 */
export function uuid2BigInt(uuidBytes) {
  // Ensure we have a proper 16-byte UUID
  if (uuidBytes.byteLength !== 16) {
    throw new Error("UUID must be exactly 16 bytes");
  }

  let result = window.BigInt(0);

  // Convert each byte to BigInt and shift appropriately
  for (let i = 0; i < 16; i++) {
    // Shift left 8 bits for each byte position and add the current byte
    result = (result << window.BigInt(8)) | window.BigInt(uuidBytes.at(i));
  }

  return result;
}

/**
 *
 * @param {BigInt} value
 */
export function bigInt2uuid(value) {
  const uuidBytes = new Uint8Array(16);

  for (let i = 15; i >= 0; i--) {
    uuidBytes[i] = Number(value & window.BigInt(0xff));
    value >>= window.BigInt(8);
  }
  return uuidBytes;
}
