/**
 * Utility functions for handling custom bullet characters in lists
 */

/**
 * Extract custom bullet character from list-style-type CSS property
 * @param {String} listStyleType - The value of list-style-type CSS property
 * @returns {String|null} - The extracted custom bullet character or null if not found
 */
function extractCustomBulletChar(listStyleType) {
  if (!listStyleType) {
    return null;
  }

  // Common named bullet styles mapping
  const bulletMap = {
    disc: '•',
    circle: '○',
    square: '■',
    none: ' ',
    // You can add more mappings as needed
  };

  // Check if list-style-type is a named value in our map
  if (bulletMap[listStyleType.toLowerCase()]) {
    return bulletMap[listStyleType.toLowerCase()];
  }

  // Check if list-style-type is wrapped in quotes (as a custom character)
  const quoteMatch = listStyleType.match(/^(['"'])(.*?)\1$/);
  if (quoteMatch) {
    return quoteMatch[2]; // Return the content between quotes
  }

  // If it's not a standard value and not in quotes, it might be directly specified
  // Only use it if it's a single character or an emoji (which can be multiple chars but render as one glyph)
  if (listStyleType.length === 1 || /\p{Emoji}/u.test(listStyleType)) {
    return listStyleType;
  }

  return null;
}

/**
 * Determine the best font to use for a bullet character
 * @param {String} bulletChar - The bullet character
 * @param {String} defaultFont - The font being used for the surrounding text
 * @returns {String} - The font to use for the bullet character
 */
function getBulletFont(bulletChar, defaultFont) {
  // Standard bullets that work best with Symbol font
  const symbolChars = ['•', '○', '■', '□', '◊', '▪', '▫', '◦', '\uF0B7', '\u25E6', '\u25AA'];

  // Check if this is a standard bullet that needs Symbol font
  if (symbolChars.includes(bulletChar)) {
    return 'Symbol';
  }

  // For emojis and other special characters that might not render well in standard fonts
  if (/\p{Emoji}/u.test(bulletChar)) {
    // Try to use a font known for good emoji support, but fall back to the document's font
    return 'Segoe UI Emoji';
  }

  // For most custom characters, using the text's font provides better visual consistency
  return defaultFont || 'Arial';
}

module.exports = {
  extractCustomBulletChar,
  getBulletFont,
};
