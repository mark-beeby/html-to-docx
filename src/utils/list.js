// Import custom bullet utilities
import { extractCustomBulletChar, getBulletFont } from './custom-bullets';

class ListStyleBuilder {
  // defaults is an object passed in from constants.js / numbering with the following properties:
  // defaultOrderedListStyleType: 'decimal' (unless otherwise specified)
  constructor(defaults) {
    this.defaults = defaults || { defaultOrderedListStyleType: 'decimal' };

    // Define standard bullet characters for different levels
    this.defaultBulletChars = [
      '\uF0B7', // Level 0 - Default bullet (•)
      '\uF0B7', // Level 1
      '\u25E6', // Level 2 - Circle (◦)
      '\u25AA', // Level 3 - Square (▪)
      '\u25AB', // Level 4
      '\u2713', // Level 5 - Checkmark (✓)
      '\u25B8', // Level 6 - Right-pointing triangle (▸)
      '\u2192', // Level 7 - Right arrow (→)
    ];

    // Define bullet character font mappings
    this.bulletFonts = {
      '\uF0B7': 'Symbol',
      '\u25E6': 'Symbol',
      '\u25AA': 'Symbol',
      '\u25AB': 'Symbol',
      '\u2713': 'Symbol',
      '\u25B8': 'Symbol',
      '\u2192': 'Symbol',
      // Default fonts for Unicode ranges
      emoji: 'Segoe UI Emoji',
      default: 'Symbol',
    };
  }

  /**
   * Get bullet character for unordered list
   * @param {Object} style - The style properties of the list
   * @param {Number} lvl - The level of the list
  }

  // eslint-disable-next-line class-methods-use-this
  getListStyleType(listType) {
    switch (listType) {
      case 'upper-roman':
        return 'upperRoman';
      case 'lower-roman':
        return 'lowerRoman';
      case 'upper-alpha':
      case 'upper-alpha-bracket-end':
        return 'upperLetter';
      case 'lower-alpha':
      case 'lower-alpha-bracket-end':
        return 'lowerLetter';
      case 'decimal':
      case 'decimal-bracket':
        return 'decimal';
      default:
        return this.defaults.defaultOrderedListStyleType;
    }
  }

  /**
   * Get the bullet character to use for unordered lists
   * @param {Object} properties - List properties including style and attributes
   * @param {Number} level - List nesting level
   * @returns {String} - The bullet character to use
   */
  getBulletChar(properties, level) {
    // Check for custom bullet in list-style-type or data-bullet-style
    let customBullet = null;

    // First check style attribute
    if (properties.style && properties.style['list-style-type']) {
      customBullet = extractCustomBulletChar(properties.style['list-style-type']);
    }

    // Then check data attribute if set by preprocessor
    if (!customBullet && properties.attributes && properties.attributes['data-bullet-style']) {
      customBullet = extractCustomBulletChar(properties.attributes['data-bullet-style']);
    }

    // Return custom bullet if found
    if (customBullet) {
      return customBullet;
    }

    // Fall back to default bullet for this level
    return this.defaultBulletChars[level] || this.defaultBulletChars[0];
  }

  /**
   * Get the appropriate font for a bullet character
   * @param {Object} properties - List properties including style and attributes
   * @param {String} bulletChar - The bullet character
   * @param {String} documentFont - The document's default font
   * @returns {String} - The font to use
   */
  // eslint-disable-next-line class-methods-use-this
  getBulletFont(properties, bulletChar, documentFont) {
    // First try to get font from data attribute if set by preprocessor
    let fontFamily = null;

    if (properties.attributes && properties.attributes['data-bullet-font']) {
      fontFamily = properties.attributes['data-bullet-font'];
    }

    // If no specific font found, use document default
    if (!fontFamily) {
      fontFamily = documentFont || 'Arial';
    }

    // Get appropriate font for this bullet character
    return getBulletFont(bulletChar, fontFamily);
  }

  getListPrefixSuffix(style, lvl) {
    let listType = this.defaults.defaultOrderedListStyleType;

    if (style && style['list-style-type']) {
      listType = style['list-style-type'];
    }

    switch (listType) {
      case 'upper-roman':
      case 'lower-roman':
      case 'upper-alpha':
      case 'lower-alpha':
        return `%${lvl + 1}.`;
      case 'upper-alpha-bracket-end':
      case 'lower-alpha-bracket-end':
      case 'decimal-bracket-end':
        return `%${lvl + 1})`;
      case 'decimal-bracket':
        return `(%${lvl + 1})`;
      case 'decimal':
      default:
        return `%${lvl + 1}.`;
    }
  }
}

export default ListStyleBuilder;
