import { create, fragment } from 'xmlbuilder2';
import { nanoid } from 'nanoid';
import https from 'https';
import http from 'http';
import sharp from 'sharp';
import * as fs from 'fs/promises';
import { createHash } from 'crypto';
import { convertToODTTF } from './utils/odttf';
import {
  generateCoreXML,
  generateStylesXML,
  generateNumberingXMLTemplate,
  generateThemeXML,
  documentRelsXML as documentRelsXMLString,
  settingsXML as settingsXMLString,
  webSettingsXML as webSettingsXMLString,
  contentTypesXML as contentTypesXMLString,
  fontTableXML as fontTableXMLString,
  genericRelsXML as genericRelsXMLString,
  generateDocumentTemplate,
} from './schemas';
import { convertVTreeToXML } from './helpers';
import namespaces from './namespaces';
import {
  footerType as footerFileType,
  headerType as headerFileType,
  themeType as themeFileType,
  landscapeMargins,
  portraitMargins,
  defaultOrientation,
  landscapeWidth,
  landscapeHeight,
  applicationName,
  defaultFont,
  defaultFontSize,
  defaultLang,
  hyperlinkType,
  documentFileName,
  imageType,
  fontType,
  defaultDocumentOptions,
} from './constants';
import ListStyleBuilder from './utils/list';
import { fontFamilyToTableObject } from './utils/font-family-conversion';
import { pixelRegex, pixelToHIP, pointRegex, pointToHIP } from './utils/unit-conversion';
import * as xmlBuilder from './helpers/xml-builder';

function generateContentTypesFragments(contentTypesXML, type, objects) {
  if (objects && Array.isArray(objects)) {
    objects.forEach((object) => {
      const contentTypesFragment = fragment({ defaultNamespace: { ele: namespaces.contentTypes } })
        .ele('Override')
        .att('PartName', `/word/${type}${object[`${type}Id`]}.xml`)
        .att(
          'ContentType',
          `application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml`
        )
        .up();

      contentTypesXML.root().import(contentTypesFragment);
    });
  }
}

function generateSectionReferenceXML(documentXML, documentSectionType, objects, isEnabled) {
  if (isEnabled && objects && Array.isArray(objects) && objects.length) {
    const xmlFragment = fragment();
    objects.forEach(({ relationshipId, type }) => {
      const objectFragment = fragment({ namespaceAlias: { w: namespaces.w, r: namespaces.r } })
        .ele('@w', `${documentSectionType}Reference`)
        .att('@r', 'id', `rId${relationshipId}`)
        .att('@w', 'type', type)
        .up();
      xmlFragment.import(objectFragment);
    });

    documentXML.root().first().first().import(xmlFragment);
  }
}

function generateXMLString(xmlString) {
  const xmlDocumentString = create({ encoding: 'UTF-8', standalone: true }, xmlString);
  return xmlDocumentString.toString({ prettyPrint: true });
}

async function generateSectionXML(vTree, type = 'header') {
  const sectionXML = create({
    encoding: 'UTF-8',
    standalone: true,
    namespaceAlias: {
      w: namespaces.w,
      ve: namespaces.ve,
      o: namespaces.o,
      r: namespaces.r,
      v: namespaces.v,
      wp: namespaces.wp,
      w10: namespaces.w10,
    },
  }).ele('@w', type === 'header' ? 'hdr' : 'ftr');

  const XMLFragment = fragment();
  await convertVTreeToXML(this, vTree, XMLFragment);
  if (type === 'footer' && XMLFragment.first().node.tagName === 'p' && this.pageNumber) {
    XMLFragment.first().import(
      fragment({ namespaceAlias: { w: namespaces.w } })
        .ele('@w', 'fldSimple')
        .att('@w', 'instr', 'PAGE')
        .ele('@w', 'r')
        .up()
        .up()
    );
  }
  sectionXML.root().import(XMLFragment);

  const referenceName = type === 'header' ? 'Header' : 'Footer';
  this[`last${referenceName}Id`] += 1;

  return { [`${type}Id`]: this[`last${referenceName}Id`], [`${type}XML`]: sectionXML };
}

class DocxDocument {
  constructor(properties) {
    this.zip = properties.zip;
    this.htmlString = properties.htmlString;
    this.orientation = properties.orientation;
    this.pageSize = properties.pageSize || defaultDocumentOptions.pageSize;
    const spacing = {
      above: properties.spacing.defaultParagraphSpacing?.above ?? 0.17,
      below: properties.spacing.defaultParagraphSpacing?.below ?? 0.17,
      line: properties.spacing.defaultLineSpacing ?? 12,
    };
    const spacingVals = this.convertToDocxSpacing(spacing.above, spacing.below, spacing.line);
    this.defaultSpacing = {
      paragraphSpacing: {
        above: spacingVals.above,
        below: spacingVals.below,
      },
      lineSpacing: spacingVals.line,
    };

    const isPortraitOrientation = this.orientation === defaultOrientation;
    const height = this.pageSize.height ? this.pageSize.height : landscapeHeight;
    const width = this.pageSize.width ? this.pageSize.width : landscapeWidth;

    this.width = isPortraitOrientation ? width : height;
    this.height = isPortraitOrientation ? height : width;

    const marginsObject = properties.margins;
    this.margins =
      // eslint-disable-next-line no-nested-ternary
      marginsObject && Object.keys(marginsObject).length
        ? marginsObject
        : isPortraitOrientation
        ? portraitMargins
        : landscapeMargins;

    this.availableDocumentSpace = this.width - this.margins.left - this.margins.right;
    this.title = properties.title || '';
    this.subject = properties.subject || '';
    this.creator = properties.creator || applicationName;
    this.keywords = properties.keywords || [applicationName];
    this.description = properties.description || '';
    this.lastModifiedBy = properties.lastModifiedBy || applicationName;
    this.revision = properties.revision || 1;
    this.createdAt = properties.createdAt || new Date();
    this.modifiedAt = properties.modifiedAt || new Date();
    this.headerType = properties.headerType || 'default';
    this.header = properties.header || false;
    this.footerType = properties.footerType || 'default';
    this.footer = properties.footer || false;
    this.suppressFooterMargins = properties.suppressFooterMargins || false;
    this.font = properties.font || defaultFont;
    this.fontSize = properties.fontSize || defaultFontSize;
    this.complexScriptFontSize = properties.complexScriptFontSize || defaultFontSize;
    this.lang = properties.lang || defaultLang;
    this.tableRowCantSplit =
      (properties.table && properties.table.row && properties.table.row.cantSplit) || false;
    this.pageNumber = properties.pageNumber || false;
    this.skipFirstHeaderFooter = properties.skipFirstHeaderFooter || false;
    this.lineNumber = properties.lineNumber ? properties.lineNumberOptions : null;

    this.lastNumberingId = 0;
    this.lastMediaId = 0;
    this.lastHeaderId = 0;
    this.lastFooterId = 0;
    this.defaultLineHeight = properties.defaultLineHeight ? properties.defaultLineHeight : 1.5;
    this.stylesObjects = [];
    this.numberingObjects = [];
    this.fontTableObjects = [];
    this.fonts = properties.fonts || [];
    this.relationshipFilename = documentFileName;
    this.relationships = [{ fileName: documentFileName, lastRelsId: 5, rels: [] }];
    this.mediaFiles = [];
    this.headerObjects = {};
    // if you define multiple sections with the same header you get duplication, this tracks when one has been added
    this.headerAdded = false;
    this.footerObjects = {};
    // if you define multiple sections with the same footer you get duplication, this tracks when one has been added
    this.footerAdded = false;

    // Support for multiple sections from data-section divs
    this.sections = [];
    this.currentSectionId = 0;
    this.documentXML = null;

    this.headerVariants = new Map(); // Track all header combinations
    this.backgroundImageCache = new Map(); // Deduplicate background images
    this.headerDeduplicationMap = new Map(); // Deduplicate identical headers

    this.generateContentTypesXML = this.generateContentTypesXML.bind(this);
    this.generateDocumentXML = this.generateDocumentXML.bind(this);
    this.generateCoreXML = this.generateCoreXML.bind(this);
    this.generateSettingsXML = this.generateSettingsXML.bind(this);
    this.generateWebSettingsXML = this.generateWebSettingsXML.bind(this);
    this.generateStylesXML = this.generateStylesXML.bind(this);
    this.generateFontTableXML = this.generateFontTableXML.bind(this);
    this.generateThemeXML = this.generateThemeXML.bind(this);
    this.generateNumberingXML = this.generateNumberingXML.bind(this);
    this.generateRelsXML = this.generateRelsXML.bind(this);
    this.createMediaFile = this.createMediaFile.bind(this);
    this.createDocumentRelationships = this.createDocumentRelationships.bind(this);
    this.generateHeaderXML = this.generateHeaderXML.bind(this);
    this.generateFooterXML = this.generateFooterXML.bind(this);
    this.generateSectionXML = generateSectionXML.bind(this);
    this.embedFonts = this.embedFonts.bind(this);
    this.ListStyleBuilder = new ListStyleBuilder({
      defaultOrderedListStyleType: 'decimal',
    });
  }

  calculateHeaderHeight() {
    let headerHeight = 720; // Base header height (720 TWIPs = 0.5 inches)
    if (this.backgroundImageHeight) {
      headerHeight = Math.max(headerHeight, this.backgroundImageHeight);
    }

    if (this.logoHeights && this.logoHeights.length > 0) {
      const maxLogoHeight = Math.max(...this.logoHeights);
      headerHeight = Math.max(headerHeight, maxLogoHeight);
    }

    if (this.vTreeHeight) {
      headerHeight = Math.max(720 + this.vTreeHeight, headerHeight);
    }

    headerHeight += 240;

    return headerHeight;
  }

  /**
   * Converts spacing measurements to DOCX XML values
   * @param aboveLines
   * @param belowLines
   * @param {number} lineSpacingPt - Line spacing in points
   * @param baseLineHeightPt
   * @returns {Object} DOCX spacing values
   */
  // eslint-disable-next-line class-methods-use-this
  convertToDocxSpacing(aboveLines, belowLines, lineSpacingPt = null, baseLineHeightPt = 15.6) {
    // Convert paragraph spacing from line units to points
    const abovePt = aboveLines * baseLineHeightPt;
    const belowPt = belowLines * baseLineHeightPt;

    // Convert to twentieths of a point
    const above = Math.round(abovePt * 20);
    const below = Math.round(belowPt * 20);

    // Calculate line spacing in twentieths (if provided)
    const line = typeof lineSpacingPt !== 'undefined' ? Math.round(lineSpacingPt * 20) : null;

    // Return object with the correct values
    return {
      above,
      below,
      line,
    };
  }

  calculateFooterHeight() {
    let footerHeight = 720; // Base header height (720 TWIPs = 0.5 inches)
    if (this.footerImageHeight) {
      footerHeight = Math.max(footerHeight, this.backgroundImageHeight);
    }

    if (this.logoHeights && this.logoHeights.length > 0) {
      const maxLogoHeight = Math.max(...this.logoHeights);
      footerHeight = Math.max(footerHeight, maxLogoHeight);
    }

    if (this.vTreeHeight) {
      footerHeight = Math.max(720 + this.vTreeHeight, footerHeight);
    }

    footerHeight += 240;

    return footerHeight;
  }

  generateContentTypesXML() {
    const contentTypesXML = create({ encoding: 'UTF-8', standalone: true }, contentTypesXMLString);

    generateContentTypesFragments(contentTypesXML, 'header', this.headerObjects);
    generateContentTypesFragments(contentTypesXML, 'footer', this.footerObjects);

    // Add font content type
    if (this.fonts && this.fonts.length > 0) {
      const contentTypesFragment = fragment({
        defaultNamespace: { ele: namespaces.contentTypes },
      })
        .ele('Default')
        .att('Extension', 'odttf')
        .att('ContentType', 'application/vnd.openxmlformats-officedocument.obfuscatedFont')
        .up();

      contentTypesXML.root().import(contentTypesFragment);
    }

    return contentTypesXML.toString({ prettyPrint: true });
  }

  generateDocumentXML() {
    const documentXML = create(
      { encoding: 'UTF-8', standalone: true },
      generateDocumentTemplate(this.width, this.height, this.orientation, this.margins)
    );

    if (this.suppressFooterMargins) {
      this.margins.footer = 1;
      this.margins.bottom = 1;
    }

    documentXML.root().att('xmlns:w', namespaces.w);
    documentXML.root().att('xmlns:r', namespaces.r);
    documentXML.root().att('xmlns:wp', namespaces.wp);
    documentXML.root().att('xmlns:a', namespaces.a);
    documentXML.root().att('xmlns:pic', namespaces.pic);
    documentXML.root().att('xmlns:ve', namespaces.ve);
    documentXML.root().att('xmlns:o', namespaces.o);
    documentXML.root().att('xmlns:v', namespaces.v);
    documentXML.root().att('xmlns:w10', namespaces.w10);
    documentXML.root().first().import(this.documentXML);

    generateSectionReferenceXML(documentXML, 'header', this.headerObjects, this.header);
    generateSectionReferenceXML(documentXML, 'footer', this.footerObjects, this.footer);

    if ((this.header || this.footer) && this.skipFirstHeaderFooter) {
      documentXML.root().first().ele('@w', 'titlePg').up();
    }

    if (this.lineNumber) {
      const { countBy, start, restart } = this.lineNumber;
      documentXML
        .root()
        .first()
        .ele('@w', 'lnNumType')
        .att('@w', 'countBy', countBy)
        .att('@w', 'start', start)
        .att('@w', 'restart', restart)
        .up();
    }

    // Check if we have custom sections from data-section divs
    const hasCustomSections = this.sections && this.sections.length > 0;

    // Modify section properties for responsive header
    if (this.sectionProperties) {
      const body = documentXML.root().first();

      if (hasCustomSections) {
        console.log(
          `Document has ${this.sections.length} custom sections - skipping document-level section properties`
        );

        // Remove any existing document-level sectPr to avoid duplicate section properties
        body.each((child) => {
          if (child.node.nodeName === 'w:sectPr') {
            child.remove();
          }
        });

        // Skip all remaining section property work
        return documentXML.toString({ prettyPrint: true });
      }

      // Standard case for documents without custom sections
      // Find existing sectPr or create a new one
      let sectPr;
      // eslint-disable-next-line consistent-return
      body.each((child) => {
        if (child.node.nodeName === 'w:sectPr') {
          sectPr = child;
          return false; // Stop iteration
        }
      });

      if (!sectPr) {
        sectPr = body.ele('w:sectPr');
      }

      // Find or create pgMar in sectPr
      let pgMar;
      sectPr.each((child) => {
        if (child.node.nodeName === 'w:pgMar') {
          pgMar = child;
          return false; // Stop iteration
        }
        return true; // Continue iteration
      });

      if (!pgMar) {
        pgMar = sectPr.ele('w:pgMar');
      }

      // Ensure other necessary elements from this.sectionProperties are present
      if (typeof this.sectionProperties.each === 'function') {
        this.sectionProperties.each((child) => {
          if (child.node.nodeName !== 'w:pgMar') {
            let existingChild;
            sectPr.each((sectionChild) => {
              if (sectionChild.node.nodeName === child.node.nodeName) {
                existingChild = sectionChild;
                return false; // Stop iteration
              }
              return true;
            });

            if (!existingChild) {
              sectPr.import(child);
            }
          }
        });
      }
    }

    // Use the updated section properties
    if (this.sectionProperties) {
      documentXML.root().first().import(this.sectionProperties);
    }

    return documentXML.toString({ prettyPrint: true });
  }

  generateCoreXML() {
    return generateXMLString(
      generateCoreXML(
        this.title,
        this.subject,
        this.creator,
        this.keywords,
        this.description,
        this.lastModifiedBy,
        this.revision,
        this.createdAt,
        this.modifiedAt
      )
    );
  }

  // eslint-disable-next-line class-methods-use-this
  generateSettingsXML() {
    return generateXMLString(settingsXMLString(!!this.fonts?.length));
  }

  // eslint-disable-next-line class-methods-use-this
  generateWebSettingsXML() {
    return generateXMLString(webSettingsXMLString);
  }

  generateStylesXML() {
    return generateXMLString(
      generateStylesXML(
        this.font,
        this.fontSize,
        this.complexScriptFontSize,
        this.lang,
        this.defaultSpacing
      )
    );
  }

  generateFontTableXML() {
    const fontTableXML = create(
      { encoding: 'UTF-8', standalone: true },
      fontTableXMLString(this.fonts)
    );
    return fontTableXML.toString({ prettyPrint: true });
  }

  generateThemeXML() {
    return generateXMLString(generateThemeXML(this.font));
  }

  generateNumberingXML() {
    const numberingXML = create(
      { encoding: 'UTF-8', standalone: true },
      generateNumberingXMLTemplate()
    );

    const abstractNumberingFragments = fragment();
    const numberingFragments = fragment();
    this.numberingObjects.forEach(({ numberingId, type, properties }) => {
      const abstractNumberingFragment = fragment({ namespaceAlias: { w: namespaces.w } })
        .ele('@w', 'abstractNum')
        .att('@w', 'abstractNumId', String(numberingId));

      [...Array(8).keys()].forEach((level) => {
        const levelFragment = fragment({ namespaceAlias: { w: namespaces.w } })
          .ele('@w', 'lvl')
          .att('@w', 'ilvl', level)
          .ele('@w', 'start')
          .att(
            '@w',
            'val',
            type === 'ol'
              ? (properties.attributes && properties.attributes['data-start']) || 1
              : '1'
          )
          .up()
          .ele('@w', 'numFmt')
          .att(
            '@w',
            'val',
            type === 'ol'
              ? this.ListStyleBuilder.getListStyleType(
                  properties.style && properties.style['list-style-type']
                )
              : 'bullet'
          )
          .up();

        // Get the appropriate bullet character or list prefix/suffix
        if (type === 'ol') {
          // For ordered lists, use the existing logic
          levelFragment
            .ele('@w', 'lvlText')
            .att('@w', 'val', this.ListStyleBuilder.getListPrefixSuffix(properties.style, level))
            .up();
        } else {
          // For unordered lists, get the custom bullet character
          const bulletChar = this.ListStyleBuilder.getBulletChar(properties, level);
          levelFragment.ele('@w', 'lvlText').att('@w', 'val', bulletChar).up();
        }

        levelFragment
          .ele('@w', 'lvlJc')
          .att('@w', 'val', 'left')
          .up()
          .ele('@w', 'pPr')
          .ele('@w', 'tabs')
          .ele('@w', 'tab')
          .att('@w', 'val', 'num')
          .att('@w', 'pos', (level + 1) * 480)
          .up()
          .up()
          .ele('@w', 'ind')
          .att('@w', 'left', (level + 1) * 480)
          .att('@w', 'hanging', 260)
          .up()
          .up()
          .up();

        // Create rPr fragment for both ul and ol
        const rPrFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'rPr');

        // Add font settings based on list type
        if (type === 'ul') {
          // For unordered lists, determine the appropriate font for the bullet character
          const bulletChar = this.ListStyleBuilder.getBulletChar(properties, level);
          const documentDefaultFont = this.font || defaultFont || 'Arial';
          const bulletFont = this.ListStyleBuilder.getBulletFont(
            properties,
            bulletChar,
            documentDefaultFont
          );

          rPrFragment
            .ele('@w', 'rFonts')
            .att('@w', 'ascii', bulletFont)
            .att('@w', 'hAnsi', bulletFont)
            .att('@w', 'hint', 'default')
            .up();

          // The Symbol character needs special handling with the ascii char attribute
          if (bulletFont === 'Symbol') {
            rPrFragment.ele('@w', 'ascii').att('@w', 'char', '•').up();
          }
        }
        // Add color if specified in properties (for both ul and ol)
        if (properties.style && properties.style.primaryColour) {
          rPrFragment
            .ele('@w', 'color')
            .att('@w', 'val', properties.style.primaryColour.replace('#', ''));
        }

        // Add color if specified in properties (for both ul and ol)
        if (properties.style && properties.style['font-size']) {
          const fontSizeString = properties.style['font-size'];
          let fontSize = properties.style['font-size'];
          if (pointRegex.test(fontSizeString)) {
            const matchedParts = fontSizeString.match(pointRegex);
            // convert point to half point
            fontSize = pointToHIP(matchedParts[1]);
          } else if (pixelRegex.test(fontSizeString)) {
            const matchedParts = fontSizeString.match(pixelRegex);
            // convert pixels to half point
            fontSize = pixelToHIP(matchedParts[1]);
          }

          rPrFragment.ele('@w', 'sz').att('@w', 'val', fontSize);
        }

        levelFragment.last().import(rPrFragment);
        abstractNumberingFragment.import(levelFragment);
      });
      abstractNumberingFragment.up();
      abstractNumberingFragments.import(abstractNumberingFragment);

      numberingFragments.import(
        fragment({ namespaceAlias: { w: namespaces.w } })
          .ele('@w', 'num')
          .att('@w', 'numId', String(numberingId))
          .ele('@w', 'abstractNumId')
          .att('@w', 'val', String(numberingId))
          .up()
          .up()
      );
    });

    numberingXML.root().import(abstractNumberingFragments);
    numberingXML.root().import(numberingFragments);

    return numberingXML.toString({ prettyPrint: true });
  }

  // eslint-disable-next-line class-methods-use-this
  appendRelationships(xmlFragment, relationships) {
    relationships.forEach(({ relationshipId, type, target, targetMode }) => {
      xmlFragment.import(
        fragment({ defaultNamespace: { ele: namespaces.relationship } })
          .ele('Relationship')
          .att('Id', `rId${relationshipId}`)
          .att('Type', type)
          .att('Target', target)
          .att('TargetMode', targetMode)
          .up()
      );
    });
  }

  generateRelsXML() {
    const relationshipXMLStrings = this.relationships.map(({ fileName, rels }) => {
      const xmlFragment = create(
        { encoding: 'UTF-8', standalone: true },
        fileName === documentFileName ? documentRelsXMLString : genericRelsXMLString
      );
      this.appendRelationships(xmlFragment.root(), rels);

      return { fileName, xmlString: xmlFragment.toString({ prettyPrint: true }) };
    });

    return relationshipXMLStrings;
  }

  createNumbering(type, properties) {
    this.lastNumberingId += 1;
    // Ensure style object exists and contains colour if specified
    const style = properties?.style || {};
    if (properties?.primaryColour && !style.primaryColour) {
      style.primaryColour = properties.primaryColour;
    }

    this.numberingObjects.push({
      numberingId: this.lastNumberingId,
      type,
      properties: {
        ...properties,
        style,
      },
    });

    return this.lastNumberingId;
  }

  createFont(fontFamily) {
    const fontTableObject = fontFamilyToTableObject(fontFamily, this.font);
    this.fontTableObjects.push(fontTableObject);
    return fontTableObject.fontName;
  }

  createMediaFile(base64String) {
    // eslint-disable-next-line no-useless-escape
    const matches = base64String.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
    if (matches.length !== 3) {
      throw new Error('Invalid base64 string');
    }

    const base64FileContent = matches[2];
    // matches array contains file type in base64 format - image/jpeg and base64 stringified data
    const fileExtension =
      matches[1].match(/\/(.*?)$/)[1] === 'octet-stream' ? 'png' : matches[1].match(/\/(.*?)$/)[1];

    const fileNameWithExtension = `image-${nanoid()}.${fileExtension}`;

    this.lastMediaId += 1;

    return { id: this.lastMediaId, fileContent: base64FileContent, fileNameWithExtension };
  }

  createDocumentRelationships(fileName = 'document', type, target, targetMode = 'External', opts) {
    const ridOverride = opts?.ridOverride;
    let relationshipObject = this.relationships.find(
      (relationship) => relationship.fileName === fileName
    );
    let lastRelsId = 1;
    if (!ridOverride || !relationshipObject) {
      if (relationshipObject) {
        lastRelsId = relationshipObject.lastRelsId + 1;
        relationshipObject.lastRelsId = lastRelsId;
      } else {
        relationshipObject = { fileName, lastRelsId, rels: [] };
        this.relationships.push(relationshipObject);
      }
    }
    if (ridOverride) {
      lastRelsId = ridOverride;
    }

    let relationshipType;
    switch (type) {
      case hyperlinkType:
        relationshipType = namespaces.hyperlinks;
        break;
      case imageType:
        relationshipType = namespaces.images;
        break;
      case headerFileType:
        relationshipType = namespaces.headers;
        break;
      case footerFileType:
        relationshipType = namespaces.footers;
        break;
      case themeFileType:
        relationshipType = namespaces.themes;
        break;
      case fontType:
        relationshipType = namespaces.fonts;
        break;
    }

    relationshipObject.rels.push({
      relationshipId: lastRelsId,
      type: relationshipType,
      target,
      targetMode,
    });

    return lastRelsId;
  }

  generateEmptyHeaderXML() {
    const headerId = this.lastHeaderId + 1;
    const headerTypeName = 'none';
    this.lastHeaderId = headerId;
    const XMLFragment = create({
      encoding: 'UTF-8',
      standalone: true,
      namespaceAlias: {
        w: namespaces.w,
        r: namespaces.r,
        wp: namespaces.wp,
        a: namespaces.a,
        pic: namespaces.pic,
        ve: namespaces.ve,
        o: namespaces.o,
        v: namespaces.v,
        w10: namespaces.w10,
      },
    }).ele('@w', 'hdr');

    XMLFragment.import(
      fragment({ namespaceAlias: { w: namespaces.w } })
        .ele('@w', 'p')
        .ele('@w', 'pPr')
        .ele('@w', 'pStyle')
        .att('@w', 'val', 'header')
        .up()
        .up()
        .up()
    );

    // Store the header XML by type name to support multiple sections
    if (!this.headerXMLs) {
      this.headerXMLs = {};
    }

    // Store by type name
    this.headerXMLs[headerTypeName] = generateXMLString(
      XMLFragment.toString({ prettyPrint: true }),
      `word/header${headerId}.xml`
    );

    // Store header object by type name for sectPr references
    this.headerObjects[headerTypeName] = {
      headerId: `rId${headerId}`,
      height: 0,
    };

    // Return all the header information including type name
    return {
      headerId: `rId${headerId}`,
      headerXML: XMLFragment,
      headerHeight: 0,
      typeName: headerTypeName,
    };
  }

  generateEmptyFooterXML() {
    const footerId = this.lastFooterId + 1;
    const footerTypeName = 'none';
    this.lastFooterId = footerId;
    const XMLFragment = create({
      encoding: 'UTF-8',
      standalone: true,
      namespaceAlias: {
        w: namespaces.w,
        r: namespaces.r,
        wp: namespaces.wp,
        a: namespaces.a,
        pic: namespaces.pic,
        ve: namespaces.ve,
        o: namespaces.o,
        v: namespaces.v,
        w10: namespaces.w10,
      },
    }).ele('@w', 'ftr');

    const emptyFragment = fragment({ namespaceAlias: { w: namespaces.w } })
      .ele('@w', 'p')
      .ele('@w', 'pPr')
      .ele('@w', 'pStyle')
      .att('@w', 'val', 'footer')
      .up()
      .up()
      .up();
    XMLFragment.import(emptyFragment);

    // Return all the footer information including type name
    return {
      footerId: `rId${footerId}`,
      footerXML: XMLFragment,
      footerHeight: 0,
      typeName: footerTypeName,
    };
  }

  static calculateBackgroundDimensions(
    imageWidth,
    imageHeight,
    pageWidth,
    pageHeight,
    backgroundSize,
    backgroundPosition
  ) {
    let finalWidth = imageWidth * 9525; // Convert to EMUs
    let finalHeight = imageHeight * 9525;
    let containScale;
    let coverScale;

    // Handle background-size (same as before)
    switch (backgroundSize) {
      case 'stretch':
        finalWidth = pageWidth;
        finalHeight = pageHeight;
        break;
      case 'fit':
      case 'contain':
        containScale = Math.min(pageWidth / finalWidth, pageHeight / finalHeight);
        finalWidth *= containScale;
        finalHeight *= containScale;
        break;
      case 'cover':
        coverScale = Math.max(pageWidth / finalWidth, pageHeight / finalHeight);
        finalWidth *= coverScale;
        finalHeight *= coverScale;
        break;
      case 'original':
      default:
        break;
    }

    // Explicit position mapping
    const position = backgroundPosition || 'center';
    let posX = 0;
    let posY = 0;

    // Map all your specific positions
    switch (position) {
      case 'top-left':
        posX = 0;
        posY = 0;
        break;
      case 'top-center':
        posX = (pageWidth - finalWidth) / 2;
        posY = 0;
        break;
      case 'top-right':
        posX = pageWidth - finalWidth;
        posY = 0;
        break;
      case 'middle-left':
        posX = 0;
        posY = (pageHeight - finalHeight) / 2;
        break;
      case 'center':
        posX = (pageWidth - finalWidth) / 2;
        posY = (pageHeight - finalHeight) / 2;
        break;
      case 'middle-right':
        posX = pageWidth - finalWidth;
        posY = (pageHeight - finalHeight) / 2;
        break;
      case 'bottom-left':
        posX = 0;
        posY = pageHeight - finalHeight;
        break;
      case 'bottom-center':
        posX = (pageWidth - finalWidth) / 2;
        posY = pageHeight - finalHeight;
        break;
      case 'bottom-right':
        posX = pageWidth - finalWidth;
        posY = pageHeight - finalHeight;
        break;
      default:
        // Default to center
        posX = (pageWidth - finalWidth) / 2;
        posY = (pageHeight - finalHeight) / 2;
        break;
    }

    // Ensure positions don't go negative
    posX = Math.max(0, posX);
    posY = Math.max(0, posY);

    return { finalWidth, finalHeight, posX, posY };
  }

  // Add background to header
  async addBackgroundToHeader(headerXML, backgroundInfo, headerId) {
    // eslint-disable-next-line no-unused-vars
    const { backgroundUrl, backgroundSize, backgroundPosition, backgroundRepeat } = backgroundInfo;

    // Check if we've already processed this image
    let imageRelationshipId;
    let imageFile;

    if (this.backgroundImageCache.has(backgroundUrl)) {
      const cachedImage = this.backgroundImageCache.get(backgroundUrl);
      imageFile = cachedImage.imageFile;

      // Create a new relationship for this header, but reuse the image file
      imageRelationshipId = this.createDocumentRelationships(
        `header${headerId}`,
        imageType,
        `media/${imageFile.fileNameWithExtension}`,
        'Internal'
      );
    } else {
      // Process new image
      const { base64String, imageWidth, imageHeight } = await this.fetchImageAndGetDimensions(
        backgroundUrl
      );
      imageFile = this.createMediaFile(base64String);

      // Cache the image
      this.backgroundImageCache.set(backgroundUrl, {
        imageFile,
        imageWidth,
        imageHeight,
        base64String,
      });

      imageRelationshipId = this.createDocumentRelationships(
        `header${headerId}`,
        imageType,
        `media/${imageFile.fileNameWithExtension}`,
        'Internal'
      );
      // Add to zip only once
      this.zip
        .folder('word/media')
        .file(imageFile.fileNameWithExtension, imageFile.fileContent, { base64: true });
    }

    const cachedImage = this.backgroundImageCache.get(backgroundUrl);
    const { imageWidth, imageHeight } = cachedImage;

    // Calculate positioning and sizing
    const pageWidthEMU = this.width * 635;
    const pageHeightEMU = this.height * 635;

    // eslint-disable-next-line no-unused-vars
    const { finalWidth, finalHeight, posX, posY } = DocxDocument.calculateBackgroundDimensions(
      imageWidth,
      imageHeight,
      pageWidthEMU,
      pageHeightEMU,
      backgroundSize,
      backgroundPosition
    );

    // Add background to header XML (simplified structure)
    const bgParagraph = headerXML.ele('@w', 'p');

    // Add paragraph properties
    const pPr = bgParagraph.ele('@w', 'pPr');
    const spacing = pPr.ele('@w', 'spacing');
    spacing.att('@w', 'before', '0');
    spacing.att('@w', 'after', '0');
    spacing.att('@w', 'line', '0');
    spacing.att('@w', 'lineRule', 'auto');

    // Add the run with drawing
    const run = bgParagraph.ele('@w', 'r');
    const drawing = run.ele('@w', 'drawing');
    const anchor = drawing.ele('@wp', 'anchor');

    // Set anchor attributes
    anchor.att('behindDoc', '1');
    anchor.att('distT', '0');
    anchor.att('distB', '0');
    anchor.att('distL', '0');
    anchor.att('distR', '0');
    anchor.att('simplePos', '0');
    anchor.att('relativeHeight', '0');
    anchor.att('locked', '1');
    anchor.att('layoutInCell', '0');
    anchor.att('allowOverlap', '1');

    // Add positioning
    const simplePos = anchor.ele('@wp', 'simplePos');
    simplePos.att('x', '0');
    simplePos.att('y', '0');

    const positionH = anchor.ele('@wp', 'positionH');
    positionH.att('relativeFrom', 'page');
    const posOffsetH = positionH.ele('@wp', 'posOffset');
    posOffsetH.txt(posX.toString());

    const positionV = anchor.ele('@wp', 'positionV');
    positionV.att('relativeFrom', 'page'); // Make sure this is 'page' not 'paragraph' or 'margin'
    const posOffsetV = positionV.ele('@wp', 'posOffset');
    posOffsetV.txt(posY.toString());

    // Continue with extent, docPr, graphic, etc. (similar to previous implementation)
    const extent = anchor.ele('@wp', 'extent');
    extent.att('cx', finalWidth.toString());
    extent.att('cy', finalHeight.toString());

    const effectExtent = anchor.ele('@wp', 'effectExtent');
    effectExtent.att('l', '0');
    effectExtent.att('t', '0');
    effectExtent.att('r', '0');
    effectExtent.att('b', '0');

    anchor.ele('@wp', 'wrapNone');

    const docPr = anchor.ele('@wp', 'docPr');
    docPr.att('id', imageFile.id);
    docPr.att('name', 'Page Background');

    // Add the rest of the graphic structure...
    const cNvGraphicFramePr = anchor.ele('@wp', 'cNvGraphicFramePr');
    const graphicFrameLocks = cNvGraphicFramePr.ele('@a', 'graphicFrameLocks');
    graphicFrameLocks.att('noChangeAspect', '1');

    const graphic = anchor.ele('@a', 'graphic');
    const graphicData = graphic.ele('@a', 'graphicData');
    graphicData.att('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture');

    const pic = graphicData.ele('@pic', 'pic');
    const nvPicPr = pic.ele('@pic', 'nvPicPr');
    const cNvPr = nvPicPr.ele('@pic', 'cNvPr');
    cNvPr.att('id', '0');
    cNvPr.att('name', 'Page Background');
    nvPicPr.ele('@pic', 'cNvPicPr');

    const blipFill = pic.ele('@pic', 'blipFill');
    const blip = blipFill.ele('@a', 'blip');
    blip.att('@r', 'embed', `rId${imageRelationshipId}`);

    const stretch = blipFill.ele('@a', 'stretch');
    stretch.ele('@a', 'fillRect');

    const spPr = pic.ele('@pic', 'spPr');
    const xfrm = spPr.ele('@a', 'xfrm');
    const off = xfrm.ele('@a', 'off');
    off.att('x', '0');
    off.att('y', '0');

    const ext = xfrm.ele('@a', 'ext');
    ext.att('cx', finalWidth.toString());
    ext.att('cy', finalHeight.toString());

    const prstGeom = spPr.ele('@a', 'prstGeom');
    prstGeom.att('prst', 'rect');
    prstGeom.ele('@a', 'avLst');

    return imageRelationshipId;
  }

  // Generate header with background only - no complex content processing
  async generateSectionHeader(sectionInfo) {
    const { headerType, backgroundUrl, backgroundSize, backgroundPosition, backgroundRepeat } =
      sectionInfo;

    // Create variant name based on background
    const variantName = `${headerType}_bg_${this.hashBackgroundInfo(sectionInfo)}`;

    // Check if we already have this variant
    if (this.headerVariants.has(variantName)) {
      return this.headerVariants.get(variantName);
    }

    const headerId = this.lastHeaderId + 1;
    this.lastHeaderId = headerId;

    const headerXML = create({
      encoding: 'UTF-8',
      standalone: true,
      namespaceAlias: {
        w: namespaces.w,
        r: namespaces.r,
        wp: namespaces.wp,
        a: namespaces.a,
        pic: namespaces.pic,
        ve: namespaces.ve,
        o: namespaces.o,
        v: namespaces.v,
        w10: namespaces.w10,
      },
    }).ele('@w', 'hdr');

    // Add background
    await this.addBackgroundToHeader(
      headerXML,
      {
        backgroundUrl,
        backgroundSize,
        backgroundPosition,
        backgroundRepeat,
      },
      headerId
    );

    // Add a simple header paragraph (no complex content processing)
    const headerParagraph = headerXML.ele('@w', 'p');
    const pPr = headerParagraph.ele('@w', 'pPr');
    const pStyle = pPr.ele('@w', 'pStyle');
    pStyle.att('@w', 'val', 'Header');

    const result = {
      headerId: `${headerId}`,
      headerXML,
      headerHeight: 720, // Simple fixed height for background headers
      variantName,
    };

    // Cache the result
    this.headerVariants.set(variantName, result);

    return result;
  }

  // Create a short hash for background info
  // eslint-disable-next-line class-methods-use-this
  hashBackgroundInfo(backgroundInfo) {
    const info = `${backgroundInfo.backgroundUrl || ''}|${backgroundInfo.backgroundSize || ''}|${
      backgroundInfo.backgroundPosition || ''
    }|${backgroundInfo.backgroundRepeat || ''}`;
    return createHash('md5').update(info).digest('hex').substring(0, 8);
  }

  async generateHeaderXML(vTree, headerConfig, headerTypeName = 'default') {
    const headerId = this.lastHeaderId + 1;
    this.lastHeaderId = headerId;
    const pageWidthEMU = this.width * 635;
    const pageHeightEMU = this.height * 635;
    const headerXML = create({
      encoding: 'UTF-8',
      standalone: true,
      namespaceAlias: {
        w: namespaces.w,
        r: namespaces.r,
        wp: namespaces.wp,
        a: namespaces.a,
        pic: namespaces.pic,
        ve: namespaces.ve,
        o: namespaces.o,
        v: namespaces.v,
        w10: namespaces.w10,
      },
    }).ele('@w', 'hdr');

    // Reset these for each header generation
    this.backgroundImageHeight = 0;
    this.logoHeights = [];
    this.vTreeHeight = 0; // Add this reset

    let headerHeight = null;

    // Add page background if present in headerConfig
    if (headerConfig && headerConfig.pageBackground && headerConfig.pageBackground.url) {
      await this.addBackgroundToHeader(
        headerXML,
        {
          backgroundUrl: headerConfig.pageBackground.url,
          backgroundSize: headerConfig.pageBackground.size,
          backgroundPosition: headerConfig.pageBackground.position,
          backgroundRepeat: headerConfig.pageBackground.repeat,
        },
        headerId
      );
    }

    // Process vTree content
    if (vTree) {
      const XMLFragment = fragment();
      await convertVTreeToXML(this, vTree, XMLFragment);

      const firstParagraph = XMLFragment.first();
      if (
        firstParagraph.node?.nodeName !== 'p' &&
        ((headerConfig && headerConfig.logos && Array.isArray(headerConfig.logos)) ||
          (headerConfig && headerConfig.backgroundImage && headerConfig.backgroundImage?.url))
      ) {
        const paragraphFragment = await xmlBuilder.buildParagraph(vTree, {}, this);
        headerXML.import(paragraphFragment);
        headerXML.import(XMLFragment);
      } else {
        headerXML.import(XMLFragment);
      }

      this.vTreeHeight = Math.ceil(this.estimateVTreeHeight(XMLFragment) / 635);
    }

    // Process other headerConfig items
    if (headerConfig) {
      if (headerConfig.backgroundImage && headerConfig.backgroundImage?.url) {
        const backgroundHeight = await this.addBackgroundImage(
          headerXML,
          headerConfig.backgroundImage,
          headerId,
          pageWidthEMU,
          pageHeightEMU,
          'header'
        );
        this.backgroundImageHeight = Math.ceil(backgroundHeight / 635);
      }
      if (headerConfig.logos && Array.isArray(headerConfig.logos)) {
        // eslint-disable-next-line no-restricted-syntax
        for (const logo of headerConfig.logos) {
          // eslint-disable-next-line no-await-in-loop
          const logoHeight = await this.addLogo(headerXML, logo, headerId);
          this.logoHeights.push(Math.ceil(logoHeight / 635));
        }
      }
    }

    // Calculate final height
    if (vTree || headerConfig) {
      headerHeight = this.calculateHeaderHeight();
    }

    // Store header object by type name for sectPr references
    this.headerObjects[headerTypeName] = {
      headerId: `${headerId}`,
      height: headerHeight,
    };

    return { headerId: `${headerId}`, headerXML, headerHeight, typeName: headerTypeName };
  }

  async generateFooterXML(vTree, footerConfig, footerTypeName = 'default') {
    const footerId = this.lastFooterId + 1;
    this.lastFooterId = footerId;
    const pageWidthEMU = this.width * 635;
    const pageHeightEMU = this.height * 635;
    const footerXML = create({
      encoding: 'UTF-8',
      standalone: true,
      namespaceAlias: {
        w: namespaces.w,
        r: namespaces.r,
        wp: namespaces.wp,
        a: namespaces.a,
        pic: namespaces.pic,
        ve: namespaces.ve,
        o: namespaces.o,
        v: namespaces.v,
        w10: namespaces.w10,
      },
    }).ele('@w', 'ftr');

    this.backgroundImageHeight = 0;
    this.logoHeights = [];

    let footerHeight = null;

    if (vTree) {
      const XMLFragment = fragment();
      await convertVTreeToXML(this, vTree, XMLFragment);

      const firstParagraph = XMLFragment.first();
      if (
        firstParagraph.node?.nodeName !== 'p' &&
        ((footerConfig.logos && Array.isArray(footerConfig?.logos)) ||
          (footerConfig.backgroundImage && footerConfig.backgroundImage?.url))
      ) {
        const paragraphFragment = await xmlBuilder.buildParagraph(vTree, {}, this);
        footerXML.import(paragraphFragment);
        footerXML.import(XMLFragment);
      } else {
        footerXML.import(XMLFragment);
      }

      this.vTreeHeight = Math.ceil(this.estimateVTreeHeight(XMLFragment) / 635); // Convert EMUs to TWIPs
    }

    if (footerConfig) {
      if (footerConfig.backgroundImage && footerConfig.backgroundImage?.url) {
        const backgroundHeight = await this.addBackgroundImage(
          footerXML,
          footerConfig.backgroundImage,
          footerId,
          pageWidthEMU,
          pageHeightEMU,
          'footer'
        );
        this.backgroundImageHeight = Math.ceil(backgroundHeight / 635); // Convert EMUs to TWIPs
      }
      if (footerConfig.logos && Array.isArray(footerConfig.logos)) {
        // eslint-disable-next-line no-restricted-syntax
        for (const logo of footerConfig.logos) {
          // eslint-disable-next-line no-await-in-loop
          const logoHeight = await this.addLogo(footerXML, logo, footerId);
          this.logoHeights.push(Math.ceil(logoHeight / 635)); // Convert EMUs to TWIPs
        }
      }
    }

    if (footerConfig) {
      // Calculate the footer height
      footerHeight = this.calculateFooterHeight();
    }

    // Store header object by type name for sectPr references
    this.footerObjects[footerTypeName] = {
      headerId: `rId${footerId}`,
      height: footerHeight,
    };

    // Return footerId, footerXML, and calculated footerHeight
    return { footerId, footerXML, footerHeight };
  }

  // eslint-disable-next-line class-methods-use-this
  estimateVTreeHeight(xmlFragment) {
    const xmlString = xmlFragment.toString();

    // Key structure detection patterns
    const tableRegex =
      /(?:<w:tbl\b[^>]*>|<tbl\b[^>]*xmlns="http:\/\/schemas\.openxmlformats\.org\/wordprocessingml\/2006\/main"[^>]*>)/g;
    const tableRowRegex = /(?:<w:tr\b[^>]*>|<tr\b[^>]*>)/g;
    const drawingRegex = /(?:<w:drawing\b[^>]*>|<drawing\b[^>]*>)/g;
    const textRegex = /(?:<w:t\b[^>]*>|<t\b[^>]*>)([^<]*)(?:<\/w:t>|<\/t>)/g;

    // Count key elements
    const hasTables = xmlString.match(tableRegex) !== null;
    const tableRows = (xmlString.match(tableRowRegex) || []).length;
    const images = (xmlString.match(drawingRegex) || []).length;

    // Basic text content check
    const textMatches = [...xmlString.matchAll(textRegex)];
    const hasSubstantialText =
      textMatches.reduce((total, match) => total + (match[1] ? match[1].length : 0), 0) > 50;

    // Base height - minimal for simple text footers
    let baseHeight = 360000; // 0.25 inches

    // Add for structure complexity
    if (hasTables) {
      // Base table structure
      baseHeight = 432000; // 0.3 inches for table footers

      // Add small increment per row (capped)
      if (tableRows > 1) {
        baseHeight += Math.min(tableRows - 1, 3) * 72000; // 0.05" per additional row, max 3
      }
    }

    // Add for images (conservatively)
    if (images > 0) {
      baseHeight += Math.min(images, 3) * 36000; // 0.025" per image, max 3 images
    }

    // Add for substantial text
    if (hasSubstantialText && !hasTables) {
      baseHeight += 72000; // 0.05" extra for text-heavy footers
    }

    return baseHeight;
  }

  // eslint-disable-next-line consistent-return
  async addBackgroundImage(
    XML,
    backgroundImage,
    uniqId,
    pageWidthEMU,
    pageHeightEMU,
    type = 'header'
  ) {
    const relationshipType = type === 'header' ? `header${uniqId}` : `footer${uniqId}`;
    const styleVal = type === 'header' ? 'Header' : 'Footer';
    const { url } = backgroundImage;
    const { base64String, imageWidth, imageHeight } = await this.fetchImageAndGetDimensions(url);
    const imageFile = this.createMediaFile(base64String);
    const imageRelationshipId = this.createDocumentRelationships(
      relationshipType,
      imageType,
      `media/${imageFile.fileNameWithExtension}`,
      'Internal'
    );
    // Add the image file to the zip
    this.zip
      .folder('word/media')
      .file(imageFile.fileNameWithExtension, imageFile.fileContent, { base64: true });

    // Calculate the height while maintaining aspect ratio
    const aspectRatio = imageWidth / imageHeight;
    const imageHeightEMU = Math.round(pageWidthEMU / aspectRatio);
    XML.ele('@w', 'p')
      .ele('@w', 'pPr')
      .ele('@w', 'pStyle')
      .att('@w', 'val', styleVal)
      .up()
      .up()
      .ele('@w', 'r')
      .ele('@w', 'drawing')
      .ele('@wp', 'anchor')
      .att('behindDoc', '1')
      .att('distT', '0')
      .att('distB', '0')
      .att('distL', '0')
      .att('distR', '0')
      .att('simplePos', '0')
      .att('relativeHeight', '0')
      .att('locked', '0')
      .att('layoutInCell', '1')
      .att('allowOverlap', '1')
      .ele('@wp', 'simplePos')
      .att('x', '0')
      .att('y', '0')
      .up()
      .ele('@wp', 'positionH')
      .att('relativeFrom', 'page')
      .ele('@wp', 'posOffset')
      .txt('0')
      .up()
      .up()
      .ele('@wp', 'positionV')
      .att('relativeFrom', 'page')
      .ele('@wp', 'posOffset')
      .txt(type === 'header' ? '0' : (pageHeightEMU - imageHeightEMU).toString())
      .up()
      .up()
      .ele('@wp', 'extent')
      .att('cx', pageWidthEMU)
      .att('cy', imageHeightEMU)
      .up()
      .ele('@wp', 'effectExtent')
      .att('l', '0')
      .att('t', '0')
      .att('r', '0')
      .att('b', '0')
      .up()
      .ele('@wp', 'wrapNone')
      .up()
      .ele('@wp', 'docPr')
      .att('id', '1')
      .att('name', 'Background Picture')
      .up()
      .ele('@wp', 'cNvGraphicFramePr')
      .ele('@a', 'graphicFrameLocks')
      .att('noChangeAspect', '1')
      .up()
      .up()
      .ele('@a', 'graphic')
      .ele('@a', 'graphicData')
      .att('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
      .ele('@pic', 'pic')
      .ele('@pic', 'nvPicPr')
      .ele('@pic', 'cNvPr')
      .att('id', '0')
      .att('name', 'Background Picture')
      .up()
      .ele('@pic', 'cNvPicPr')
      .up()
      .up()
      .ele('@pic', 'blipFill')
      .ele('@a', 'blip')
      .att('@r', 'embed', `rId${imageRelationshipId}`)
      .up()
      .ele('@a', 'stretch')
      .ele('@a', 'fillRect')
      .up()
      .up()
      .up()
      .ele('@pic', 'spPr')
      .ele('@a', 'xfrm')
      .ele('@a', 'off')
      .att('x', '0')
      .att('y', '0')
      .up()
      .ele('@a', 'ext')
      .att('cx', pageWidthEMU)
      .att('cy', imageHeightEMU)
      .up()
      .up()
      .ele('@a', 'prstGeom')
      .att('prst', 'rect')
      .ele('@a', 'avLst')
      .up()
      .up()
      .up()
      .up()
      .up()
      .up()
      .up()
      .up()
      .up()
      .up();

    return imageHeightEMU;
  }

  // eslint-disable-next-line consistent-return
  async getImageHeight(image) {
    const { url, width, height } = image;

    const { imageWidth, imageHeight } = await this.fetchImageAndGetDimensions(url);

    // Calculate dimensions based on provided width or height, maintaining aspect ratio
    const aspectRatio = imageWidth / imageHeight;
    let widthEMU;
    if (width && !height) {
      widthEMU = Math.round(parseFloat(width) * 9525); // 1 px = 9525 EMUs
      return Math.round(widthEMU / aspectRatio);
    }
    if (!width && height) {
      return Math.round(parseFloat(height) * 9525);
    }
    if (width && height) {
      return Math.round(parseFloat(height) * 9525);
    }
    // If neither width nor height is provided, use original dimensions
    return Math.round(imageHeight * 9525);
  }

  async addLogo(headerXML, logo, headerId) {
    const { url, width, height, alignment } = logo;

    try {
      const { base64String, imageWidth, imageHeight } = await this.fetchImageAndGetDimensions(url);
      const imageFile = this.createMediaFile(base64String);
      const imageRelationshipId = this.createDocumentRelationships(
        `header${headerId}`,
        imageType,
        `media/${imageFile.fileNameWithExtension}`,
        'Internal'
      );

      // Add the image file to the zip
      this.zip
        .folder('word/media')
        .file(imageFile.fileNameWithExtension, imageFile.fileContent, { base64: true });

      // Calculate dimensions based on provided width or height, maintaining aspect ratio
      const aspectRatio = imageWidth / imageHeight;
      let widthEMU;
      let heightEMU;

      if (width && !height) {
        widthEMU = Math.round(parseFloat(width) * 9525); // 1 px = 9525 EMUs
        heightEMU = Math.round(widthEMU / aspectRatio);
      } else if (!width && height) {
        heightEMU = Math.round(parseFloat(height) * 9525);
        widthEMU = Math.round(heightEMU * aspectRatio);
      } else if (width && height) {
        widthEMU = Math.round(parseFloat(width) * 9525);
        heightEMU = Math.round(parseFloat(height) * 9525);
      } else {
        // If neither width nor height is provided, use original dimensions
        widthEMU = Math.round(imageWidth * 9525);
        heightEMU = Math.round(imageHeight * 9525);
      }

      // Find the first paragraph or create a new paragraph for the header
      let lastParagraph = null;
      let paragraphProperties = null;
      try {
        lastParagraph = headerXML.first('@w', 'p');
      } catch (e) {
        lastParagraph = headerXML.ele('@w', 'p');
      }

      if (!lastParagraph) {
        lastParagraph = headerXML.ele('@w', 'p');
      }

      // Add a paragraph style if not exists
      try {
        paragraphProperties = lastParagraph.first('@w', 'pPr');
      } catch (e) {
        lastParagraph.ele('@w', 'pPr').ele('@w', 'pStyle').att('@w', 'val', 'Header');
      }

      if (!paragraphProperties) {
        lastParagraph.ele('@w', 'pPr').ele('@w', 'pStyle').att('@w', 'val', 'Header');
      }

      const pageWidthTwips = 11906; // in twips
      const pageWidthEMU = pageWidthTwips * 635; // converts twips to EMUs
      const leftOffset = 180000; // 1 cm gap from the left edge
      const centerOffset = (pageWidthEMU - widthEMU) / 2;
      const rightOffset = pageWidthEMU - widthEMU - 180000; // 1 cm gap from the right edge
      // eslint-disable-next-line no-nested-ternary
      const posOffset =
        // eslint-disable-next-line no-nested-ternary
        alignment === 'center' ? centerOffset : alignment === 'right' ? rightOffset : leftOffset;

      // Create the run with the drawing
      const run = lastParagraph.ele('@w', 'r');

      // Create the drawing with anchor
      // eslint-disable-next-line no-unused-vars
      const drawing = run
        .ele('@w', 'drawing')
        .ele('@wp', 'anchor')
        .att('behindDoc', '1')
        .att('distT', '0')
        .att('distB', '0')
        .att('distL', '0')
        .att('distR', '0')
        .att('simplePos', '0')
        .att('relativeHeight', '0')
        .att('locked', '0')
        .att('layoutInCell', '1')
        .att('allowOverlap', '1')
        .ele('@wp', 'simplePos')
        .att('x', '0')
        .att('y', '0')
        .up()
        .ele('@wp', 'positionH')
        .att('relativeFrom', 'page') // Use 'page' for the width of the document context
        .ele('@wp', 'posOffset')
        .txt(posOffset)
        .up()
        .up()
        .ele('@wp', 'positionV')
        .att('relativeFrom', 'paragraph') // Try 'paragraph' instead of 'page'
        .ele('@wp', 'posOffset')
        .txt('180000')
        .up()
        .up()
        .ele('@wp', 'extent')
        .att('cx', widthEMU)
        .att('cy', heightEMU)
        .up()
        .ele('@wp', 'effectExtent')
        .att('l', '0')
        .att('t', '0')
        .att('r', '0')
        .att('b', '0')
        .up()
        .ele('@wp', 'wrapNone')
        .up()
        .ele('@wp', 'docPr')
        .att('id', '1')
        .att('name', 'Logo')
        .up()
        .ele('@wp', 'cNvGraphicFramePr')
        .ele('@a', 'graphicFrameLocks')
        .att('noChangeAspect', '1')
        .up()
        .up()
        .ele('@a', 'graphic')
        .ele('@a', 'graphicData')
        .att('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
        .ele('@pic', 'pic')
        .ele('@pic', 'nvPicPr')
        .ele('@pic', 'cNvPr')
        .att('id', '0')
        .att('name', 'Logo')
        .up()
        .ele('@pic', 'cNvPicPr')
        .up()
        .up()
        .ele('@pic', 'blipFill')
        .ele('@a', 'blip')
        .att('@r', 'embed', `rId${imageRelationshipId}`)
        .up()
        .ele('@a', 'stretch')
        .ele('@a', 'fillRect')
        .up()
        .up()
        .up()
        .ele('@pic', 'spPr')
        .ele('@a', 'xfrm')
        .ele('@a', 'off')
        .att('x', '0')
        .att('y', '0')
        .up()
        .ele('@a', 'ext')
        .att('cx', widthEMU)
        .att('cy', heightEMU)
        .up()
        .up()
        .ele('@a', 'prstGeom')
        .att('prst', 'rect')
        .ele('@a', 'avLst')
        .up()
        .up()
        .up()
        .up()
        .up()
        .up()
        .up()
        .up();

      return heightEMU; // Return the height of the logo in EMUs
    } catch (error) {
      // eslint-disable-next-line no-console
      console.warn('Error processing logo:', error);
      return 0; // Return 0 if there was an error
    }
  }

  async fetchImageAndGetDimensions(url) {
    return new Promise((resolve, reject) => {
      const protocol = url.startsWith('https') ? https : http;
      protocol
        .get(url, (response) => {
          if (response.statusCode !== 200) {
            reject(new Error(`Failed to fetch image: ${response.statusCode}`));
            return;
          }

          const chunks = [];
          response.on('data', (chunk) => chunks.push(chunk));
          response.on('end', async () => {
            const buffer = Buffer.concat(chunks);
            const base64 = buffer.toString('base64');
            const mimeType = response.headers['content-type'];

            try {
              if (mimeType === 'image/svg+xml') {
                const svgString = buffer.toString('utf-8');
                const { width, height } = await this.getSVGDimensions(svgString);
                resolve({
                  base64String: `data:${mimeType};base64,${base64}`,
                  imageWidth: width,
                  imageHeight: height,
                  mimeType,
                });
              } else {
                // For other image types, use sharp
                const metadata = await sharp(buffer).metadata();
                resolve({
                  base64String: `data:${mimeType};base64,${base64}`,
                  imageWidth: metadata.width,
                  imageHeight: metadata.height,
                  mimeType,
                });
              }
            } catch (err) {
              reject(err);
            }
          });
        })
        .on('error', reject);
    });
  }

  async getSVGDimensions(svgString) {
    // First, try to get dimensions from viewBox
    const viewBoxMatch = svgString.match(/viewBox="([^"]+)"/);
    if (viewBoxMatch) {
      // eslint-disable-next-line no-unused-vars
      const [minX, minY, width, height] = viewBoxMatch[1].split(/\s+/).map(Number);
      return { width, height };
    }

    // If viewBox is not available, fall back to width and height attributes
    const widthMatch = svgString.match(/width="([^"]+)"/);
    const heightMatch = svgString.match(/height="([^"]+)"/);

    let width = widthMatch ? this.parseLength(widthMatch[1]) : null;
    let height = heightMatch ? this.parseLength(heightMatch[1]) : null;

    // If dimensions are still not found, use default values
    width = width || 300; // Default width
    height = height || 150; // Default height

    return { width, height };
  }

  // eslint-disable-next-line class-methods-use-this
  parseLength(length) {
    if (typeof length === 'number') return length;

    const match = length.match(/^(\d+(?:\.\d+)?)(px|pt|em|ex|%)?$/);
    if (!match) return null;

    const value = parseFloat(match[1]);
    const unit = match[2] || 'px';

    switch (unit) {
      case 'px':
        return value;
      case 'pt':
        return value * 1.33333; // 1pt = 1.33333px
      case 'em':
      case 'ex':
      case '%':
        // For relative units, we'll just use the numeric value as pixels
        return value;
      default:
        return value;
    }
  }

  addImage(imageUrl) {
    const imageFile = this.createMediaFile(imageUrl);
    this.zip
      .folder('word/media')
      .file(imageFile.fileNameWithExtension, imageFile.fileContent, { base64: true });
    return imageFile;
  }

  async embedFonts() {
    if (!this.fonts || !this.fonts.length) return;

    // eslint-disable-next-line no-restricted-syntax
    for (const font of this.fonts) {
      try {
        // Read the font file
        // eslint-disable-next-line no-await-in-loop
        let fontData = await fs.readFile(font.path);
        if (font.path.trim().toLowerCase().endsWith('.ttf')) {
          try {
            const { guid, data } = convertToODTTF(font.path);
            font.guid = guid;
            fontData = data;
          } catch (error) {
            // eslint-disable-next-line no-console
            console.warn('Error converting font:', error);
            throw new Error(`Error converting font: ${error}`);
          }
        }

        const base64Font = fontData.toString('base64');

        // Add font file to the zip
        const fontFileName = `${font.name.toLowerCase().replace(/\s+/g, '_')}.odttf`;
        this.zip.folder('word/fonts').file(fontFileName, base64Font, { base64: true });

        this.createDocumentRelationships(
          'fontTable',
          fontType,
          `fonts/${fontFileName.replaceAll(' ', '')}`,
          'Internal',
          {
            ridOverride: font.name.replaceAll(' ', ''),
          }
        );
      } catch (error) {
        // eslint-disable-next-line no-console
        console.warn(`Error embedding font ${font.name}:`, error);
      }
    }

    // Generate and add font table XML once guids have been created
    const fontTableXML = this.generateFontTableXML();
    this.zip.file('word/fontTable.xml', fontTableXML);
  }
}

export default DocxDocument;
