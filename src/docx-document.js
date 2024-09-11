import { create, fragment } from 'xmlbuilder2';
import { nanoid } from 'nanoid';
import https from 'https';
import http from 'http';
import { URL } from 'url';
import sharp from 'sharp';
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
  defaultDocumentOptions,
} from './constants';
import ListStyleBuilder from './utils/list';
import { fontFamilyToTableObject } from './utils/font-family-conversion';

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
    this.stylesObjects = [];
    this.numberingObjects = [];
    this.fontTableObjects = [];
    this.relationshipFilename = documentFileName;
    this.relationships = [{ fileName: documentFileName, lastRelsId: 5, rels: [] }];
    this.mediaFiles = [];
    this.headerObjects = [];
    this.footerObjects = [];
    this.documentXML = null;

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

    this.ListStyleBuilder = new ListStyleBuilder(properties.numbering);
  }

  generateContentTypesXML() {
    const contentTypesXML = create({ encoding: 'UTF-8', standalone: true }, contentTypesXMLString);

    generateContentTypesFragments(contentTypesXML, 'header', this.headerObjects);
    generateContentTypesFragments(contentTypesXML, 'footer', this.footerObjects);

    return contentTypesXML.toString({ prettyPrint: true });
  }

  generateDocumentXML() {
    const documentXML = create(
      { encoding: 'UTF-8', standalone: true },
      generateDocumentTemplate(this.width, this.height, this.orientation, this.margins)
    );

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

    // Modify section properties for responsive header
    if (this.sectionProperties) {
      const body = documentXML.root().first();

      // Iterate through child elements to find and modify existing sectPr
      body.each((child) => {
        if (child.node.nodeName === 'w:sectPr') {
          // Remove the existing sectPr
          child.remove();
        }
      });

      // Append the new sectPr at the end of the body
      body.import(this.sectionProperties.root());
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
    return generateXMLString(settingsXMLString);
  }

  // eslint-disable-next-line class-methods-use-this
  generateWebSettingsXML() {
    return generateXMLString(webSettingsXMLString);
  }

  generateStylesXML() {
    return generateXMLString(
      generateStylesXML(this.font, this.fontSize, this.complexScriptFontSize, this.lang)
    );
  }

  generateFontTableXML() {
    const fontTableXML = create({ encoding: 'UTF-8', standalone: true }, fontTableXMLString);
    const fontNames = [
      'Arial',
      'Calibri',
      'Calibri Light',
      'Courier New',
      'Symbol',
      'Times New Roman',
    ];
    this.fontTableObjects.forEach(({ fontName, genericFontName }) => {
      if (!fontNames.includes(fontName)) {
        fontNames.push(fontName);
        const fontFragment = fragment({
          namespaceAlias: { w: namespaces.w },
        })
          .ele('@w', 'font')
          .att('@w', 'name', fontName);

        switch (genericFontName) {
          case 'serif':
            fontFragment.ele('@w', 'altName').att('@w', 'val', 'Times New Roman');
            fontFragment.ele('@w', 'family').att('@w', 'val', 'roman');
            fontFragment.ele('@w', 'pitch').att('@w', 'val', 'variable');
            break;
          case 'sans-serif':
            fontFragment.ele('@w', 'altName').att('@w', 'val', 'Arial');
            fontFragment.ele('@w', 'family').att('@w', 'val', 'swiss');
            fontFragment.ele('@w', 'pitch').att('@w', 'val', 'variable');
            break;
          case 'monospace':
            fontFragment.ele('@w', 'altName').att('@w', 'val', 'Courier New');
            fontFragment.ele('@w', 'family').att('@w', 'val', 'modern');
            fontFragment.ele('@w', 'pitch').att('@w', 'val', 'fixed');
            break;
          default:
            break;
        }

        fontTableXML.root().import(fontFragment);
      }
    });

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
          .up()
          .ele('@w', 'lvlText')
          .att(
            '@w',
            'val',
            type === 'ol' ? this.ListStyleBuilder.getListPrefixSuffix(properties.style, level) : ''
          )
          .up()
          .ele('@w', 'lvlJc')
          .att('@w', 'val', 'left')
          .up()
          .ele('@w', 'pPr')
          .ele('@w', 'tabs')
          .ele('@w', 'tab')
          .att('@w', 'val', 'num')
          .att('@w', 'pos', (level + 1) * 720)
          .up()
          .up()
          .ele('@w', 'ind')
          .att('@w', 'left', (level + 1) * 720)
          .att('@w', 'hanging', 360)
          .up()
          .up()
          .up();

        if (type === 'ul') {
          levelFragment.last().import(
            fragment({ namespaceAlias: { w: namespaces.w } })
              .ele('@w', 'rPr')
              .ele('@w', 'rFonts')
              .att('@w', 'ascii', 'Symbol')
              .att('@w', 'hAnsi', 'Symbol')
              .att('@w', 'hint', 'default')
              .up()
              .up()
          );
        }
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
    this.numberingObjects.push({ numberingId: this.lastNumberingId, type, properties });

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

  createDocumentRelationships(fileName = 'document', type, target, targetMode = 'External') {
    let relationshipObject = this.relationships.find(
      (relationship) => relationship.fileName === fileName
    );
    let lastRelsId = 1;
    if (relationshipObject) {
      lastRelsId = relationshipObject.lastRelsId + 1;
      relationshipObject.lastRelsId = lastRelsId;
    } else {
      relationshipObject = { fileName, lastRelsId, rels: [] };
      this.relationships.push(relationshipObject);
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
    }

    relationshipObject.rels.push({
      relationshipId: lastRelsId,
      type: relationshipType,
      target,
      targetMode,
    });

    return lastRelsId;
  }

  adjustSectionProperties() {
    // const headerHeight = Math.min(Math.max(imageHeightTwips || 0, 0), 9000);

    this.sectionProperties = create({ encoding: 'UTF-8', standalone: true }).ele('w:sectPr', {
      'xmlns:w': namespaces.w,
      'xmlns:r': namespaces.r,
    });

    if (this.width && this.height) {
      this.sectionProperties
        .ele('w:pgSz')
        .att('w:w', this.width.toString())
        .att('w:h', this.height.toString())
        .up();
    }
    if (this.pageNumberStart) {
      this.sectionProperties
        .ele('w:pgNumType')
        .att('w:start', this.pageNumberStart.toString())
        .up();
    }

    if (this.headerObjects && this.headerObjects[0] && this.headerObjects[0].relationshipId) {
      this.sectionProperties
        .ele('w:headerReference')
        .att('w:type', 'default')
        .att('r:id', `rId${this.headerObjects[0].relationshipId}`)
        .up();
    }

    return this.sectionProperties;
  }

  async generateHeaderXML(vTree, headerImageUrl) {
    const headerId = this.lastHeaderId + 1;
    this.lastHeaderId = headerId;

    // Convert page width from twips to EMUs (1 twip = 635 EMUs)
    const pageWidthEMU = this.width * 635;

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

    if (headerImageUrl) {
      try {
        // Fetch the image and get its dimensions
        const { base64String, width, height } = await new Promise((resolve, reject) => {
          const url = new URL(headerImageUrl);
          const protocol = url.protocol === 'https:' ? https : http;

          protocol
            .get(url, (response) => {
              if (response.statusCode !== 200) {
                reject(new Error(`Failed to fetch image: ${response.statusCode}`));
                return;
              }

              const chunks = [];
              response.on('data', (chunk) => chunks.push(chunk));
              response.on('end', () => {
                const buffer = Buffer.concat(chunks);
                const base64 = buffer.toString('base64');
                const mimeType = response.headers['content-type'];

                // Get image dimensions
                sharp(buffer)
                  .metadata()
                  .then((metadata) => {
                    resolve({
                      base64String: `data:${mimeType};base64,${base64}`,
                      width: metadata.width,
                      height: metadata.height,
                    });
                  })
                  .catch((err) => reject(err));
              });
            })
            .on('error', reject);
        });

        // Calculate the height while maintaining aspect ratio
        const aspectRatio = width / height;
        const imageHeightEMU = Math.round(pageWidthEMU / aspectRatio);

        // Store the calculated height in twips for later use
        this.headerHeight = Math.round(imageHeightEMU / 635);

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

        headerXML
          .ele('@w', 'p')
          .ele('@w', 'pPr')
          .ele('@w', 'pStyle')
          .att('@w', 'val', 'Header')
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
          .txt('0')
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
          .att('name', 'Picture 1')
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
          .att('name', 'Picture 1')
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
      } catch (error) {
        console.error('Error processing header image:', error);
      }
    }

    if (vTree) {
      const XMLFragment = fragment();
      await convertVTreeToXML(this, vTree, XMLFragment);
      headerXML.import(XMLFragment);
    }

    return { headerId, headerXML };
  }

  addImage(imageUrl) {
    const imageFile = this.createMediaFile(imageUrl);
    this.zip
      .folder('word/media')
      .file(imageFile.fileNameWithExtension, imageFile.fileContent, { base64: true });
    return imageFile;
  }

  generateFooterXML(vTree) {
    return this.generateSectionXML(vTree, 'footer');
  }
}

export default DocxDocument;
