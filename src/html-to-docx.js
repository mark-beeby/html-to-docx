import { create } from 'xmlbuilder2';
import VNode from 'virtual-dom/vnode/vnode';
import VText from 'virtual-dom/vnode/vtext';
// eslint-disable-next-line import/no-named-default
import { default as HTMLToVDOM } from 'html-to-vdom';
import { decode } from 'html-entities';

import { relsXML } from './schemas';
import DocxDocument from './docx-document';
import { renderDocumentFile } from './helpers';
import {
  pixelRegex,
  pixelToTWIP,
  cmRegex,
  cmToTWIP,
  inchRegex,
  inchToTWIP,
  pointRegex,
  pointToHIP,
} from './utils/unit-conversion';
import {
  defaultDocumentOptions,
  relsFolderName,
  headerFileName,
  footerFileName,
  themeFileName,
  documentFileName,
  headerType,
  footerType,
  internalRelationship,
  wordFolder,
  themeFolder,
  themeType,
} from './constants';

const convertHTML = HTMLToVDOM({
  VNode,
  VText,
});

const mergeOptions = (options, patch) => ({ ...options, ...patch });

const fixupFontSize = (fontSize) => {
  let normalizedFontSize;
  if (pointRegex.test(fontSize)) {
    const matchedParts = fontSize.match(pointRegex);

    normalizedFontSize = pointToHIP(matchedParts[1]);
  } else if (fontSize) {
    // assuming it is already in HIP
    normalizedFontSize = fontSize;
  } else {
    normalizedFontSize = null;
  }

  return normalizedFontSize;
};

const normalizeUnits = (dimensioningObject, defaultDimensionsProperty) => {
  let normalizedUnitResult = {};
  if (typeof dimensioningObject === 'object' && dimensioningObject !== null) {
    Object.keys(dimensioningObject).forEach((key) => {
      if (pixelRegex.test(dimensioningObject[key])) {
        const matchedParts = dimensioningObject[key].match(pixelRegex);
        normalizedUnitResult[key] = pixelToTWIP(matchedParts[1]);
      } else if (cmRegex.test(dimensioningObject[key])) {
        const matchedParts = dimensioningObject[key].match(cmRegex);
        normalizedUnitResult[key] = cmToTWIP(matchedParts[1]);
      } else if (inchRegex.test(dimensioningObject[key])) {
        const matchedParts = dimensioningObject[key].match(inchRegex);
        normalizedUnitResult[key] = inchToTWIP(matchedParts[1]);
      } else if (dimensioningObject[key]) {
        normalizedUnitResult[key] = dimensioningObject[key];
      } else {
        // incase value is something like 0
        normalizedUnitResult[key] = defaultDimensionsProperty[key];
      }
    });
  } else {
    // eslint-disable-next-line no-param-reassign
    normalizedUnitResult = null;
  }

  return normalizedUnitResult;
};

const normalizeDocumentOptions = (documentOptions) => {
  const normalizedDocumentOptions = { ...documentOptions };
  Object.keys(documentOptions).forEach((key) => {
    // eslint-disable-next-line default-case
    switch (key) {
      case 'pageSize':
      case 'margins':
        normalizedDocumentOptions[key] = normalizeUnits(
          documentOptions[key],
          defaultDocumentOptions[key]
        );
        break;
      case 'fontSize':
      case 'complexScriptFontSize':
        normalizedDocumentOptions[key] = fixupFontSize(documentOptions[key]);
        break;
    }
  });

  return normalizedDocumentOptions;
};

function generateEmptyHeaderAndFooter(docxDocument, zip) {
  // generate an empty header and footer xml in case these require suppression
  // eslint-disable-next-line no-param-reassign
  docxDocument.relationshipFilename = headerFileName;
  const { headerId, headerXML } = docxDocument.generateEmptyHeaderXML();
  // eslint-disable-next-line no-param-reassign
  docxDocument.relationshipFilename = footerFileName;
  const { footerId, footerXML } = docxDocument.generateEmptyFooterXML();
  // eslint-disable-next-line no-param-reassign
  docxDocument.relationshipFilename = documentFileName;
  [
    {
      type: headerType,
      id: headerId,
      xml: headerXML,
    },
    {
      type: footerType,
      id: footerId,
      xml: footerXML,
    },
  ].forEach(({ type, id, xml }) => {
    const fileNameWithExt = `${type}${id}.xml`;
    const relationshipId = docxDocument.createDocumentRelationships(
      docxDocument.relationshipFilename,
      type,
      fileNameWithExt,
      internalRelationship
    );
    zip
      .folder(wordFolder)
      .file(fileNameWithExt, xml.toString({ prettyPrint: true }), { createFolders: false });
    // eslint-disable-next-line no-param-reassign
    docxDocument[`${type}Objects`].none = { id, relationshipId, type };
  });
}

// Ref: https://en.wikipedia.org/wiki/Office_Open_XML_file_formats
// http://officeopenxml.com/anatomyofOOXML.php
async function addFilesToContainer(
  zip,
  htmlString,
  suppliedDocumentOptions,
  headerHTMLString,
  footerHTMLString,
  headerConfig,
  footerConfig
) {
  const normalizedDocumentOptions = normalizeDocumentOptions(suppliedDocumentOptions);
  const documentOptions = mergeOptions(defaultDocumentOptions, normalizedDocumentOptions);

  if (documentOptions.header && !headerHTMLString) {
    // eslint-disable-next-line no-param-reassign
    headerHTMLString = '';
  }
  if (documentOptions.footer && !footerHTMLString) {
    // eslint-disable-next-line no-param-reassign
    footerHTMLString = '';
  }
  if (documentOptions.decodeUnicode) {
    // eslint-disable-next-line no-param-reassign
    headerHTMLString = decode(headerHTMLString);
    // eslint-disable-next-line no-param-reassign
    htmlString = decode(htmlString);
    // eslint-disable-next-line no-param-reassign
    footerHTMLString = decode(footerHTMLString);
  }

  documentOptions.suppressFooterMargins = footerHTMLString && footerHTMLString.length;

  const docxDocument = new DocxDocument({ zip, htmlString, ...documentOptions });

  // Embed fonts before generating document XML
  await docxDocument.embedFonts();

  zip
    .folder(relsFolderName)
    .file(
      '.rels',
      create({ encoding: 'UTF-8', standalone: true }, relsXML).toString({ prettyPrint: true }),
      { createFolders: false }
    );
  zip.folder('docProps').file('core.xml', docxDocument.generateCoreXML(), { createFolders: false });

  generateEmptyHeaderAndFooter(docxDocument, zip);

  if (docxDocument.header && (headerHTMLString || headerConfig)) {
    const vTree = headerHTMLString ? convertHTML(headerHTMLString) : null;
    docxDocument.relationshipFilename = headerFileName;
    const { headerId, headerXML, headerHeight } = await docxDocument.generateHeaderXML(
      vTree,
      headerConfig
    );

    if (headerHeight !== null) {
      docxDocument.margins.header = 300;
      docxDocument.margins.top = Math.max(headerHeight + 180, docxDocument.margins.top);
    }

    docxDocument.relationshipFilename = documentFileName;
    const fileNameWithExt = `${headerType}${headerId}.xml`;
    const relationshipId = docxDocument.createDocumentRelationships(
      docxDocument.relationshipFilename,
      headerType,
      fileNameWithExt,
      internalRelationship
    );

    zip
      .folder(wordFolder)
      .file(fileNameWithExt, headerXML.toString({ prettyPrint: true }), { createFolders: false });
    docxDocument.headerObjects.default = {
      headerId,
      relationshipId,
      type: docxDocument.headerType,
    };
  }

  // Handle footer in similar way
  if (docxDocument.footer && (footerHTMLString || footerConfig)) {
    const vTree = footerHTMLString ? convertHTML(footerHTMLString) : null;
    docxDocument.relationshipFilename = footerFileName;
    const { footerId, footerXML, footerHeight } = await docxDocument.generateFooterXML(
      vTree,
      footerConfig
    );

    if (footerHeight !== null) {
      docxDocument.margins.footer = 300;
      docxDocument.margins.bottom = Math.max(footerHeight + 180, docxDocument.margins.bottom);
    }

    docxDocument.relationshipFilename = documentFileName;
    const fileNameWithExt = `${footerType}${footerId}.xml`;
    const relationshipId = docxDocument.createDocumentRelationships(
      docxDocument.relationshipFilename,
      footerType,
      fileNameWithExt,
      internalRelationship
    );

    zip
      .folder(wordFolder)
      .file(fileNameWithExt, footerXML.toString({ prettyPrint: true }), { createFolders: false });
    docxDocument.footerObjects.default = {
      footerId,
      relationshipId,
      type: docxDocument.footerType,
    };
  }

  docxDocument.documentXML = await renderDocumentFile(docxDocument);

  const themeFileNameWithExt = `${themeFileName}.xml`;
  docxDocument.createDocumentRelationships(
    docxDocument.relationshipFilename,
    themeType,
    `${themeFolder}/${themeFileNameWithExt}`,
    internalRelationship
  );

  zip
    .folder(wordFolder)
    .folder(themeFolder)
    .file(themeFileNameWithExt, docxDocument.generateThemeXML(), { createFolders: false });
  // require('fs').writeFileSync(`document.xml`, docxDocument.generateDocumentXML());
  zip
    .folder(wordFolder)
    .file('document.xml', docxDocument.generateDocumentXML(), { createFolders: false })
    .file('fontTable.xml', docxDocument.generateFontTableXML(), { createFolders: false })
    .file('styles.xml', docxDocument.generateStylesXML(), { createFolders: false })
    .file('numbering.xml', docxDocument.generateNumberingXML(), { createFolders: false })
    .file('settings.xml', docxDocument.generateSettingsXML(), { createFolders: false })
    .file('webSettings.xml', docxDocument.generateWebSettingsXML(), { createFolders: false });

  const relationshipXMLs = docxDocument.generateRelsXML();
  if (relationshipXMLs && Array.isArray(relationshipXMLs)) {
    relationshipXMLs.forEach(({ fileName, xmlString }) => {
      zip
        .folder(wordFolder)
        .folder(relsFolderName)
        .file(`${fileName}.xml.rels`, xmlString, { createFolders: false });
    });
  }

  zip.file('[Content_Types].xml', docxDocument.generateContentTypesXML(), { createFolders: false });

  return zip;
}

export default addFilesToContainer;
