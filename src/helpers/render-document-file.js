/* eslint-disable no-await-in-loop */
/* eslint-disable no-case-declarations */
import { fragment } from 'xmlbuilder2';
import VNode from 'virtual-dom/vnode/vnode';
import VText from 'virtual-dom/vnode/vtext';
import isVNode from 'virtual-dom/vnode/is-vnode';
import isVText from 'virtual-dom/vnode/is-vtext';
// eslint-disable-next-line import/no-named-default
import { default as HTMLToVDOM } from 'html-to-vdom';
import sizeOf from 'image-size';
import imageToBase64 from 'image-to-base64';
import mimeTypes from 'mime-types';

// FIXME: remove the cyclic dependency
// eslint-disable-next-line import/no-cycle
import * as xmlBuilder from './xml-builder';
// eslint-disable-next-line import/no-cycle
import { buildTableRow } from './xml-builder';
import namespaces from '../namespaces';
import { imageType, internalRelationship } from '../constants';
import { vNodeHasChildren } from '../utils/vnode';
import { isValidUrl } from '../utils/url';

const convertHTML = HTMLToVDOM({
  VNode,
  VText,
});

// eslint-disable-next-line consistent-return, no-shadow
export const buildImage = async (docxDocumentInstance, vNode, maximumWidth = null) => {
  let response = null;
  let base64Uri = null;
  try {
    const imageSource = vNode.properties.src;
    if (isValidUrl(imageSource)) {
      const base64String = await imageToBase64(imageSource).catch((error) => {
        // eslint-disable-next-line no-console
        console.warning(`skipping image download and conversion due to ${error}`);
      });

      if (base64String) {
        base64Uri = `data:${mimeTypes.lookup(imageSource)};base64, ${base64String}`;
      }
    } else {
      base64Uri = decodeURIComponent(vNode.properties.src);
    }
    if (base64Uri) {
      response = docxDocumentInstance.createMediaFile(base64Uri);
    }
  } catch (error) {
    // NOOP
  }
  if (response) {
    docxDocumentInstance.zip
      .folder('word')
      .folder('media')
      .file(response.fileNameWithExtension, Buffer.from(response.fileContent, 'base64'), {
        createFolders: false,
      });

    const documentRelsId = docxDocumentInstance.createDocumentRelationships(
      docxDocumentInstance.relationshipFilename,
      imageType,
      `media/${response.fileNameWithExtension}`,
      internalRelationship
    );

    const imageBuffer = Buffer.from(response.fileContent, 'base64');
    const imageProperties = sizeOf(imageBuffer);

    const imageFragment = await xmlBuilder.buildParagraph(
      vNode,
      {
        type: 'picture',
        inlineOrAnchored: true,
        relationshipId: documentRelsId,
        ...response,
        description: vNode.properties.alt,
        maximumWidth: maximumWidth || docxDocumentInstance.availableDocumentSpace,
        originalWidth: imageProperties.width,
        originalHeight: imageProperties.height,
      },
      docxDocumentInstance
    );

    return imageFragment;
  }
};

export const buildList = async (vNode, docxDocumentInstance, xmlFragment) => {
  // For lists, we want to remove any empty paragraphs that were added before
  const { lastChild } = xmlFragment.node;
  if (lastChild && lastChild.nodeName === 'w:p') {
    const textNodes = lastChild.getElementsByTagName('w:t');
    let hasText = false;
    // eslint-disable-next-line no-plusplus
    for (let i = 0; i < textNodes.length; i++) {
      if (textNodes[i].textContent.trim()) {
        hasText = true;
        break;
      }
    }
    if (!hasText) {
      xmlFragment.node.removeChild(lastChild);
    }
  }

  let vNodeObjects = [
    {
      node: vNode,
      level: 0,
      type: vNode.tagName,
      numberingId: docxDocumentInstance.createNumbering(vNode.tagName, {
        level: 0,
        style: {
          ...vNode.properties?.style,
          primaryColour:
            vNode.properties?.attributes?.['data-primary-colour'] || vNode.properties?.style?.color,
        },
      }),
    },
  ];
  while (vNodeObjects.length) {
    const tempVNodeObject = vNodeObjects.shift();

    if (
      isVText(tempVNodeObject.node) ||
      (isVNode(tempVNodeObject.node) && !['ul', 'ol', 'li'].includes(tempVNodeObject.node.tagName))
    ) {
      const paragraphFragment = await xmlBuilder.buildParagraph(
        tempVNodeObject.node,
        {
          numbering: { levelId: tempVNodeObject.level, numberingId: tempVNodeObject.numberingId },
        },
        docxDocumentInstance
      );

      xmlFragment.import(paragraphFragment);
    }

    if (
      tempVNodeObject.node.children &&
      tempVNodeObject.node.children.length &&
      ['ul', 'ol', 'li'].includes(tempVNodeObject.node.tagName)
    ) {
      const tempVNodeObjects = tempVNodeObject.node.children.reduce((accumulator, childVNode) => {
        if (['ul', 'ol'].includes(childVNode.tagName)) {
          accumulator.push({
            node: childVNode,
            level: tempVNodeObject.level + 1,
            type: childVNode.tagName,
            numberingId: docxDocumentInstance.createNumbering(childVNode.tagName, {
              level: tempVNodeObject.level + 1,
              style: {
                ...childVNode.properties?.style,
                primaryColour:
                  childVNode.properties?.attributes?.['data-primary-colour'] ||
                  childVNode.properties?.style?.color,
              },
            }),
          });
        } else {
          // eslint-disable-next-line no-lonely-if
          if (
            accumulator.length > 0 &&
            isVNode(accumulator[accumulator.length - 1].node) &&
            accumulator[accumulator.length - 1].node.tagName.toLowerCase() === 'p'
          ) {
            accumulator[accumulator.length - 1].node.children.push(childVNode);
          } else {
            const paragraphVNode = new VNode(
              'p',
              null,
              // eslint-disable-next-line no-nested-ternary
              isVText(childVNode)
                ? [childVNode]
                : // eslint-disable-next-line no-nested-ternary
                isVNode(childVNode)
                ? childVNode.tagName.toLowerCase() === 'li'
                  ? [...childVNode.children]
                  : [childVNode]
                : []
            );
            accumulator.push({
              // eslint-disable-next-line prettier/prettier, no-nested-ternary
              node: isVNode(childVNode)
                ? // eslint-disable-next-line prettier/prettier, no-nested-ternary
                  childVNode.tagName.toLowerCase() === 'li'
                  ? childVNode
                  : childVNode.tagName.toLowerCase() !== 'p'
                  ? paragraphVNode
                  : childVNode
                : // eslint-disable-next-line prettier/prettier
                  paragraphVNode,
              level: tempVNodeObject.level,
              type: tempVNodeObject.type,
              numberingId: tempVNodeObject.numberingId,
            });
          }
        }

        return accumulator;
      }, []);
      vNodeObjects = tempVNodeObjects.concat(vNodeObjects);
    }
  }

  return xmlFragment;
};

export async function convertVTreeToXML(docxDocumentInstance, vTree, xmlFragment) {
  if (!vTree) {
    return xmlFragment;
  }
  if (Array.isArray(vTree) && vTree.length) {
    // eslint-disable-next-line no-restricted-syntax
    for (const vNode of vTree) {
      await convertVTreeToXML(docxDocumentInstance, vNode, xmlFragment);
    }
  } else if (isVNode(vTree)) {
    if (
      vTree.properties &&
      vTree.properties.attributes &&
      vTree.properties.attributes['data-docx-column-group']
    ) {
      // Handle columns
      const columns = vTree.children.filter(
        (child) =>
          child.properties &&
          child.properties.attributes &&
          child.properties.attributes['data-docx-column']
      );
      if (columns.length > 0) {
        const tableFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('w:tbl');

        // Set parent table properties
        const tablePropertiesFragment = fragment({ namespaceAlias: { w: namespaces.w } })
          .ele('w:tblPr')
          .ele('w:tblW')
          .att('w:w', '5000')
          .att('w:type', 'pct')
          .up()
          .ele('w:jc')
          .att('w:val', 'left')
          .up()
          .ele('w:tblInd')
          .att('w:w', '0')
          .att('w:type', 'dxa')
          .up()
          .ele('w:tblLayout')
          .att('w:type', 'fixed')
          .up()
          .ele('w:tblCellMar')
          .ele('w:top')
          .att('w:w', '0')
          .att('w:type', 'dxa')
          .up()
          .ele('w:left')
          .att('w:w', '100')
          .att('w:type', 'dxa')
          .up()
          .ele('w:bottom')
          .att('w:w', '0')
          .att('w:type', 'dxa')
          .up()
          .ele('w:right')
          .att('w:w', '100')
          .att('w:type', 'dxa')
          .up()
          .up();

        tableFragment.import(tablePropertiesFragment);

        // Add colgroup based on columns
        const tableGridFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele(
          'w:tblGrid'
        );
        columns.forEach((column) => {
          const colWidth = parseInt(column.properties.attributes['data-docx-column'], 10) || 1;
          const colWidthTwips = Math.floor(
            (colWidth / 12) * docxDocumentInstance.availableDocumentSpace
          );
          tableGridFragment.ele('w:gridCol').att('w:w', colWidthTwips).up();
          // We want to align top these columns to help it behave like the HTML equivalent, as
          // when it's a normal table the default behaviour would be middle.
          // eslint-disable-next-line no-param-reassign
          column.properties.attributes = column.properties.attributes || {};
          // eslint-disable-next-line no-param-reassign
          column.properties.attributes.valign = 'top';
        });
        tableGridFragment.up();

        tableFragment.import(tableGridFragment);

        // Create the table row and cells for the columns
        const tableRowFragment = await buildTableRow(docxDocumentInstance, columns);
        tableFragment.import(tableRowFragment);

        tableFragment.up();
        xmlFragment.import(tableFragment);
        // eslint-disable-next-line consistent-return
        return;
      }
    }
    // eslint-disable-next-line no-use-before-define
    await findXMLEquivalent(docxDocumentInstance, vTree, xmlFragment);
  } else if (isVText(vTree)) {
    // Only create paragraphs for text nodes that have meaningful content
    if (vTree.text.trim()) {
      const paragraphFragment = await xmlBuilder.buildParagraph(vTree, {}, docxDocumentInstance);
      if (paragraphFragment) {
        xmlFragment.import(paragraphFragment);
      }
    }
  }
  return xmlFragment;
}

async function findXMLEquivalent(docxDocumentInstance, vNode, xmlFragment) {
  if (
    isVNode(vNode) &&
    vNode.tagName === 'div' &&
    vNode.properties &&
    vNode.properties.attributes &&
    vNode.properties.attributes['data-docx-column-group']
  ) {
    // nested column group requires processing from the top
    await convertVTreeToXML(docxDocumentInstance, vNode, xmlFragment);
    return;
  }
  if (
    isVNode(vNode) &&
    vNode.tagName === 'div' &&
    vNode.properties &&
    vNode.properties.attributes &&
    vNode.properties.attributes['data-docx-column']
  ) {
    return; // Handled in parent function to convert columns
  }

  // Handle both regular page breaks and data-section divs that require section breaks
  if (
    vNode.tagName === 'div' &&
    (vNode.properties.attributes.class === 'page-break' ||
      (vNode.properties.style && vNode.properties.style['page-break-after']) ||
      (vNode.properties.attributes.class && vNode.properties.attributes['data-section']))
  ) {
    const isProfilePage =
      vNode.properties.attributes.class && vNode.properties.attributes['data-section'];

    // Track section information in the DocxDocument instance
    if (isProfilePage && vNode.properties.attributes['data-section-break'] === 'true') {
      const sectionIndex = parseInt(
        vNode.properties.attributes['data-section-index'] || docxDocumentInstance.currentSectionId,
        10
      );
      const headerType = vNode.properties.attributes['data-header-type'] || 'default';
      const footerType = vNode.properties.attributes['data-footer-type'] || 'default';

      // Store section information
      let margins = null;
      if (vNode.properties.attributes['data-margins']) {
        try {
          margins = JSON.parse(vNode.properties.attributes['data-margins']);
        } catch (e) {
          console.error('Error parsing margins JSON:', e);
        }
      }

      docxDocumentInstance.sections.push({
        index: sectionIndex,
        headerType,
        footerType,
        margins,
      });
      // eslint-disable-next-line no-param-reassign
      docxDocumentInstance.currentSectionId = sectionIndex + 1;
    }

    // For data-section divs, we'll create a section break with its own properties
    if (isProfilePage) {
      // Create a section break with its own properties
      const sectionBreakPara = fragment({ namespaceAlias: { w: namespaces.w, r: namespaces.r } })
        .ele('w:p')
        .ele('w:r')
        .ele('w:br')
        .att('w:type', 'page')
        .up()
        .up()
        .up()
        .ele('w:p')
        .ele('w:pPr')
        .ele('w:sectPr')
        // Add section type to ensure it starts on a new page
        .ele('w:type')
        .att('w:val', 'nextPage')
        .up();

      // Add section properties based on custom attributes if present
      // Check if this data-section has custom margins
      const hasCustomMargins = vNode.properties.attributes['data-margins'];
      let margins = null;

      if (hasCustomMargins) {
        try {
          margins = JSON.parse(vNode.properties.attributes['data-margins']);
          sectionBreakPara
            .ele('w:pgMar')
            .att('w:top', margins.top || docxDocumentInstance.margins.top)
            .att('w:right', margins.right || docxDocumentInstance.margins.right)
            .att('w:bottom', margins.bottom || docxDocumentInstance.margins.bottom)
            .att('w:left', margins.left || docxDocumentInstance.margins.left)
            .att('w:header', margins.header || docxDocumentInstance.margins.header)
            .att('w:footer', margins.footer || docxDocumentInstance.margins.footer)
            .up();
        } catch (e) {
          console.error('Error parsing margins JSON:', e);
          // Use default page margins if parsing fails
          sectionBreakPara
            .ele('w:pgMar')
            .att('w:top', docxDocumentInstance.margins.top)
            .att('w:right', docxDocumentInstance.margins.right)
            .att('w:bottom', docxDocumentInstance.margins.bottom)
            .att('w:left', docxDocumentInstance.margins.left)
            .att('w:header', docxDocumentInstance.margins.header)
            .att('w:footer', docxDocumentInstance.margins.footer)
            .up();
        }
      } else {
        // Use default page margins if no custom margins are provided
        sectionBreakPara
          .ele('w:pgMar')
          .att('w:top', docxDocumentInstance.margins.top)
          .att('w:right', docxDocumentInstance.margins.right)
          .att('w:bottom', docxDocumentInstance.margins.bottom)
          .att('w:left', docxDocumentInstance.margins.left)
          .att('w:header', docxDocumentInstance.margins.header)
          .att('w:footer', docxDocumentInstance.margins.footer)
          .up();
      }

      // Check for header/footer options
      const headerType = vNode.properties.attributes['data-header-type'] || 'default';
      const footerType = vNode.properties.attributes['data-footer-type'] || 'default';
      // Add header reference if header exists
      if (docxDocumentInstance.headerObjects && docxDocumentInstance.headerObjects[headerType]) {
        if (headerType !== 'none') {
          // eslint-disable-next-line no-param-reassign
          docxDocumentInstance.headerAdded = true;
        }
        sectionBreakPara
          .ele('w:headerReference')
          .att('w:type', 'default')
          .att('r:id', `rId${docxDocumentInstance.headerObjects[headerType].relationshipId}`)
          .up();
      } else if (!docxDocumentInstance.headerAdded) {
        console.log(`Warning: Header type '${headerType}' not found in docxDocument.headerObjects`);
      }

      // Add footer reference if footer exists
      if (docxDocumentInstance.footerObjects && docxDocumentInstance.footerObjects[footerType]) {
        if (footerType !== 'none') {
          // eslint-disable-next-line no-param-reassign
          docxDocumentInstance.footerAdded = true;
        }
        sectionBreakPara
          .ele('w:footerReference')
          .att('w:type', 'default')
          .att('r:id', `rId${docxDocumentInstance.footerObjects[footerType].relationshipId}`)
          .up();
      } else if (!docxDocumentInstance.footerAdded) {
        console.log(`Warning: Footer type '${footerType}' not found in docxDocument.footerObjects`);
      }

      // Add page size for consistency
      sectionBreakPara
        .ele('w:pgSz')
        .att('w:w', docxDocumentInstance.width)
        .att('w:h', docxDocumentInstance.height)
        .att('w:orient', docxDocumentInstance.orientation || 'portrait')
        .up();

      // Register this section with the DocxDocument for later reference
      const sectionIndex = parseInt(
        vNode.properties.attributes['data-section-index'] || docxDocumentInstance.currentSectionId,
        10
      );
      docxDocumentInstance.sections.push({
        index: sectionIndex,
        headerType,
        footerType,
        margins,
      });
      // eslint-disable-next-line no-param-reassign
      docxDocumentInstance.currentSectionId = sectionIndex + 1;

      // Complete the section break fragment
      sectionBreakPara.up().up();

      // Now process the data-section div's children
      // eslint-disable-next-line no-restricted-syntax
      for (const child of vNode.children || []) {
        await findXMLEquivalent(docxDocumentInstance, child, xmlFragment);
      }

      xmlFragment.import(sectionBreakPara.up().up());
      return;
    }
    // Regular page break (not a section break)
    const paragraphFragment = fragment({ namespaceAlias: { w: namespaces.w } })
      .ele('w:p')
      .ele('w:r')
      .ele('w:br')
      .att('w:type', 'page')
      .up()
      .up()
      .up();
    xmlFragment.import(paragraphFragment);
    return;
  }

  // Helper function to check if a node has direct text content
  const hasDirectTextContent = (node) => {
    if (!node || !vNodeHasChildren(node)) return false;
    return node.children.some((child) => isVText(child) && child.text.trim());
  };

  // Helper function to check if a node has block elements as children
  const hasBlockChildren = (node) => {
    if (!node || !vNodeHasChildren(node)) return false;
    return node.children.some(
      (child) =>
        isVNode(child) &&
        [
          'div',
          'p',
          'h1',
          'h2',
          'h3',
          'h4',
          'h5',
          'h6',
          'ul',
          'ol',
          'li',
          'blockquote',
          'table',
          'figure',
        ].includes(child.tagName)
    );
  };
  const spacingAfter = vNode.properties?.attributes?.['data-spacing-after'];
  switch (vNode.tagName) {
    case 'div':
      // First process div's children
      // eslint-disable-next-line no-restricted-syntax
      for (const child of vNode.children || []) {
        await findXMLEquivalent(docxDocumentInstance, child, xmlFragment);
      }

      // Check for data-spacing-after attribute
      if (spacingAfter) {
        const spacingParagraph = fragment({ namespaceAlias: { w: namespaces.w } })
          .ele('@w', 'p')
          .ele('@w', 'pPr')
          .ele('@w', 'spacing')
          .att('@w', 'after', spacingAfter)
          .att('@w', 'before', '0')
          .att('@w', 'line', '240')
          .att('@w', 'lineRule', 'auto')
          .up()
          .up()
          .up();
        xmlFragment.import(spacingParagraph);
      }
      return;

    case 'ul':
    case 'ol':
      await buildList(vNode, docxDocumentInstance, xmlFragment);
      return;

    case 'p':
      // Check for spacing attribute even if empty
      if (spacingAfter) {
        const spacingParagraph = fragment({ namespaceAlias: { w: namespaces.w } })
          .ele('@w', 'p')
          .ele('@w', 'pPr')
          .ele('@w', 'spacing')
          .att('@w', 'after', spacingAfter)
          .att('@w', 'before', '0')
          .att('@w', 'line', '240')
          .att('@w', 'lineRule', 'auto')
          .up()
          .up()
          .up();
        xmlFragment.import(spacingParagraph);
        return;
      }

      // Process it normally
      const paragraphFragment = await xmlBuilder.buildParagraph(vNode, {}, docxDocumentInstance);
      xmlFragment.import(paragraphFragment);
      return;

    case 'h1':
    case 'h2':
    case 'h3':
    case 'h4':
    case 'h5':
    case 'h6':
      const headingFragment = await xmlBuilder.buildParagraph(
        vNode,
        { paragraphStyle: `Heading${vNode.tagName[1]}` },
        docxDocumentInstance
      );
      xmlFragment.import(headingFragment);
      return;

    case 'span':
    case 'strong':
    case 'b':
    case 'em':
    case 'i':
    case 'u':
    case 'ins':
    case 'strike':
    case 'del':
    case 's':
    case 'sub':
    case 'sup':
    case 'mark':
    case 'a':
    case 'blockquote':
    case 'code':
    case 'pre':
      // Create paragraph if:
      // 1. Node has direct text content, OR
      // 2. Node is empty (no children) and is a block element, OR
      // 3. Node has no block children
      if (
        hasDirectTextContent(vNode) ||
        (!vNodeHasChildren(vNode) && ['p', 'div', 'blockquote'].includes(vNode.tagName)) ||
        !hasBlockChildren(vNode)
      ) {
        // eslint-disable-next-line no-shadow
        const paragraphFragment = await xmlBuilder.buildParagraph(vNode, {}, docxDocumentInstance);
        xmlFragment.import(paragraphFragment);

        // Add empty paragraph after block elements with mb-6 class
        if (vNode.properties?.attributes?.class?.includes('mb-6')) {
          const spacerParagraph = await xmlBuilder.buildParagraph(
            null,
            { isSpacerParagraph: true },
            docxDocumentInstance
          );
          xmlFragment.import(spacerParagraph);
        }
      } else if (vNodeHasChildren(vNode)) {
        // If we have block children, process them
        // eslint-disable-next-line no-restricted-syntax
        for (const childVNode of vNode.children) {
          await convertVTreeToXML(docxDocumentInstance, childVNode, xmlFragment);
        }
      }
      return;

    case 'figure':
      if (vNodeHasChildren(vNode)) {
        // eslint-disable-next-line no-restricted-syntax
        for (const childVNode of vNode.children) {
          if (childVNode.tagName === 'table') {
            const tableFragment = await xmlBuilder.buildTable(
              childVNode,
              {
                maximumWidth: docxDocumentInstance.availableDocumentSpace,
                rowCantSplit: docxDocumentInstance.tableRowCantSplit,
              },
              docxDocumentInstance
            );
            xmlFragment.import(tableFragment);
            // Only add empty paragraph if there's more content after the table
            const hasMoreContent = childVNode !== vNode.children[vNode.children.length - 1];
            if (hasMoreContent) {
              const emptyParagraphFragment = await xmlBuilder.buildParagraph(null, {});
              xmlFragment.import(emptyParagraphFragment);
            }
          } else if (childVNode.tagName === 'img') {
            const imageFragment = await buildImage(docxDocumentInstance, childVNode);
            if (imageFragment) {
              xmlFragment.import(imageFragment);
            }
          }
        }
      }
      return;

    case 'table':
      const tableFragment = await xmlBuilder.buildTable(
        vNode,
        {
          maximumWidth: docxDocumentInstance.availableDocumentSpace,
          rowCantSplit: docxDocumentInstance.tableRowCantSplit,
        },
        docxDocumentInstance
      );
      xmlFragment.import(tableFragment);
      return;

    case 'img':
      const imageFragment = await buildImage(docxDocumentInstance, vNode);
      if (imageFragment) {
        xmlFragment.import(imageFragment);
      }
      return;

    case 'br':
      // Create a line break within the current paragraph instead of a new paragraph
      const linebreakFragment = fragment({ namespaceAlias: { w: namespaces.w } })
        .ele('w:r')
        .ele('w:br')
        .up()
        .up();
      xmlFragment.import(linebreakFragment);
      return;

    case 'head':
      return;

    default:
      if (vNodeHasChildren(vNode)) {
        // eslint-disable-next-line no-restricted-syntax
        for (const childVNode of vNode.children) {
          await convertVTreeToXML(docxDocumentInstance, childVNode, xmlFragment);
        }
      }
  }
}

async function renderDocumentFile(docxDocumentInstance) {
  const vTree = convertHTML(docxDocumentInstance.htmlString);

  const xmlFragment = fragment({ namespaceAlias: { w: namespaces.w } });

  const populatedXmlFragment = await convertVTreeToXML(docxDocumentInstance, vTree, xmlFragment);

  return populatedXmlFragment;
}

export default renderDocumentFile;
