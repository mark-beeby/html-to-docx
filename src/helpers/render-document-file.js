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
// eslint-disable-next-line import/no-cycle
import * as xmlBuilder from './xml-builder';
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
  const listElements = [];

  let vNodeObjects = [
    {
      node: vNode,
      level: 0,
      type: vNode.tagName,
      numberingId: docxDocumentInstance.createNumbering(vNode.tagName, vNode.properties),
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
            numberingId: docxDocumentInstance.createNumbering(
              childVNode.tagName,
              childVNode.properties
            ),
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

  return listElements;
};

export async function convertVTreeToXML(docxDocumentInstance, vTree, xmlFragment) {
  if (!vTree) return xmlFragment;
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
    // Only create a paragraph for text nodes that have actual content
    if (vTree.text.trim()) {
      const paragraphFragment = await xmlBuilder.buildParagraph(vTree, {}, docxDocumentInstance);
      xmlFragment.import(paragraphFragment);
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
    vNode.properties.attributes['data-docx-column']
  ) {
    return; // Handled in parent function to convert columns
  }

  if (
    vNode.tagName === 'div' &&
    (vNode.properties.attributes.class === 'page-break' ||
      (vNode.properties.style && vNode.properties.style['page-break-after']))
  ) {
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

  // Helper function to check if a node has the mb-6 class
  const hasMb6Class = (node) => {
    if (!node || !node.properties || !node.properties.className) return false;
    return node.properties.className.includes('mb-6');
  };

  switch (vNode.tagName) {
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
    case 'p':
    case 'div':
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
        const paragraphFragment = await xmlBuilder.buildParagraph(vNode, {}, docxDocumentInstance);
        xmlFragment.import(paragraphFragment);

        // Add empty paragraph after block elements with mb-6 class
        if (hasMb6Class(vNode)) {
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

    case 'ul':
    case 'ol':
      await buildList(vNode, docxDocumentInstance, xmlFragment);
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
        .ele('@w', 'r')
        .ele('@w', 'br')
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
