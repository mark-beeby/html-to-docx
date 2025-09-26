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
import { buildTableRow, modifiedStyleAttributesBuilder } from './xml-builder';
import namespaces from '../namespaces';
import { imageType, internalRelationship } from '../constants';
import { vNodeHasChildren } from '../utils/vnode';
import { isValidUrl } from '../utils/url';

const convertHTML = HTMLToVDOM({
  VNode,
  VText,
});

const isInlineContent = (child) => {
  if (isVText(child)) return true;
  return !!(
    isVNode(child) &&
    [
      'br',
      'strong',
      'b',
      'em',
      'i',
      'u',
      'span',
      'a',
      'del',
      's',
      'ins',
      'sub',
      'sup',
      'mark',
    ].includes(child.tagName)
  );
};

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

    return xmlBuilder.buildParagraph(
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

// Helper function to clean up whitespace around br tags
const cleanWhitespaceAroundBr = (nodes) => {
  // eslint-disable-next-line no-plusplus
  for (let i = 0; i < nodes.length; i++) {
    if (isVNode(nodes[i]) && nodes[i].tagName === 'br') {
      // Clean whitespace before br
      if (i > 0 && isVText(nodes[i - 1]) && !nodes[i - 1].text.trim()) {
        nodes.splice(i - 1, 1);
        // eslint-disable-next-line no-plusplus
        i--; // Adjust index after removal
      }

      // Clean whitespace after br
      if (i + 1 < nodes.length && isVText(nodes[i + 1]) && !nodes[i + 1].text.trim()) {
        nodes.splice(i + 1, 1);
      }
    }
  }
  return nodes;
};
const processChildrenInGroups = async (
  docxDocumentInstance,
  xmlFragment,
  children,
  parentVNode
) => {
  let currentParagraphContent = [];

  const paragraphOptions = {};

  if (
    parentVNode &&
    parentVNode.properties?.style?.['text-align'] &&
    ['left', 'right', 'center', 'justify'].includes(
      parentVNode.properties.style['text-align'].toLowerCase()
    )
  ) {
    paragraphOptions.textAlign = parentVNode.properties.style['text-align'].toLowerCase();
  }

  const finishCurrentParagraph = async () => {
    if (currentParagraphContent.length > 0) {
      const paragraphVNode = new VNode(
        'p',
        parentVNode?.properties || null,
        currentParagraphContent
      );
      const paragraphFragment = await xmlBuilder.buildParagraph(
        paragraphVNode,
        paragraphOptions,
        docxDocumentInstance
      );
      xmlFragment.import(paragraphFragment);
      currentParagraphContent = [];
    }
  };

  const createEmptyParagraph = async () => {
    const emptyParagraphFragment = await xmlBuilder.buildParagraph(
      null,
      { ...paragraphOptions, beforeSpacing: 0, afterSpacing: 0, lineSpacing: 240 },
      docxDocumentInstance
    );
    xmlFragment.import(emptyParagraphFragment);
  };

  const cleanedChildren = cleanWhitespaceAroundBr([...children]);

  // eslint-disable-next-line no-plusplus
  for (let i = 0; i < cleanedChildren.length; i++) {
    const child = cleanedChildren[i];

    if (isVNode(child) && child.tagName === 'br') {
      // Count consecutive <br> tags starting from current position
      let consecutiveBrCount = 1;
      let j = i + 1;

      while (
        j < cleanedChildren.length &&
        isVNode(cleanedChildren[j]) &&
        cleanedChildren[j].tagName === 'br'
      ) {
        // eslint-disable-next-line no-plusplus
        consecutiveBrCount++;
        // eslint-disable-next-line no-plusplus
        j++;
      }

      if (consecutiveBrCount >= 2) {
        await finishCurrentParagraph();

        const emptyLinesToCreate = Math.ceil(consecutiveBrCount / 2);

        // eslint-disable-next-line no-plusplus
        for (let k = 0; k < emptyLinesToCreate; k++) {
          await createEmptyParagraph();
        }

        i = j - 1;
      } else if (currentParagraphContent.length > 0) {
        // Single <br> tag
        // Content exists - check if there's content after this <br>
        let hasContentAfter = false;
        // eslint-disable-next-line no-plusplus
        for (let k = i + 1; k < cleanedChildren.length; k++) {
          if (
            isInlineContent(cleanedChildren[k]) &&
            !(isVText(cleanedChildren[k]) && !cleanedChildren[k].text.trim())
          ) {
            hasContentAfter = true;
            break;
          } else if (!isInlineContent(cleanedChildren[k])) {
            break;
          }
        }

        if (hasContentAfter) {
          // Single <br> with content after: add as line break within paragraph
          currentParagraphContent.push(child);
        } else {
          // Single <br> at end: finish paragraph (br just ends the line)
          await finishCurrentParagraph();
        }
      } else {
        // Standalone <br>: check if there's content after
        let hasContentAfter = false;
        // eslint-disable-next-line no-plusplus
        for (let k = i + 1; k < cleanedChildren.length; k++) {
          if (
            isInlineContent(cleanedChildren[k]) &&
            !(isVText(cleanedChildren[k]) && !cleanedChildren[k].text.trim())
          ) {
            hasContentAfter = true;
            break;
          } else if (!isInlineContent(cleanedChildren[k])) {
            break;
          }
        }

        if (hasContentAfter) {
          await createEmptyParagraph();
        }
      }
    } else if (isInlineContent(child)) {
      if (isVText(child) && !child.text.trim()) {
        // eslint-disable-next-line no-continue
        continue;
      }
      currentParagraphContent.push(child);
    } else {
      await finishCurrentParagraph();
      // eslint-disable-next-line no-use-before-define
      await findXMLEquivalent(docxDocumentInstance, child, xmlFragment);
    }
  }

  await finishCurrentParagraph();
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

      // Background data
      const backgroundUrl = vNode.properties.attributes['data-background-url'];
      const backgroundSize = vNode.properties.attributes['data-background-size'];
      const backgroundPosition = vNode.properties.attributes['data-background-position'];
      const backgroundRepeat = vNode.properties.attributes['data-background-repeat'];

      // Store section information
      let margins = null;
      if (vNode.properties.attributes['data-margins']) {
        const marginStr = vNode.properties.attributes['data-margins'].replaceAll('&quot;', '"');
        try {
          margins = JSON.parse(marginStr);
        } catch (e) {
          console.error('Error parsing margins JSON:', e);
        }
      }

      docxDocumentInstance.sections.push({
        index: sectionIndex,
        headerType,
        footerType,
        backgroundUrl,
        backgroundSize,
        backgroundPosition,
        backgroundRepeat,
        margins,
      });
      // eslint-disable-next-line no-param-reassign
      docxDocumentInstance.currentSectionId = sectionIndex + 1;
    }

    // For data-section divs, we'll create a section break with its own properties
    // For data-section divs, we'll create a section break with its own properties
    if (isProfilePage) {
      // Extract section data
      let headerType = vNode.properties.attributes['data-header-type'] || 'default';
      const footerType = vNode.properties.attributes['data-footer-type'] || 'default';
      const backgroundUrl = vNode.properties.attributes['data-background-url'];
      const backgroundSize = vNode.properties.attributes['data-background-size'];
      const backgroundPosition = vNode.properties.attributes['data-background-position'];
      const backgroundRepeat = vNode.properties.attributes['data-background-repeat'];

      // Only generate header if headerType is not 'none'
      if (headerType !== 'none') {
        // Get the existing header content from the document-level header
        let existingHeaderVTree = null;
        let existingHeaderConfig = {};

        if (headerType === 'default' && docxDocumentInstance.defaultHeaderVTree) {
          existingHeaderVTree = docxDocumentInstance.defaultHeaderVTree;
        }

        if (headerType === 'default' && docxDocumentInstance.defaultHeaderConfig) {
          // Make a deep copy to avoid modifying the original
          existingHeaderConfig = JSON.parse(
            JSON.stringify(docxDocumentInstance.defaultHeaderConfig)
          );
        }

        // Create modified header config that includes page background if present
        const modifiedHeaderConfig = { ...existingHeaderConfig };

        if (backgroundUrl) {
          modifiedHeaderConfig.pageBackground = {
            url: backgroundUrl.replaceAll('&amp;', '&'),
            size: backgroundSize,
            position: backgroundPosition,
            repeat: backgroundRepeat,
          };
        }

        // Create unique header type name based on whether it has background
        const uniqueHeaderTypeName = backgroundUrl
          ? `${headerType}_bg_${docxDocumentInstance.hashBackgroundInfo({
              backgroundUrl,
              backgroundSize,
              backgroundPosition,
              backgroundRepeat,
            })}`
          : headerType;

        // Only generate if we don't already have this header variant
        if (!docxDocumentInstance.headerObjects[uniqueHeaderTypeName]) {
          // Use your existing generateHeaderXML method
          const headerResult = await docxDocumentInstance.generateHeaderXML(
            existingHeaderVTree,
            modifiedHeaderConfig,
            uniqueHeaderTypeName
          );

          // Create header file and relationship
          const headerFileName = `header${headerResult.headerId}.xml`;
          docxDocumentInstance.zip
            .folder('word')
            .file(headerFileName, headerResult.headerXML.toString({ prettyPrint: true }), {
              createFolders: false,
            });

          const headerRelationshipId = docxDocumentInstance.createDocumentRelationships(
            docxDocumentInstance.relationshipFilename,
            'header',
            headerFileName,
            'Internal'
          );

          // Store header
          // eslint-disable-next-line no-param-reassign
          docxDocumentInstance.headerObjects[uniqueHeaderTypeName] = {
            headerId: headerResult.headerId,
            relationshipId: headerRelationshipId,
            height: headerResult.headerHeight || 0,
          };
        }

        // Update headerType to use the unique key
        headerType = uniqueHeaderTypeName;
      } else if (backgroundUrl) {
        // headerType is 'none' - but we still might need a background-only header
        const noContentHeaderResult = await docxDocumentInstance.generateSectionHeader({
          headerType,
          backgroundUrl: backgroundUrl.replaceAll('&amp;', '&'),
          backgroundSize,
          backgroundPosition,
          backgroundRepeat,
        });

        // Create header file and relationship for background-only header
        const headerFileName = `header${noContentHeaderResult.headerId}.xml`;
        docxDocumentInstance.zip
          .folder('word')
          .file(headerFileName, noContentHeaderResult.headerXML.toString({ prettyPrint: true }), {
            createFolders: false,
          });

        const headerRelationshipId = docxDocumentInstance.createDocumentRelationships(
          docxDocumentInstance.relationshipFilename,
          'header',
          headerFileName,
          'Internal'
        );

        // Store header with background-only key
        const backgroundOnlyHeaderKey = `${headerType}_bg_only`;
        // eslint-disable-next-line no-param-reassign
        docxDocumentInstance.headerObjects[backgroundOnlyHeaderKey] = {
          headerId: noContentHeaderResult.headerId,
          relationshipId: headerRelationshipId,
          height: noContentHeaderResult.headerHeight || 0,
          variantName: noContentHeaderResult.variantName,
        };

        // Update headerType to use the background-only key
        headerType = backgroundOnlyHeaderKey;
        // If no background URL and headerType is 'none', we don't generate any header
      }

      // Create a section break with its own properties (SAME STRUCTURE AS BEFORE)
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

      // Calculate heights (same as before)
      const showHeader = headerType !== 'none';
      const showFooter = footerType !== 'none';
      let headerHeight = showHeader ? docxDocumentInstance.margins.header : 0;
      let footerHeight = showFooter ? docxDocumentInstance.margins.footer : 0;

      if (docxDocumentInstance.headerObjects[headerType] && showHeader) {
        headerHeight = docxDocumentInstance.headerObjects[headerType].height ?? headerHeight;
      }
      if (docxDocumentInstance.footerObjects[footerType] && showFooter) {
        footerHeight = docxDocumentInstance.footerObjects[footerType].height ?? footerHeight;
      }

      const hasCustomMargins = vNode.properties.attributes['data-margins'];
      let margins = null;

      if (hasCustomMargins) {
        try {
          const marginStr = vNode.properties.attributes['data-margins'].replaceAll('&quot;', '"');
          margins = JSON.parse(marginStr);

          const finalTopMargin = Math.round(
            (margins.top ?? docxDocumentInstance.margins.top) + headerHeight
          );
          const finalBottomMargin = Math.round(
            (margins.bottom ?? docxDocumentInstance.margins.bottom) + footerHeight
          );

          sectionBreakPara
            .ele('w:pgMar')
            .att('w:top', finalTopMargin)
            .att('w:right', Math.round(margins.right ?? docxDocumentInstance.margins.right))
            .att('w:bottom', finalBottomMargin)
            .att('w:left', Math.round(margins.left ?? docxDocumentInstance.margins.left))
            .att('w:header', Math.round(docxDocumentInstance.margins.header))
            .att('w:footer', Math.round(docxDocumentInstance.margins.footer))
            .up();
        } catch (e) {
          console.error('Error parsing margins JSON:', e);
          // Use default page margins if parsing fails
          const finalTopMargin = Math.round(docxDocumentInstance.margins.top + headerHeight);
          const finalBottomMargin = Math.round(docxDocumentInstance.margins.bottom + footerHeight);

          sectionBreakPara
            .ele('w:pgMar')
            .att('w:top', finalTopMargin)
            .att('w:right', Math.round(docxDocumentInstance.margins.right))
            .att('w:bottom', finalBottomMargin)
            .att('w:left', Math.round(docxDocumentInstance.margins.left))
            .att('w:header', Math.round(docxDocumentInstance.margins.header))
            .att('w:footer', Math.round(docxDocumentInstance.margins.footer))
            .up();
        }
      } else {
        // Use default page margins if no custom margins are provided
        const finalTopMargin = Math.round(docxDocumentInstance.margins.top + headerHeight);
        const finalBottomMargin = Math.round(docxDocumentInstance.margins.bottom + footerHeight);
        sectionBreakPara
          .ele('w:pgMar')
          .att('w:top', finalTopMargin)
          .att('w:right', Math.round(docxDocumentInstance.margins.right))
          .att('w:bottom', finalBottomMargin)
          .att('w:left', Math.round(docxDocumentInstance.margins.left))
          .att('w:header', Math.round(docxDocumentInstance.margins.header))
          .att('w:footer', Math.round(docxDocumentInstance.margins.footer))
          .up();
      }
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
      // Check if div has direct text content and other content that needs to be preserved in order
      if (hasDirectTextContent(vNode)) {
        // Create a new virtual node representing just the div with its text content
        // but maintain the original order of all children
        // Process each child in order, handling text nodes specially
        // eslint-disable-next-line no-restricted-syntax
        for (const child of vNode.children || []) {
          if (isVText(child) && child.text.trim()) {
            // For text nodes, create a paragraph to hold them
            const paragraphOptions = {};

            // Preserve text-align from the parent div if it exists
            if (
              vNode.properties?.style?.['text-align'] &&
              ['left', 'right', 'center', 'justify'].includes(
                vNode.properties.style['text-align'].toLowerCase()
              )
            ) {
              paragraphOptions.textAlign = vNode.properties.style['text-align'].toLowerCase();
            }

            // Create a new virtual node that looks like a paragraph but with the div's properties
            // and just this text node as content
            const textParagraphVNode = new VNode('p', vNode.properties, [child]);

            // Build paragraph for just this text node
            const textParagraphFragment = await xmlBuilder.buildParagraph(
              textParagraphVNode,
              paragraphOptions,
              docxDocumentInstance
            );

            if (textParagraphFragment) {
              xmlFragment.import(textParagraphFragment);
            }
          } else if (isVNode(child)) {
            // Process non-text children normally
            await findXMLEquivalent(docxDocumentInstance, child, xmlFragment);
          }
        }
      }

      if (vNodeHasChildren(vNode)) {
        await processChildrenInGroups(docxDocumentInstance, xmlFragment, vNode.children, vNode);
      }

      // Check for data-spacing-after attribute
      if (spacingAfter) {
        const spacingParagraph = fragment({ namespaceAlias: { w: namespaces.w } })
          .ele('@w', 'p')
          .ele('@w', 'pPr')
          .ele('@w', 'spacing')
          .att('@w', 'after', spacingAfter)
          .att('@w', 'before', spacingAfter)
          .att('@w', 'line', spacingAfter)
          .att('@w', 'lineRule', 'exactly')
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
          .att('@w', 'before', spacingAfter)
          .att('@w', 'line', spacingAfter)
          .att('@w', 'lineRule', 'exactly')
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
      const headingAttributes = modifiedStyleAttributesBuilder(
        docxDocumentInstance,
        vNode,
        {},
        { isParagraph: true }
      );
      headingAttributes.paragraphStyle = `Heading${vNode.tagName[1]}`;
      const headingFragment = await xmlBuilder.buildParagraph(
        vNode,
        headingAttributes,
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
      // Check if this inline element has display: block style
      const hasDisplayBlock = vNode.properties?.style?.display === 'block';

      if (hasDisplayBlock) {
        if (vNodeHasChildren(vNode)) {
          await processChildrenInGroups(docxDocumentInstance, xmlFragment, vNode.children, vNode);
        }
        return;
      }

      // Regular handling for inline elements (unchanged)
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
