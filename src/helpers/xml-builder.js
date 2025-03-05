/* eslint-disable no-await-in-loop */
/* eslint-disable radix */
/* eslint-disable no-param-reassign */
/* eslint-disable no-case-declarations */
/* eslint-disable no-plusplus */
/* eslint-disable no-else-return */
import { fragment } from 'xmlbuilder2';
import isVNode from 'virtual-dom/vnode/is-vnode';
import isVText from 'virtual-dom/vnode/is-vtext';
import colorNames from 'color-name';
import { cloneDeep } from 'lodash';
import imageToBase64 from 'image-to-base64';
import mimeTypes from 'mime-types';
import sizeOf from 'image-size';

import VirtualText from 'virtual-dom/vnode/vtext';
import namespaces from '../namespaces';
import {
  hex3Regex,
  hex3ToHex,
  hexRegex,
  hslRegex,
  hslToHex,
  rgbRegex,
  rgbToHex,
} from '../utils/color-conversion';
import {
  HIPToTWIP,
  percentageRegex,
  pixelRegex,
  pixelToEMU,
  pixelToHIP,
  pixelToTWIP,
  pointRegex,
  pointToHIP,
  pointToTWIP,
  TWIPToEMU,
} from '../utils/unit-conversion';
// eslint-disable-next-line import/no-cycle
import { convertVTreeToXML } from './render-document-file';
import {
  colorlessColors,
  defaultFont,
  hyperlinkType,
  imageType,
  internalRelationship,
  paragraphBordersObject,
  verticalAlignValues,
} from '../constants';
import { vNodeHasChildren } from '../utils/vnode';
import { isValidUrl } from '../utils/url';

const valignMapping = {
  top: 'top',
  middle: 'center',
  bottom: 'bottom',
  baseline: 'top', // No direct equivalent, typically treated as 'top'
};

const cssVerticalAlignMapping = {
  baseline: 'bottom', // No direct equivalent, typically treated as 'bottom'
  sub: 'bottom', // No direct equivalent, typically treated as 'bottom'
  super: 'top', // No direct equivalent, typically treated as 'top'
  'text-top': 'top',
  'text-bottom': 'bottom',
  middle: 'center',
  top: 'top',
  bottom: 'bottom',
  center: 'center',
  inherit: 'center', // Default to center when inheriting
};

// eslint-disable-next-line consistent-return
const fixupColorCode = (colorCodeString) => {
  if (Object.prototype.hasOwnProperty.call(colorNames, colorCodeString.toLowerCase())) {
    const [red, green, blue] = colorNames[colorCodeString.toLowerCase()];

    return rgbToHex(red, green, blue);
  } else if (rgbRegex.test(colorCodeString)) {
    const matchedParts = colorCodeString.match(rgbRegex);
    const red = matchedParts[1];
    const green = matchedParts[2];
    const blue = matchedParts[3];

    return rgbToHex(red, green, blue);
  } else if (hslRegex.test(colorCodeString)) {
    const matchedParts = colorCodeString.match(hslRegex);
    const hue = matchedParts[1];
    const saturation = matchedParts[2];
    const luminosity = matchedParts[3];

    return hslToHex(hue, saturation, luminosity);
  } else if (hexRegex.test(colorCodeString)) {
    const matchedParts = colorCodeString.match(hexRegex);

    return matchedParts[1];
  } else if (hex3Regex.test(colorCodeString)) {
    const matchedParts = colorCodeString.match(hex3Regex);
    const red = matchedParts[1];
    const green = matchedParts[2];
    const blue = matchedParts[3];

    return hex3ToHex(red, green, blue);
  } else {
    return '000000';
  }
};

const recursiveRunOrHyperlink = async (vNode, attributes, docxDocumentInstance) => {
  let runFragments = [];
  // eslint-disable-next-line no-use-before-define
  const combinedAttributes = modifiedStyleAttributesBuilder(
    docxDocumentInstance,
    vNode,
    attributes
  );
  // eslint-disable-next-line no-restricted-syntax
  for (const childVNode of vNode.children) {
    if (isVNode(childVNode)) {
      // eslint-disable-next-line no-use-before-define
      const combinedChildAttributes = modifiedStyleAttributesBuilder(
        docxDocumentInstance,
        childVNode,
        combinedAttributes
      );

      // eslint-disable-next-line no-shadow
      let fragment = null;

      if (childVNode.tagName === 'a') {
        // eslint-disable-next-line no-use-before-define
        fragment = await buildRunOrHyperLink(
          childVNode,
          combinedChildAttributes,
          docxDocumentInstance
        );
      } else if (childVNode.tagName === 'span') {
        fragment = await recursiveRunOrHyperlink(
          childVNode,
          combinedChildAttributes,
          docxDocumentInstance
        );
      } else {
        // eslint-disable-next-line no-use-before-define
        fragment = await buildRunOrRuns(childVNode, combinedChildAttributes, docxDocumentInstance);
      }

      if (fragment) {
        runFragments = runFragments.concat(Array.isArray(fragment) ? fragment : [fragment]);
      }
    } else if (isVText(childVNode)) {
      // Ensure styles are captured from parent attributes
      // eslint-disable-next-line no-use-before-define
      const textFragment = buildTextElement(childVNode.text);
      const styledRunFragment = fragment({ namespaceAlias: { w: namespaces.w } })
        .ele('@w', 'r')
        // eslint-disable-next-line no-use-before-define
        .import(buildRunProperties(combinedAttributes)) // Use "attributes" incorporating parent styles
        .import(textFragment);
      styledRunFragment.up();
      runFragments.push(styledRunFragment);
    }
  }

  return runFragments;
};

const buildRunFontFragment = (fontName = defaultFont) =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'rFonts')
    .att('@w', 'ascii', fontName)
    .att('@w', 'hAnsi', fontName)
    .up();

const buildRunStyleFragment = (type = 'Hyperlink') =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'rStyle')
    .att('@w', 'val', type)
    .up();

const buildColor = (colorCode) =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'color')
    .att('@w', 'val', colorCode)
    .up();

const buildFontSize = (fontSize) =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'sz')
    .att('@w', 'val', fontSize)
    .up();

const buildShading = (colorCode) =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'shd')
    .att('@w', 'val', 'clear')
    .att('@w', 'fill', colorCode)
    .up();

const buildHighlight = (color = 'yellow') =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'highlight')
    .att('@w', 'val', color)
    .up();

const buildVertAlign = (type = 'baseline') =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'vertAlign')
    .att('@w', 'val', type)
    .up();

const buildStrike = () =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'strike')
    .att('@w', 'val', true)
    .up();

const buildBold = () =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'b')
    .up();

const buildItalics = () =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'i')
    .up();

const buildUnderline = (type = 'single') =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'u')
    .att('@w', 'val', type)
    .up();

const buildLineBreak = (type = 'textWrapping') =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'br')
    .att('@w', 'type', type)
    .up();

const buildBorder = (
  borderSide = 'top',
  borderSize = 0,
  borderSpacing = 0,
  borderColor = fixupColorCode('black'),
  borderStroke = 'single'
) =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', borderSide)
    .att('@w', 'val', borderStroke)
    .att('@w', 'sz', borderSize)
    .att('@w', 'space', borderSpacing)
    .att('@w', 'color', borderColor)
    .up();

const buildTextElement = (text) =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 't')
    .att('@xml', 'space', 'preserve')
    .txt(text)
    .up();

// eslint-disable-next-line consistent-return
const fixupLineHeight = (lineHeight, fontSize) => {
  // FIXME: If line height is anything other than a number
  // eslint-disable-next-line no-restricted-globals
  if (!isNaN(lineHeight)) {
    if (fontSize) {
      const actualLineHeight = +lineHeight * fontSize;

      return HIPToTWIP(actualLineHeight);
    } else {
      // 240 TWIP or 12 point is default line height
      return +lineHeight * 240;
    }
  } else {
    // 240 TWIP or 12 point is default line height
    return 240;
  }
};

// eslint-disable-next-line consistent-return
const fixupFontSize = (fontSizeString) => {
  if (pointRegex.test(fontSizeString)) {
    const matchedParts = fontSizeString.match(pointRegex);
    // convert point to half point
    return pointToHIP(matchedParts[1]);
  } else if (pixelRegex.test(fontSizeString)) {
    const matchedParts = fontSizeString.match(pixelRegex);
    // convert pixels to half point
    return pixelToHIP(matchedParts[1]);
  }
};

// eslint-disable-next-line consistent-return
const fixupMargin = (marginString) => {
  if (pointRegex.test(marginString)) {
    const matchedParts = marginString.match(pointRegex);
    // convert point to half point
    return pointToTWIP(matchedParts[1]);
  } else if (pixelRegex.test(marginString)) {
    const matchedParts = marginString.match(pixelRegex);
    // convert pixels to half point
    return pixelToTWIP(matchedParts[1]);
  }
};

const modifiedStyleAttributesBuilder = (docxDocumentInstance, vNode, attributes, options) => {
  const modifiedAttributes = { ...attributes };

  if (isVNode(vNode) && vNode.properties && vNode.properties.style) {
    const { style } = vNode.properties;

    if (style['font-family']) {
      modifiedAttributes.font = docxDocumentInstance.createFont(style['font-family']);
    }
    if (style['text-decoration'] && style['text-decoration'] === 'underline') {
      modifiedAttributes.u = true;
    }
    if (style['font-style'] && style['font-style'] === 'italic') {
      modifiedAttributes.i = true;
    }
    if (style['font-size']) {
      modifiedAttributes.fontSize = fixupFontSize(style['font-size']);
    }
    if (style.color && !colorlessColors.includes(style.color)) {
      modifiedAttributes.color = fixupColorCode(style.color);
    }
    if (style['background-color'] && !colorlessColors.includes(style['background-color'])) {
      modifiedAttributes.backgroundColor = fixupColorCode(style['background-color']);
    }
    if (style['vertical-align'] && verticalAlignValues.includes(style['vertical-align'])) {
      modifiedAttributes.verticalAlign = style['vertical-align'];
    }
    if (
      style['text-align'] &&
      ['left', 'right', 'center', 'justify'].includes(style['text-align'])
    ) {
      modifiedAttributes.textAlign = style['text-align'];
    }
    if (style['font-weight']) {
      const weight = style['font-weight'];
      if (weight === 'bold' || parseInt(weight, 10) >= 700) {
        modifiedAttributes.strong = true;
      }
    }
    if (style['line-height']) {
      modifiedAttributes.lineHeight = fixupLineHeight(
        style['line-height'],
        modifiedAttributes.fontSize
      );
    }
    if (style['margin-top']) {
      modifiedAttributes.marginTop = fixupMargin(style['margin-top']);
    }
    if (style['margin-bottom']) {
      modifiedAttributes.marginBottom = fixupMargin(style['margin-bottom']);
    }
    if (
      style['margin-left'] ||
      style['margin-right'] ||
      style['padding-left'] ||
      style['padding-right']
    ) {
      const leftMargin = style['margin-left']
        ? fixupMargin(style['margin-left'])
        : fixupMargin(style['padding-left']);
      const rightMargin = style['margin-right']
        ? fixupMargin(style['margin-right'])
        : fixupMargin(style['padding-right']);
      const indentation = {};
      if (leftMargin) {
        indentation.left = leftMargin;
      }
      if (rightMargin) {
        indentation.right = rightMargin;
      }
      if (leftMargin || rightMargin) {
        modifiedAttributes.indentation = indentation;
      }
    }
    if (style.display) {
      modifiedAttributes.display = style.display;
    }
    if (style.width) {
      modifiedAttributes.width = style.width;
    }
  }

  // Handle classes
  if (isVNode(vNode) && vNode.properties && vNode.properties.className) {
    const classes = vNode.properties.className.split(' ');
    classes.forEach((className) => {
      if (className.startsWith('mb-')) {
        const value = parseInt(className.slice(3), 10) * 4;
        modifiedAttributes.marginBottom = pixelToTWIP(value);
      }
      if (className.startsWith('mt-')) {
        const value = parseInt(className.slice(3), 10) * 4;
        modifiedAttributes.marginTop = pixelToTWIP(value);
      }
    });
  }

  // paragraph only
  if (options && options.isParagraph) {
    if (isVNode(vNode) && vNode.tagName === 'blockquote') {
      modifiedAttributes.indentation = { left: 284 };
      modifiedAttributes.textAlign = 'justify';
    } else if (isVNode(vNode) && vNode.tagName === 'code') {
      modifiedAttributes.highlightColor = 'lightGray';
    } else if (isVNode(vNode) && vNode.tagName === 'pre') {
      modifiedAttributes.font = 'Courier';
    }
  }

  return modifiedAttributes;
};

// html tag to formatting function
// options are passed to the formatting function if needed
const buildFormatting = (htmlTag, options) => {
  switch (htmlTag) {
    case 'strong':
    case 'b':
      return buildBold();
    case 'em':
    case 'i':
      return buildItalics();
    case 'ins':
    case 'u':
      return buildUnderline();
    case 'strike':
    case 'del':
    case 's':
      return buildStrike();
    case 'sub':
      return buildVertAlign('subscript');
    case 'sup':
      return buildVertAlign('superscript');
    case 'mark':
      return buildHighlight();
    case 'code':
      return buildHighlight('lightGray');
    case 'highlightColor':
      return buildHighlight(options && options.color ? options.color : 'lightGray');
    case 'font':
      return buildRunFontFragment(options.font);
    case 'pre':
      return buildRunFontFragment('Courier');
    case 'color':
      return buildColor(options && options.color ? options.color : 'black');
    case 'backgroundColor':
      return buildShading(options && options.color ? options.color : 'black');
    case 'fontSize':
      // does this need a unit of measure?
      return buildFontSize(options && options.fontSize ? options.fontSize : 10);
    case 'hyperlink':
      return buildRunStyleFragment('Hyperlink');
  }

  return null;
};

const buildRunProperties = (attributes) => {
  const runPropertiesFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'rPr');
  if (attributes && attributes.constructor === Object) {
    Object.keys(attributes).forEach((key) => {
      const options = {};
      if (key === 'color' || key === 'backgroundColor' || key === 'highlightColor') {
        options.color = attributes[key];
      }

      if (key === 'fontSize' || key === 'font') {
        options[key] = attributes[key];
      }

      // Handle new attributes
      if (key === 'textAlign' || key === 'lineHeight') {
        // These are paragraph-level properties, so we don't need to handle them here
        return;
      }

      const formattingFragment = buildFormatting(key, options);
      if (formattingFragment) {
        runPropertiesFragment.import(formattingFragment);
      }
    });
  }
  runPropertiesFragment.up();

  return runPropertiesFragment;
};

const buildRun = async (vNode, attributes, docxDocumentInstance) => {
  const runFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'r');
  const runPropertiesFragment = buildRunProperties(cloneDeep(attributes));

  // case where we have recursive spans representing font changes
  if (isVNode(vNode) && vNode.tagName === 'span') {
    // eslint-disable-next-line no-use-before-define
    return buildRunOrRuns(vNode, attributes, docxDocumentInstance);
  }

  if (
    isVNode(vNode) &&
    [
      'strong',
      'b',
      'em',
      'i',
      'u',
      'ins',
      'strike',
      'del',
      's',
      'sub',
      'sup',
      'mark',
      'blockquote',
      'code',
      'pre',
    ].includes(vNode.tagName)
  ) {
    const runFragmentsArray = [];

    let vNodes = [vNode];
    // create temp run fragments to split the paragraph into different runs
    let tempAttributes = cloneDeep(attributes);
    let tempRunFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'r');
    while (vNodes.length) {
      const tempVNode = vNodes.shift();
      if (isVText(tempVNode)) {
        const textFragment = buildTextElement(tempVNode.text);
        const tempRunPropertiesFragment = buildRunProperties({ ...attributes, ...tempAttributes });
        tempRunFragment.import(tempRunPropertiesFragment);
        tempRunFragment.import(textFragment);
        runFragmentsArray.push(tempRunFragment);

        // re initialize temp run fragments with new fragment
        tempAttributes = cloneDeep(attributes);
        tempRunFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'r');
      } else if (isVNode(tempVNode)) {
        if (tempVNode.tagName === 'br') {
          const lineBreakFragment = buildLineBreak('textWrapping');
          tempRunFragment.import(lineBreakFragment);
          runFragmentsArray.push(tempRunFragment);

          // re initialize temp run fragments with new fragment
          tempAttributes = cloneDeep(attributes);
          tempRunFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'r');

          // If the next node is a text node, trim its leading whitespace
          if (vNodes.length > 0 && isVText(vNodes[0])) {
            vNodes[0].text = vNodes[0].text.replace(/^\s+/, '');
            // If the text is now empty, skip it
            if (!vNodes[0].text) {
              vNodes.shift();
            }
          }
          // eslint-disable-next-line no-continue
          continue;
        }
        if (
          [
            'strong',
            'b',
            'em',
            'i',
            'u',
            'ins',
            'strike',
            'del',
            's',
            'sub',
            'sup',
            'mark',
            'code',
            'pre',
          ].includes(tempVNode.tagName)
        ) {
          tempAttributes = {};
          switch (tempVNode.tagName) {
            case 'strong':
            case 'b':
              tempAttributes.strong = true;
              break;
            case 'i':
            case 'em':
              tempAttributes.i = true;
              break;
            case 'u':
              tempAttributes.u = true;
              break;
            case 'sub':
              tempAttributes.sub = true;
              break;
            case 'sup':
              tempAttributes.sup = true;
              break;
          }
          const formattingFragment = buildFormatting(tempVNode);

          if (formattingFragment) {
            runPropertiesFragment.import(formattingFragment);
          }
          // go a layer deeper if there is a span somewhere in the children
        } else if (tempVNode.tagName === 'span') {
          // eslint-disable-next-line no-use-before-define
          const spanFragment = await buildRunOrRuns(
            tempVNode,
            { ...attributes, ...tempAttributes },
            docxDocumentInstance
          );

          // if spanFragment is an array, we need to add each fragment to the runFragmentsArray. If the fragment is an array, perform a depth first search on the array to add each fragment to the runFragmentsArray
          if (Array.isArray(spanFragment)) {
            spanFragment.flat(Infinity);
            runFragmentsArray.push(...spanFragment);
          } else {
            runFragmentsArray.push(spanFragment);
          }

          // do not slice and concat children since this is already accounted for in the buildRunOrRuns function
          // eslint-disable-next-line no-continue
          continue;
        }
      }

      if (tempVNode.children && tempVNode.children.length) {
        if (tempVNode.children.length > 1) {
          attributes = { ...attributes, ...tempAttributes };
        }
        vNodes = tempVNode.children.slice().concat(vNodes);
      }
    }
    if (runFragmentsArray.length) {
      return runFragmentsArray;
    }
  }

  runFragment.import(runPropertiesFragment);

  if (isVText(vNode) && vNode.text.trim().length > 0) {
    const textFragment = buildTextElement(vNode.text);
    runFragment.import(textFragment);
  } else if (attributes && attributes.type === 'picture') {
    let response = null;

    const base64Uri = decodeURIComponent(vNode.properties.src);
    if (base64Uri) {
      response = docxDocumentInstance.createMediaFile(base64Uri);
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

      attributes.inlineOrAnchored = true;
      attributes.relationshipId = documentRelsId;
      attributes.id = response.id;
      attributes.fileContent = response.fileContent;
      attributes.fileNameWithExtension = response.fileNameWithExtension;
    }

    const { type, inlineOrAnchored, ...otherAttributes } = attributes;
    // eslint-disable-next-line no-use-before-define
    const imageFragment = buildDrawing(inlineOrAnchored, type, otherAttributes);
    runFragment.import(imageFragment);
  } else if (isVNode(vNode) && vNode.tagName === 'br') {
    const lineBreakFragment = buildLineBreak('textWrapping');
    runFragment.import(lineBreakFragment);
  }
  runFragment.up();

  return runFragment;
};

const buildRunOrRuns = async (vNode, attributes, docxDocumentInstance) => {
  if (isVNode(vNode) && vNode.tagName === 'span') {
    let runFragments = [];

    if (
      vNode.children.length === 0 &&
      vNode.properties?.attributes?.['data-force-space'] === 'true'
    ) {
      vNode.children.push(new VirtualText(' '));
    }

    for (let index = 0; index < vNode.children.length; index++) {
      const childVNode = vNode.children[index];
      const modifiedAttributes = modifiedStyleAttributesBuilder(
        docxDocumentInstance,
        vNode,
        attributes
      );
      const tempRunFragments = await buildRun(childVNode, modifiedAttributes, docxDocumentInstance);
      runFragments = runFragments.concat(
        Array.isArray(tempRunFragments) ? tempRunFragments : [tempRunFragments]
      );
    }

    return runFragments;
  } else {
    const tempRunFragments = await buildRun(vNode, attributes, docxDocumentInstance);
    return tempRunFragments;
  }
};

const buildRunOrHyperLink = async (vNode, attributes, docxDocumentInstance) => {
  if (isVNode(vNode) && vNode.tagName === 'a') {
    const relationshipId = docxDocumentInstance.createDocumentRelationships(
      docxDocumentInstance.relationshipFilename,
      hyperlinkType,
      vNode.properties && vNode.properties.href ? vNode.properties.href : ''
    );
    const hyperlinkFragment = fragment({ namespaceAlias: { w: namespaces.w, r: namespaces.r } })
      .ele('@w', 'hyperlink')
      .att('@r', 'id', `rId${relationshipId}`);

    const modifiedAttributes = { ...attributes };
    modifiedAttributes.hyperlink = true;
    let runFragments = null;
    if (isVNode(vNode) && vNode.tagName === 'span') {
      runFragments = await recursiveRunOrHyperlink(
        vNode.children[0],
        modifiedAttributes,
        docxDocumentInstance
      );
    } else {
      runFragments = await buildRunOrRuns(
        vNode.children[0],
        modifiedAttributes,
        docxDocumentInstance
      );
    }

    if (Array.isArray(runFragments)) {
      for (let iteratorIndex = 0; iteratorIndex < runFragments.length; iteratorIndex++) {
        const runOrHyperlinkFragment = runFragments[iteratorIndex];

        hyperlinkFragment.import(runOrHyperlinkFragment);
      }
    } else {
      hyperlinkFragment.import(runFragments);
    }
    hyperlinkFragment.up();

    return hyperlinkFragment;
  }

  if (isVNode(vNode) && vNode.tagName === 'span') {
    if (
      vNode.children.length === 0 &&
      vNode.properties?.attributes?.['data-force-space'] === 'true'
    ) {
      // Create a virtual text node with a single space character
      vNode.children.push(new VirtualText(' '));
    }
    return recursiveRunOrHyperlink(vNode, attributes, docxDocumentInstance);
  } else {
    return buildRunOrRuns(vNode, attributes, docxDocumentInstance);
  }
};

const buildNumberingProperties = (levelId, numberingId) =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'numPr')
    .ele('@w', 'ilvl')
    .att('@w', 'val', String(levelId))
    .up()
    .ele('@w', 'numId')
    .att('@w', 'val', String(numberingId))
    .up()
    .up();

const buildNumberingInstances = () =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'num')
    .ele('@w', 'abstractNumId')
    .up()
    .up();

const buildSpacing = (lineSpacing, beforeSpacing, afterSpacing) => {
  const spacingFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'spacing');

  if (lineSpacing || lineSpacing === 0) {
    spacingFragment.att('@w', 'line', lineSpacing);
    spacingFragment.att('@w', 'lineRule', 'auto');
  }
  if (beforeSpacing || beforeSpacing === 0) {
    spacingFragment.att('@w', 'before', beforeSpacing);
  }
  if (afterSpacing || afterSpacing === 0) {
    spacingFragment.att('@w', 'after', afterSpacing);
  }

  spacingFragment.up();

  return spacingFragment;
};

const buildIndentation = ({ left, right }) => {
  const indentationFragment = fragment({
    namespaceAlias: { w: namespaces.w },
  }).ele('@w', 'ind');

  if (left) {
    indentationFragment.att('@w', 'left', left);
  }
  if (right) {
    indentationFragment.att('@w', 'right', right);
  }

  indentationFragment.up();

  return indentationFragment;
};

const buildPStyle = (style = 'Normal') =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'pStyle')
    .att('@w', 'val', style)
    .up();

const buildHorizontalAlignment = (horizontalAlignment) => {
  if (horizontalAlignment === 'justify') {
    horizontalAlignment = 'both';
  }
  return fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'jc')
    .att('@w', 'val', horizontalAlignment)
    .up()
    .ele('@w', 'spacing')
    .att('@w', 'line', '0')
    .att('@w', 'lineRule', 'auto')
    .att('@w', 'before', '0')
    .att('@w', 'after', '0')
    .up();
};

const buildParagraphBorder = () => {
  const paragraphBorderFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele(
    '@w',
    'pBdr'
  );
  const bordersObject = cloneDeep(paragraphBordersObject);

  Object.keys(bordersObject).forEach((borderName) => {
    if (bordersObject[borderName]) {
      const { size, spacing, color } = bordersObject[borderName];

      const borderFragment = buildBorder(borderName, size, spacing, color);
      paragraphBorderFragment.import(borderFragment);
    }
  });

  paragraphBorderFragment.up();

  return paragraphBorderFragment;
};

const buildParagraphProperties = (attributes) => {
  const paragraphPropertiesFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele(
    '@w',
    'pPr'
  );
  if (attributes && attributes.constructor === Object) {
    Object.keys(attributes).forEach((key) => {
      if (typeof attributes[key] === 'undefined' && !attributes[key]) {
        return;
      }
      switch (key) {
        case 'numbering':
          const { levelId, numberingId } = attributes[key];
          const numberingPropertiesFragment = buildNumberingProperties(levelId, numberingId);
          paragraphPropertiesFragment.import(numberingPropertiesFragment);
          // eslint-disable-next-line no-param-reassign
          delete attributes.numbering;
          break;
        case 'textAlign':
          const horizontalAlignmentFragment = buildHorizontalAlignment(attributes[key]);
          paragraphPropertiesFragment.import(horizontalAlignmentFragment);
          // eslint-disable-next-line no-param-reassign
          delete attributes.textAlign;
          break;
        case 'backgroundColor':
          // Add shading to Paragraph Properties only if display is block
          // Essentially if background color needs to be across the row
          if (attributes.display === 'block') {
            const shadingFragment = buildShading(attributes[key]);
            paragraphPropertiesFragment.import(shadingFragment);
            // FIXME: Inner padding in case of shaded paragraphs.
            const paragraphBorderFragment = buildParagraphBorder();
            paragraphPropertiesFragment.import(paragraphBorderFragment);
            // eslint-disable-next-line no-param-reassign
            delete attributes.backgroundColor;
          }
          break;
        case 'paragraphStyle':
          const pStyleFragment = buildPStyle(attributes.paragraphStyle);
          paragraphPropertiesFragment.import(pStyleFragment);
          delete attributes.paragraphStyle;
          break;
        case 'indentation':
          const indentationFragment = buildIndentation(attributes[key]);
          paragraphPropertiesFragment.import(indentationFragment);
          // eslint-disable-next-line no-param-reassign
          delete attributes.indentation;
          break;
        case 'marginBottom':
          const spacingAfterFragment = buildSpacing(null, null, attributes[key]);
          paragraphPropertiesFragment.import(spacingAfterFragment);
          break;
      }
    });

    // Only add spacing if we have actual content or specific styles
    if (!attributes || (!attributes.isSpacerParagraph && Object.keys(attributes).length === 0)) {
      paragraphPropertiesFragment
        .ele('w:spacing')
        .att('w:line', '0')
        .att('w:lineRule', 'auto')
        .att('w:before', '0')
        .att('w:after', '0')
        .up();
      return paragraphPropertiesFragment;
    }

    // Only add spacing if it's explicitly defined
    if (attributes.lineHeight || attributes.marginTop || attributes.marginBottom) {
      const spacingFragment = buildSpacing(
        attributes.lineHeight,
        attributes.marginTop,
        attributes.marginBottom
      );
      // eslint-disable-next-line no-param-reassign
      delete attributes.lineHeight;
      // eslint-disable-next-line no-param-reassign
      delete attributes.beforeSpacing;
      // eslint-disable-next-line no-param-reassign
      delete attributes.afterSpacing;
      paragraphPropertiesFragment.import(spacingFragment);
    } else {
      // but to prevent additional spacing unwanted.
      const spacingAfterFragment = buildSpacing(0, 0, 0);
      paragraphPropertiesFragment.import(spacingAfterFragment);
    }
  }
  paragraphPropertiesFragment.up();

  return paragraphPropertiesFragment;
};

const computeImageDimensions = (vNode, attributes) => {
  const { maximumWidth, originalWidth, originalHeight } = attributes;
  const aspectRatio = originalWidth / originalHeight;
  const maximumWidthInEMU = TWIPToEMU(maximumWidth);
  let originalWidthInEMU = pixelToEMU(originalWidth);
  let originalHeightInEMU = pixelToEMU(originalHeight);
  if (originalWidthInEMU > maximumWidthInEMU) {
    originalWidthInEMU = maximumWidthInEMU;
    originalHeightInEMU = Math.round(originalWidthInEMU / aspectRatio);
  }
  let modifiedHeight;
  let modifiedWidth;

  if (vNode.properties && vNode.properties.style) {
    if (vNode.properties.style.width) {
      if (vNode.properties.style.width !== 'auto') {
        if (pixelRegex.test(vNode.properties.style.width)) {
          modifiedWidth = pixelToEMU(vNode.properties.style.width.match(pixelRegex)[1]);
        } else if (percentageRegex.test(vNode.properties.style.width)) {
          const percentageValue = vNode.properties.style.width.match(percentageRegex)[1];

          modifiedWidth = Math.round((percentageValue / 100) * originalWidthInEMU);
        }
      } else {
        // eslint-disable-next-line no-lonely-if
        if (vNode.properties.style.height && vNode.properties.style.height === 'auto') {
          modifiedWidth = originalWidthInEMU;
          modifiedHeight = originalHeightInEMU;
        }
      }
    }
    if (vNode.properties.style.height) {
      if (vNode.properties.style.height !== 'auto') {
        if (pixelRegex.test(vNode.properties.style.height)) {
          modifiedHeight = pixelToEMU(vNode.properties.style.height.match(pixelRegex)[1]);
        } else if (percentageRegex.test(vNode.properties.style.height)) {
          const percentageValue = vNode.properties.style.width.match(percentageRegex)[1];

          modifiedHeight = Math.round((percentageValue / 100) * originalHeightInEMU);
          if (!modifiedWidth) {
            modifiedWidth = Math.round(modifiedHeight * aspectRatio);
          }
        }
      } else {
        // eslint-disable-next-line no-lonely-if
        if (modifiedWidth) {
          if (!modifiedHeight) {
            modifiedHeight = Math.round(modifiedWidth / aspectRatio);
          }
        } else {
          modifiedHeight = originalHeightInEMU;
          modifiedWidth = originalWidthInEMU;
        }
      }
    }
    if (modifiedWidth && !modifiedHeight) {
      modifiedHeight = Math.round(modifiedWidth / aspectRatio);
    } else if (modifiedHeight && !modifiedWidth) {
      modifiedWidth = Math.round(modifiedHeight * aspectRatio);
    }
  } else {
    modifiedWidth = originalWidthInEMU;
    modifiedHeight = originalHeightInEMU;
  }

  // eslint-disable-next-line no-param-reassign
  attributes.width = modifiedWidth;
  // eslint-disable-next-line no-param-reassign
  attributes.height = modifiedHeight;
};

const buildParagraph = async (vNode, attributes, docxDocumentInstance) => {
  // Skip empty text nodes without attributes
  if (isVText(vNode) && !vNode.text.trim() && !attributes) {
    return null;
  }

  // Skip empty nodes without attributes, unless it's a spacer paragraph
  if (!vNode && !attributes.isSpacerParagraph && !attributes) {
    return null;
  }

  // For empty paragraphs, ensure no spacing is added
  const isEmpty = !vNode || (isVText(vNode) && !vNode.text.trim());
  if (isEmpty && !attributes.isSpacerParagraph) {
    attributes = { ...attributes, beforeSpacing: 0, afterSpacing: 0, lineSpacing: 0 };
  }

  const paragraphFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'p');

  const modifiedAttributes = modifiedStyleAttributesBuilder(
    docxDocumentInstance,
    vNode,
    attributes,
    {
      isParagraph: true,
    }
  );

  // For spacer paragraphs, ensure we have minimal spacing
  if (attributes.isSpacerParagraph) {
    modifiedAttributes.lineSpacing = 240; // Standard line spacing
    modifiedAttributes.beforeSpacing = 0;
    modifiedAttributes.afterSpacing = 0;
  }

  // For empty block elements, ensure we have the correct spacing
  if (
    (!vNode || !vNodeHasChildren(vNode)) &&
    !attributes.paragraphStyle &&
    !attributes.isSpacerParagraph
  ) {
    modifiedAttributes.lineSpacing = 0;
    modifiedAttributes.beforeSpacing = 0;
    modifiedAttributes.afterSpacing = 0;
  }

  const paragraphPropertiesFragment = buildParagraphProperties(modifiedAttributes);
  paragraphFragment.import(paragraphPropertiesFragment);

  if (isVNode(vNode) && vNodeHasChildren(vNode)) {
    if (
      [
        'span',
        'strong',
        'b',
        'em',
        'i',
        'u',
        'ins',
        'strike',
        'del',
        's',
        'sub',
        'sup',
        'mark',
        'a',
        'code',
        'pre',
      ].includes(vNode.tagName)
    ) {
      const runOrHyperlinkFragments = await buildRunOrHyperLink(
        vNode,
        modifiedAttributes,
        docxDocumentInstance
      );

      if (Array.isArray(runOrHyperlinkFragments)) {
        for (
          let iteratorIndex = 0;
          iteratorIndex < runOrHyperlinkFragments.length;
          iteratorIndex++
        ) {
          const runOrHyperlinkFragment = runOrHyperlinkFragments[iteratorIndex];

          paragraphFragment.import(runOrHyperlinkFragment);
        }
      } else {
        paragraphFragment.import(runOrHyperlinkFragments);
      }
    } else if (vNode.tagName === 'blockquote') {
      const runFragments = await buildRunOrRuns(vNode, modifiedAttributes, docxDocumentInstance);
      if (Array.isArray(runFragments)) {
        for (let index = 0; index < runFragments.length; index++) {
          const runFragment = runFragments[index];

          paragraphFragment.import(runFragment);
        }
      } else {
        paragraphFragment.import(runFragments);
      }
    } else {
      for (let index = 0; index < vNode.children.length; index++) {
        const childVNode = vNode.children[index];
        if (childVNode.tagName === 'img') {
          let base64String;
          const imageSource = childVNode.properties.src;
          if (isValidUrl(imageSource)) {
            base64String = await imageToBase64(imageSource).catch((error) => {
              // eslint-disable-next-line no-console
              console.warning(`skipping image download and conversion due to ${error}`);
            });

            if (base64String && mimeTypes.lookup(imageSource)) {
              childVNode.properties.src = `data:${mimeTypes.lookup(
                imageSource
              )};base64, ${base64String}`;
            } else {
              break;
            }
          } else {
            base64String = decodeURIComponent(childVNode.properties.src);
            const match = base64String.match(/^data:([A-Za-z-+/]+);base64,(.+)$/);
            if (match) {
              // eslint-disable-next-line prefer-destructuring
              base64String = match[2];
            }
          }

          const imageBuffer = Buffer.from(decodeURIComponent(base64String), 'base64');
          const imageProperties = sizeOf(imageBuffer);

          modifiedAttributes.maximumWidth =
            modifiedAttributes.maximumWidth || docxDocumentInstance.availableDocumentSpace;
          modifiedAttributes.originalWidth = imageProperties.width;
          modifiedAttributes.originalHeight = imageProperties.height;

          computeImageDimensions(childVNode, modifiedAttributes);
        }
        const runOrHyperlinkFragments = await buildRunOrHyperLink(
          childVNode,
          isVNode(childVNode) && childVNode.tagName === 'img'
            ? { ...modifiedAttributes, type: 'picture', description: childVNode.properties.alt }
            : modifiedAttributes,
          docxDocumentInstance
        );
        if (Array.isArray(runOrHyperlinkFragments)) {
          for (
            let iteratorIndex = 0;
            iteratorIndex < runOrHyperlinkFragments.length;
            iteratorIndex++
          ) {
            const runOrHyperlinkFragment = runOrHyperlinkFragments[iteratorIndex];

            paragraphFragment.import(runOrHyperlinkFragment);
          }
        } else {
          paragraphFragment.import(runOrHyperlinkFragments);
        }
      }
    }
  } else {
    // In case paragraphs has to be rendered where vText is present. Eg. table-cell
    // Or in case the vNode is something like img
    if (isVNode(vNode) && vNode.tagName === 'img') {
      const imageSource = vNode.properties.src;
      let base64String = imageSource;
      if (isValidUrl(imageSource)) {
        base64String = await imageToBase64(imageSource).catch((error) => {
          // eslint-disable-next-line no-console
          console.warning(`skipping image download and conversion due to ${error}`);
        });

        if (base64String && mimeTypes.lookup(imageSource)) {
          vNode.properties.src = `data:${mimeTypes.lookup(imageSource)};base64, ${base64String}`;
        } else {
          paragraphFragment.up();

          return paragraphFragment;
        }
      } else {
        base64String = decodeURIComponent(vNode.properties.src);
        const match = base64String.match(/^data:([A-Za-z-+/]+);base64,(.+)$/);
        if (match) {
          // eslint-disable-next-line prefer-destructuring
          base64String = match[2];
        }
      }

      const imageBuffer = Buffer.from(decodeURIComponent(base64String), 'base64');
      const imageProperties = sizeOf(imageBuffer);

      modifiedAttributes.maximumWidth =
        modifiedAttributes.maximumWidth || docxDocumentInstance.availableDocumentSpace;
      modifiedAttributes.originalWidth = imageProperties.width;
      modifiedAttributes.originalHeight = imageProperties.height;

      computeImageDimensions(vNode, modifiedAttributes);
    }
    const runFragments = await buildRunOrRuns(vNode, modifiedAttributes, docxDocumentInstance);
    if (Array.isArray(runFragments)) {
      for (let index = 0; index < runFragments.length; index++) {
        const runFragment = runFragments[index];

        paragraphFragment.import(runFragment);
      }
    } else {
      paragraphFragment.import(runFragments);
    }
  }
  paragraphFragment.up();

  if (modifiedAttributes.marginBottom) {
    const spacingFragment = buildSpacing(null, null, modifiedAttributes.marginBottom);
    // eslint-disable-next-line no-shadow
    const paragraphPropertiesFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele(
      '@w',
      'pPr'
    );
    paragraphPropertiesFragment.import(spacingFragment);
    paragraphFragment.import(paragraphPropertiesFragment);
  }

  paragraphFragment.up();

  return paragraphFragment;
};

// eslint-disable-next-line no-unused-vars
async function resizeNestedTable(vNode, parentColumnWidth, docxDocumentInstance) {
  const tableFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'tbl');
  if (isVNode(vNode) && vNode.properties) {
    const tableStyles = vNode.properties.style || {};
    let width = parentColumnWidth; // Set width to parent column width if data attribute doesn't exist

    if (tableStyles.width) {
      if (tableStyles.width !== 'auto') {
        if (pixelRegex.test(tableStyles.width)) {
          width = pixelToTWIP(tableStyles.width.match(pixelRegex)[1]);
        } else if (percentageRegex.test(tableStyles.width)) {
          const percentageValue = tableStyles.width.match(percentageRegex)[1];

          width = Math.round((percentageValue / 100) * parentColumnWidth);
        }
      } else {
        // eslint-disable-next-line no-lonely-if
        if (vNode.properties.style.height && vNode.properties.style.height === 'auto') {
          width = parentColumnWidth;
        }
      }
    }

    // Set table properties
    tableFragment
      .ele('@w', 'tblPr')
      .ele('@w', 'tblW')
      .att('@w', 'w', width)
      .att('@w', 'type', 'pct')
      .up()
      .ele('@w', 'tblInd')
      .att('@w', 'w', '0')
      .att('@w', 'type', 'dxa')
      .up()
      .ele('@w', 'tblLayout')
      .att('@w', 'type', 'fixed')
      .up()
      .ele('@w', 'tblCellMar')
      .ele('@w', 'tblInd')
      .att('@w', 'w', '0')
      .att('@w', 'type', 'dxa')
      .up()
      .ele('@w', 'test')
      .att('@w', 'w', '55')
      .att('@w', 'type', 'dxa')
      .up()
      .ele('@w', 'top')
      .att('@w', 'w', '555')
      .att('@w', 'type', 'dxa')
      .up()
      .ele('@w', 'left')
      .att('@w', 'w', '11')
      .att('@w', 'type', 'dxa')
      .up()
      .ele('@w', 'bottom')
      .att('@w', 'w', '0')
      .att('@w', 'type', 'dxa')
      .up()
      .ele('@w', 'right')
      .att('@w', 'w', '0')
      .att('@w', 'type', 'dxa')
      .up()
      .up()
      .up();

    const tableGridFragment = tableFragment.ele('@w', 'tblGrid');
    // Set column widths
    let colCount = 0;
    vNode.children.forEach((childVNode) => {
      if (childVNode.tagName === 'colgroup') {
        colCount = childVNode.children.length;
        childVNode.children.forEach((col) => {
          const colWidthStr = col.properties.style.width || 'auto';
          let colWidth = width / colCount; // Default evenly distributed widths

          if (percentageRegex.test(colWidthStr)) {
            colWidth = Math.round((parseFloat(colWidthStr) / 100) * width);
          } else if (pixelRegex.test(colWidthStr)) {
            colWidth = pixelToTWIP(colWidthStr.match(pixelRegex)[1]);
          }
          tableGridFragment.ele('@w', 'gridCol').att('@w', 'w', colWidth).up();
        });
      }
    });
    tableGridFragment.up();

    // Set table rows and cells
    for (let index = 0; index < vNode.children.length; index++) {
      const childVNode = vNode.children[index];
      if (['thead', 'tbody', 'tr'].includes(childVNode.tagName)) {
        // eslint-disable-next-line no-use-before-define
        const tableRowFragment = await buildTableRow(docxDocumentInstance, childVNode.children, {
          width,
        });
        tableFragment.import(tableRowFragment);
      } else if (childVNode.tagName === 'table') {
        const nestedTableFragment = await resizeNestedTable(
          childVNode,
          width,
          docxDocumentInstance
        );
        tableFragment.import(nestedTableFragment);
      }
    }
  }

  tableFragment.up();
  return tableFragment;
}

const parseCSSSpacing = (style, property) => {
  // Handle individual properties (padding-left, padding-right, etc.)
  const specificValue = style[`${property}`];
  if (specificValue) {
    if (specificValue.endsWith('px')) {
      return pixelToTWIP(parseFloat(specificValue));
    }
    // Add other unit conversions as needed
  }

  // Handle shorthand property (padding: 10px 20px, etc.)
  const shorthand = style[property.split('-')[0]];
  if (shorthand) {
    const values = shorthand.split(' ').map((v) => v.trim());
    if (values.length === 1) {
      // Same value for all sides
      return pixelToTWIP(parseFloat(values[0]));
    }
    if (values.length === 2) {
      // vertical horizontal
      return property.includes('left') || property.includes('right')
        ? pixelToTWIP(parseFloat(values[1]))
        : pixelToTWIP(parseFloat(values[0]));
    }
    if (values.length === 4) {
      // top right bottom left
      // eslint-disable-next-line no-nested-ternary
      const index = property.includes('top')
        ? 0
        : // eslint-disable-next-line no-nested-ternary
        property.includes('right')
        ? 1
        : property.includes('bottom')
        ? 2
        : 3;
      return pixelToTWIP(parseFloat(values[index]));
    }
  }

  return null;
};

const buildTableRow = async function buildTableRow(docxDocumentInstance, columns, attributes = {}) {
  const width = attributes.width || attributes.maximumWidth || '100%';
  const tableRowFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'tr');

  const needsRowHeight = columns.some((column) => {
    const style = column.properties?.style || {};
    const hasPadding = [
      style.padding,
      style['padding-top'],
      style['padding-bottom'],
      style.margin,
      style['margin-top'],
      style['margin-bottom'],
    ].some((value) => {
      if (!value) return false;
      // Convert percentage or pixel values to numbers
      const numericValue = parseFloat(value);
      // eslint-disable-next-line no-restricted-globals
      return !isNaN(numericValue) && numericValue > 0;
    });

    return hasPadding;
  });
  // Add consistent row height
  const trPr = tableRowFragment.ele('@w', 'trPr');

  if (needsRowHeight) {
    trPr
      .ele('@w', 'trHeight')
      .att('@w', 'val', '400') // Set a default height of 400 twips (about 0.28 inches)
      .att('@w', 'hRule', 'atLeast'); // atLeast ensures minimum height while allowing expansion if needed
  } else {
    trPr.ele('@w', 'trHeight').att('@w', 'hRule', 'atLeast'); // atLeast ensures minimum height while allowing expansion if needed
  }

  // Process each column
  // eslint-disable-next-line no-restricted-syntax
  for (const column of columns) {
    const colspan = parseInt(column.properties?.colSpan || '1', 10);
    const totalPadding = columns.length * 200;
    let colWidth = 'auto';
    let colWidthTwips = docxDocumentInstance.availableDocumentSpace;
    if (column.properties?.attributes?.['data-docx-column']) {
      colWidth = parseInt(column.properties?.attributes?.['data-docx-column'] || '1', 10);
      colWidthTwips =
        Math.floor((colWidth / 12) * docxDocumentInstance.availableDocumentSpace) - totalPadding;
    } else {
      const colWidthStr = column.properties.style.width || 'auto';
      colWidth = columns / columns.length; // Default evenly distributed widths

      if (percentageRegex.test(colWidthStr)) {
        colWidth = Math.round((parseFloat(colWidthStr) / 100) * width);
        colWidthTwips = colWidth;
      } else if (pixelRegex.test(colWidthStr)) {
        colWidth = pixelToTWIP(colWidthStr.match(pixelRegex)[1]);
        colWidthTwips = colWidth;
      }
    }

    const tableCellFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'tc');
    const align = column.properties?.style?.['text-align'];
    const vAlign = column.properties?.attributes?.valign;
    const cssVAlign = column.properties?.style?.['vertical-align'];
    let cellBgOverride = column.properties?.style?.['background-color'] || false;

    if (cellBgOverride) {
      // Convert RGB to hex if necessary
      if (cellBgOverride.startsWith('rgb')) {
        const rgb = cellBgOverride.match(/\d+/g);
        if (rgb && rgb.length === 3) {
          cellBgOverride = rgbToHex(parseInt(rgb[0]), parseInt(rgb[1]), parseInt(rgb[2]));
        }
      }
      // Remove '#' if present
      cellBgOverride = cellBgOverride.replace('#', '');
    }

    // eslint-disable-next-line no-nested-ternary
    const cssOrValignAlignment = cssVAlign
      ? cssVerticalAlignMapping[cssVAlign] || 'center'
      : vAlign
      ? valignMapping[vAlign] || 'center'
      : 'center';

    if (align) {
      column.children.forEach((child) => {
        if (!child.properties?.style?.['text-align']) {
          child.properties = child.properties || {};
          child.properties.style = child.properties.style || {};
          child.properties.style['text-align'] = column.properties.style['text-align'];
        }
      });
    }

    const style = column.properties?.style || {};

    // Calculate margins from both padding and margin
    const leftSpace =
      parseCSSSpacing(style, 'padding-left') || parseCSSSpacing(style, 'margin-left') || 0; // default value

    const rightSpace =
      parseCSSSpacing(style, 'padding-right') || parseCSSSpacing(style, 'margin-right') || 0; // default value

    const topSpace =
      parseCSSSpacing(style, 'padding-top') || parseCSSSpacing(style, 'margin-top') || 0; // default value

    const bottomSpace =
      parseCSSSpacing(style, 'padding-bottom') || parseCSSSpacing(style, 'margin-bottom') || 0; // default value

    // Set up column width
    tableCellFragment
      .ele('@w', 'tcPr')
      .ele('@w', 'tcW')
      .att('@w', 'w', colWidthTwips.toString())
      .att('@w', 'type', 'dxa')
      .up()
      // Add gridSpan here for colspan support
      .ele('@w', 'gridSpan')
      .att('@w', 'val', colspan.toString())
      .up()
      .ele('@w', 'vAlign')
      .att('@w', 'val', cssOrValignAlignment)
      .up()
      .ele('@w', 'tcMar')
      .ele('@w', 'top')
      .att('@w', 'w', topSpace.toString())
      .att('@w', 'type', 'dxa')
      .up()
      .ele('@w', 'left')
      .att('@w', 'w', leftSpace.toString())
      .att('@w', 'type', 'dxa')
      .up()
      .ele('@w', 'bottom')
      .att('@w', 'w', bottomSpace.toString())
      .att('@w', 'type', 'dxa')
      .up()
      .ele('@w', 'right')
      .att('@w', 'w', rightSpace.toString())
      .att('@w', 'type', 'dxa')
      .up();

    if (cellBgOverride || attributes.backgroundColor) {
      const bgColour = cellBgOverride || attributes.backgroundColor;
      tableCellFragment
        .first()
        .ele('@w', 'shd')
        .att('@w', 'val', 'clear')
        .att('@w', 'fill', bgColour.toUpperCase())
        .up();
    }

    tableCellFragment.up();

    // Import cell content
    await convertVTreeToXML(docxDocumentInstance, column.children, tableCellFragment);

    // Handle nested tables recursively
    // eslint-disable-next-line no-restricted-syntax
    for (const child of column.children) {
      if (isVNode(child) && child.tagName === 'table') {
        const nestedTableFragment = await resizeNestedTable(
          child,
          colWidthTwips,
          docxDocumentInstance
        );
        tableCellFragment.import(nestedTableFragment);
      }
    }

    tableCellFragment.up(); // Close the table cell fragment
    tableRowFragment.import(tableCellFragment);
  }

  tableRowFragment.up(); // Close the table row fragment
  return tableRowFragment;
};

const buildTableGridCol = (gridWidth) =>
  fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'gridCol')
    .att('@w', 'w', String(gridWidth));

const buildTableGrid = (vNode, attributes) => {
  const tableGridFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'tblGrid');
  if (vNodeHasChildren(vNode)) {
    const gridColumns = vNode.children.filter((childVNode) => childVNode.tagName === 'col');
    const gridWidth = (attributes.width || attributes.maximumWidth) / gridColumns.length;
    for (let index = 0; index < gridColumns.length; index++) {
      const tableGridColFragment = buildTableGridCol(gridWidth);
      tableGridFragment.import(tableGridColFragment);
    }
  }
  tableGridFragment.up();

  return tableGridFragment;
};

const buildTableProperties = (attributes) => {
  const { maximumWidth, width, jc } = attributes;
  const tableFragment = fragment({ namespaceAlias: { w: namespaces.w } });
  const tableProperties = tableFragment.ele('@w', 'tblPr');
  if (width || maximumWidth) {
    tableProperties
      .ele('@w', 'tblW')
      .att('@w', 'w', Math.round(((width || maximumWidth) / 100) * 5000))
      .att('@w', 'type', 'dxa');
  }
  if (jc) {
    tableProperties.ele('@w', 'jc').att('@w', 'val', jc);
  }

  // Add table borders if present
  if (attributes.tableBorders) {
    const tblBorders = tableProperties.ele('@w', 'tblBorders');

    ['top', 'left', 'bottom', 'right'].forEach((side) => {
      tblBorders
        .ele('@w', side)
        .att('@w', 'val', attributes.tableBorders.stroke)
        .att('@w', 'sz', attributes.tableBorders[side] * 8)
        .att('@w', 'color', attributes.tableBorders.color);
    });

    if (attributes.tableBorders.insideH || attributes.tableBorders.insideV) {
      ['insideH', 'insideV'].forEach((side) => {
        if (attributes.tableBorders[side]) {
          tblBorders
            .ele('@w', side)
            .att('@w', 'val', attributes.tableBorders.stroke)
            .att('@w', 'sz', attributes.tableBorders[side] * 8)
            .att('@w', 'color', attributes.tableBorders.color);
        }
      });
    }
  }

  tableProperties.ele('@w', 'tblLayout').att('@w', 'type', 'fixed');

  const tableCellMargins = tableProperties.ele('@w', 'tblCellMar');

  tableCellMargins.ele('@w', 'tblInd').att('@w', 'w', 0).att('@w', 'type', 'dxa');

  ['top', 'left', 'bottom', 'right'].forEach((direction) => {
    tableCellMargins
      .ele('@w', direction)
      .att(
        '@w',
        'w',
        attributes[`cellMargin${direction.charAt(0).toUpperCase()}${direction.slice(1)}`] || 0
      )
      .att('@w', 'type', 'dxa');
  });

  return tableFragment;
};

const buildTableGridFromTableRow = (vNode, width) => {
  const tableGridFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'tblGrid');

  const gridColumns = vNode.children.filter(
    (childVNode) => childVNode.tagName === 'td' || childVNode.tagName === 'th'
  );
  let hasColSpan = false;
  for (let index = 0; index < gridColumns.length; index++) {
    const col = gridColumns[index];
    if (col.properties.colSpan) {
      hasColSpan = true;
      break;
    }
  }
  // this grid can't be built from this row.
  if (hasColSpan) {
    return false;
  }

  // First pass: collect all percentage values
  let totalPercentage = 0;
  const columnWidths = gridColumns.map((col) => {
    const colWidthStr = col.attributes?.width || col.properties.style.width || 'auto';
    if (percentageRegex.test(colWidthStr)) {
      const percentage = parseFloat(colWidthStr);
      totalPercentage += percentage;
      return { type: 'percentage', value: percentage };
    } else if (pixelRegex.test(colWidthStr)) {
      const pixels = pixelToTWIP(colWidthStr.match(pixelRegex)[1]);
      return { type: 'pixels', value: pixels };
    }
    return { type: 'auto' };
  });

  // Normalize percentages if needed
  if (totalPercentage > 0 && totalPercentage < 100) {
    const scaleFactor = 100 / totalPercentage;
    columnWidths.forEach((col) => {
      if (col.type === 'percentage') {
        col.value *= scaleFactor;
      }
    });
  }

  // Second pass: calculate final widths and build grid
  for (let i = 0; i < gridColumns.length; i++) {
    let colWidth;
    const col = columnWidths[i];

    if (col.type === 'percentage') {
      colWidth = Math.round((col.value / 100) * width);
    } else if (col.type === 'pixels') {
      colWidth = col.value;
    } else {
      // Distribute remaining width evenly among 'auto' columns
      colWidth = width / gridColumns.length;
    }

    const tableGridColFragment = buildTableGridCol(colWidth);
    tableGridFragment.import(tableGridColFragment);
  }

  tableGridFragment.up();

  return tableGridFragment;
};
const parseBorderStyle = (borderStyle) => {
  if (!borderStyle) return {};

  // Match width, style, and color in any order
  // First extract any rgb/rgba values to prevent splitting them
  const rgbMatches = borderStyle.match(/rgba?\([^)]+\)|hsla?\([^)]+\)/g) || [];
  let remainingStyle = borderStyle;
  const rgbPlaceholders = {};

  // Replace rgb/rgba/hsl/hsla values with placeholders
  rgbMatches.forEach((match, index) => {
    // Ensure the match is properly formatted
    if (
      !match.match(/^(rgba?\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*(?:,\s*[\d.]+\s*)?\)|hsla?\([^)]+\))$/)
    ) {
      // eslint-disable-next-line no-console
      console.warn(`Invalid color format found in border style: ${match}`);
      return;
    }
    const placeholder = `__COLOR_${index}__`;
    rgbPlaceholders[placeholder] = match;
    remainingStyle = remainingStyle.replace(match, placeholder);
  });

  const parts = remainingStyle.split(/\s+/);
  const result = {};

  parts.forEach((part) => {
    // Restore rgb values from placeholders
    const actualPart = rgbPlaceholders[part] || part;

    if (actualPart.match(/^[0-9]+(\.[0-9]+)?(px|pt|em|rem)$/)) {
      result.width = parseInt(actualPart);
    } else if (
      actualPart.match(/^(none|hidden|dotted|dashed|solid|double|groove|ridge|inset|outset)$/)
    ) {
      result.style = actualPart;
    } else if (
      actualPart.match(/^(rgb|rgba|#|hsl|hsla)/i) ||
      Object.prototype.hasOwnProperty.call(colorNames, actualPart.toLowerCase())
    ) {
      try {
        result.color = fixupColorCode(actualPart);
      } catch (error) {
        // eslint-disable-next-line no-console
        console.warn(`Error parsing color value: ${actualPart}`, error);
        result.color = '000000'; // Fallback to black
      }
    }
  });

  return result;
};
const buildTable = async (vNode, attributes, docxDocumentInstance) => {
  const tableFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'tbl');
  const modifiedAttributes = { ...attributes };

  if (isVNode(vNode) && vNode.properties) {
    const tableAttributes = vNode.properties.attributes || {};
    const tableStyles = vNode.properties.style || {};

    // Check if border-style is hidden - this takes precedence over border attribute
    if (tableStyles['border-style'] === 'hidden') {
      modifiedAttributes.tableBorders = {
        top: 0,
        bottom: 0,
        left: 0,
        right: 0,
        stroke: 'none',
        color: 'auto',
      };
    } else {
      // Initialize border properties
      // eslint-disable-next-line prefer-const
      let [borderSize, borderStrike, borderColor] = [0, 'single', '000000'];

      // Check HTML border attribute
      // eslint-disable-next-line no-restricted-globals
      if (!isNaN(tableAttributes.border)) {
        borderSize = parseInt(tableAttributes.border, 10);
      }

      if (tableStyles?.border) {
        const border = parseBorderStyle(tableStyles.border);
        borderSize = border.width || borderSize;
        borderColor = border.color || borderColor;
      }

      if (tableStyles?.['border-color']) {
        const border = parseBorderStyle(tableStyles['border-color']);
        borderColor = border.color || borderColor;
      }

      if (
        tableStyles['background-color'] &&
        !colorlessColors.includes(tableStyles['background-color'])
      ) {
        modifiedAttributes.backgroundColor = fixupColorCode(tableStyles['background-color']);
      }

      if (tableStyles?.['border-width']) {
        const parsed = String(tableStyles?.['border-width']).replace(/[^0-9.]/g, '');
        const parsedBorderWidth = parsed === '' ? false : parseInt(parsed);
        if (parsedBorderWidth !== false) {
          borderSize = parsedBorderWidth;
        }
      }

      if (tableStyles?.['border-style'] && tableStyles?.['border-style'] === 'none') {
        borderSize = 0;
      }

      // Only add borders if we have a size > 0
      if (borderSize > 0) {
        modifiedAttributes.tableBorders = {
          top: borderSize,
          bottom: borderSize,
          left: borderSize,
          right: borderSize,
          stroke: borderStrike,
          color: borderColor,
        };

        // Add inside borders for collapsed tables
        if (tableStyles['border-collapse'] === 'collapse') {
          modifiedAttributes.tableBorders.insideH = borderSize;
          modifiedAttributes.tableBorders.insideV = borderSize;
        }
      }
    }

    let minimumWidth;
    let maximumWidth;
    let width;
    tableStyles['min-width'] = tableStyles['min-width'] || tableStyles.width;
    if (tableStyles['min-width']) {
      if (pixelRegex.test(tableStyles['min-width'])) {
        minimumWidth = pixelToTWIP(tableStyles['min-width'].match(pixelRegex)[1]);
      } else if (percentageRegex.test(tableStyles['min-width'])) {
        const percentageValue = tableStyles['min-width'].match(percentageRegex)[1];

        minimumWidth = Math.round((percentageValue / 100) * attributes.maximumWidth);
      }
    }

    const {
      margin: tblMargin,
      'margin-left': tblMarginLeft,
      'margin-right': tblMarginRight,
    } = tableStyles;
    if (
      (tblMargin && tblMargin === 'auto') ||
      (tblMarginLeft && tblMarginLeft === 'auto' && tblMarginRight && tblMarginRight === 'auto')
    ) {
      modifiedAttributes.jc = 'center';
    }
    if (tableStyles['max-width']) {
      if (pixelRegex.test(tableStyles['max-width'])) {
        pixelRegex.lastIndex = 0;
        maximumWidth = pixelToTWIP(tableStyles['max-width'].match(pixelRegex)[1]);
      } else if (percentageRegex.test(tableStyles['max-width'])) {
        percentageRegex.lastIndex = 0;
        const percentageValue = tableStyles['max-width'].match(percentageRegex)[1];
        maximumWidth = Math.round((percentageValue / 100) * attributes.maximumWidth);
      }
    }
    if (tableStyles.width) {
      if (pixelRegex.test(tableStyles.width)) {
        pixelRegex.lastIndex = 0;
        width = pixelToTWIP(tableStyles.width.match(pixelRegex)[1]);
      } else if (percentageRegex.test(tableStyles.width)) {
        percentageRegex.lastIndex = 0;
        const percentageValue = tableStyles.width.match(percentageRegex)[1];
        width = Math.round(Math.round((percentageValue / 100) * 12240));
      }
    }

    if (width) {
      modifiedAttributes.width = width;
      if (maximumWidth) {
        modifiedAttributes.width = Math.min(modifiedAttributes.width, maximumWidth);
      }
      if (minimumWidth) {
        modifiedAttributes.width = Math.max(modifiedAttributes.width, minimumWidth);
      }
    } else if (minimumWidth) {
      modifiedAttributes.width = minimumWidth;
    }
    if (modifiedAttributes.width) {
      modifiedAttributes.width = Math.min(modifiedAttributes.width, attributes.maximumWidth);
    }
  }

  // Count the number of columns
  let columnCount = 0;
  if (vNodeHasChildren(vNode)) {
    const firstRow = vNode.children.find(
      (child) =>
        child.tagName === 'tr' ||
        (child.tagName === 'thead' &&
          child.children.find((grandChild) => grandChild.tagName === 'tr')) ||
        (child.tagName === 'tbody' &&
          child.children.find((grandChild) => grandChild.tagName === 'tr'))
    );

    if (firstRow) {
      const cells =
        firstRow.tagName === 'tr'
          ? firstRow.children
          : firstRow.children.find((child) => child.tagName === 'tr').children;
      columnCount = cells.filter((cell) => cell.tagName === 'td' || cell.tagName === 'th').length;
    }
  }

  const paddingPerCell = 0; // 100 twips left + 100 twips right
  const totalPadding = columnCount * paddingPerCell;
  if (modifiedAttributes.width) {
    modifiedAttributes.width = Math.max(modifiedAttributes.width - totalPadding, 0);
  }

  const tablePropertiesFragment = buildTableProperties(modifiedAttributes);
  tableFragment.import(tablePropertiesFragment);

  // Add table borders if size > 0
  if (modifiedAttributes.tableBorders) {
    const tblBorders = tablePropertiesFragment.ele('@w', 'tblBorders');

    // Add border for each side
    ['top', 'left', 'bottom', 'right'].forEach((side) => {
      tblBorders
        .ele('@w', side)
        .att('@w', 'val', modifiedAttributes.tableBorders.stroke)
        .att('@w', 'sz', modifiedAttributes.tableBorders[side] * 8)
        .att('@w', 'color', modifiedAttributes.tableBorders.color);
    });

    // Add inside borders for collapsed borders
    if (modifiedAttributes.tableBorders.insideH || modifiedAttributes.tableBorders.insideV) {
      ['insideH', 'insideV'].forEach((side) => {
        if (modifiedAttributes.tableBorders[side]) {
          tblBorders
            .ele('@w', side)
            .att('@w', 'val', modifiedAttributes.tableBorders.stroke)
            .att('@w', 'sz', modifiedAttributes.tableBorders[side] * 8)
            .att('@w', 'color', modifiedAttributes.tableBorders.color);
        }
      });
    }
  }

  let hasGrid = false;
  if (vNodeHasChildren(vNode)) {
    for (let index = 0; index < vNode.children.length; index++) {
      const childVNode = vNode.children[index];
      if (childVNode.tagName === 'colgroup') {
        const tableGridFragment = buildTableGrid(childVNode, modifiedAttributes);
        tableFragment.import(tableGridFragment);
        hasGrid = true;
      } else if (childVNode.tagName === 'thead') {
        for (let iteratorIndex = 0; iteratorIndex < childVNode.children.length; iteratorIndex++) {
          const grandChildVNode = childVNode.children[iteratorIndex];
          if (grandChildVNode.tagName === 'tr') {
            // Extract the <td> or <th> elements as columns
            const columns = grandChildVNode.children.filter(
              (child) => child.tagName === 'td' || child.tagName === 'th'
            );

            if (iteratorIndex === 0 && !hasGrid) {
              const tableGridFragment = buildTableGridFromTableRow(
                grandChildVNode,
                modifiedAttributes.width
              );
              if (tableGridFragment) {
                tableFragment.import(tableGridFragment);
              }
            }
            const tableRowFragment = await buildTableRow(
              docxDocumentInstance,
              columns,
              modifiedAttributes
            );
            tableFragment.import(tableRowFragment);
          }
        }
      } else if (childVNode.tagName === 'tbody') {
        for (let iteratorIndex = 0; iteratorIndex < childVNode.children.length; iteratorIndex++) {
          const grandChildVNode = childVNode.children[iteratorIndex];
          if (grandChildVNode.tagName === 'tr') {
            // Extract the <td> or <th> elements as columns
            const columns = grandChildVNode.children.filter(
              (child) => child.tagName === 'td' || child.tagName === 'th'
            );

            if (iteratorIndex === 0 && !hasGrid) {
              const tableGridFragment = buildTableGridFromTableRow(
                grandChildVNode,
                modifiedAttributes.width
              );
              if (tableGridFragment) {
                tableFragment.import(tableGridFragment);
              }
            }
            const tableRowFragment = await buildTableRow(
              docxDocumentInstance,
              columns,
              modifiedAttributes
            );
            tableFragment.import(tableRowFragment);
          }
        }
      } else if (childVNode.tagName === 'tr') {
        // Extract the <td> or <th> elements as columns
        const columns = childVNode.children.filter(
          (child) => child.tagName === 'td' || child.tagName === 'th'
        );

        if (index === 0 && !hasGrid) {
          const tableGridFragment = buildTableGridFromTableRow(
            childVNode,
            modifiedAttributes.width
          );
          if (tableGridFragment) {
            tableFragment.import(tableGridFragment);
          }
        }

        const tableRowFragment = await buildTableRow(
          docxDocumentInstance,
          columns,
          modifiedAttributes
        );
        tableFragment.import(tableRowFragment);
      }
    }
  }

  tableFragment.up(); // Complete the table fragment.
  return tableFragment;
};

const buildPresetGeometry = () =>
  fragment({ namespaceAlias: { a: namespaces.a } })
    .ele('@a', 'prstGeom')
    .att('prst', 'rect')
    .up();

const buildExtents = ({ width, height }) =>
  fragment({ namespaceAlias: { a: namespaces.a } })
    .ele('@a', 'ext')
    .att('cx', width)
    .att('cy', height)
    .up();

const buildOffset = () =>
  fragment({ namespaceAlias: { a: namespaces.a } })
    .ele('@a', 'off')
    .att('x', '0')
    .att('y', '0')
    .up();

const buildGraphicFrameTransform = (attributes) => {
  const graphicFrameTransformFragment = fragment({ namespaceAlias: { a: namespaces.a } }).ele(
    '@a',
    'xfrm'
  );

  const offsetFragment = buildOffset();
  graphicFrameTransformFragment.import(offsetFragment);
  const extentsFragment = buildExtents(attributes);
  graphicFrameTransformFragment.import(extentsFragment);

  graphicFrameTransformFragment.up();

  return graphicFrameTransformFragment;
};

const buildShapeProperties = (attributes) => {
  const shapeProperties = fragment({ namespaceAlias: { pic: namespaces.pic } }).ele('@pic', 'spPr');

  const graphicFrameTransformFragment = buildGraphicFrameTransform(attributes);
  shapeProperties.import(graphicFrameTransformFragment);
  const presetGeometryFragment = buildPresetGeometry();
  shapeProperties.import(presetGeometryFragment);

  shapeProperties.up();

  return shapeProperties;
};

const buildFillRect = () =>
  fragment({ namespaceAlias: { a: namespaces.a } })
    .ele('@a', 'fillRect')
    .up();

const buildStretch = () => {
  const stretchFragment = fragment({ namespaceAlias: { a: namespaces.a } }).ele('@a', 'stretch');

  const fillRectFragment = buildFillRect();
  stretchFragment.import(fillRectFragment);

  stretchFragment.up();

  return stretchFragment;
};

const buildSrcRectFragment = () =>
  fragment({ namespaceAlias: { a: namespaces.a } })
    .ele('@a', 'srcRect')
    .att('b', '0')
    .att('l', '0')
    .att('r', '0')
    .att('t', '0')
    .up();

const buildBinaryLargeImageOrPicture = (relationshipId) =>
  fragment({
    namespaceAlias: { a: namespaces.a, r: namespaces.r },
  })
    .ele('@a', 'blip')
    .att('@r', 'embed', `rId${relationshipId}`)
    // FIXME: possible values 'email', 'none', 'print', 'hqprint', 'screen'
    .att('cstate', 'print')
    .up();

const buildBinaryLargeImageOrPictureFill = (relationshipId) => {
  const binaryLargeImageOrPictureFillFragment = fragment({
    namespaceAlias: { pic: namespaces.pic },
  }).ele('@pic', 'blipFill');
  const binaryLargeImageOrPictureFragment = buildBinaryLargeImageOrPicture(relationshipId);
  binaryLargeImageOrPictureFillFragment.import(binaryLargeImageOrPictureFragment);
  const srcRectFragment = buildSrcRectFragment();
  binaryLargeImageOrPictureFillFragment.import(srcRectFragment);
  const stretchFragment = buildStretch();
  binaryLargeImageOrPictureFillFragment.import(stretchFragment);

  binaryLargeImageOrPictureFillFragment.up();

  return binaryLargeImageOrPictureFillFragment;
};

const buildNonVisualPictureDrawingProperties = () =>
  fragment({ namespaceAlias: { pic: namespaces.pic } })
    .ele('@pic', 'cNvPicPr')
    .up();

const buildNonVisualDrawingProperties = (
  pictureId,
  pictureNameWithExtension,
  pictureDescription = ''
) =>
  fragment({ namespaceAlias: { pic: namespaces.pic } })
    .ele('@pic', 'cNvPr')
    .att('id', pictureId)
    .att('name', pictureNameWithExtension)
    .att('descr', pictureDescription)
    .up();

const buildNonVisualPictureProperties = (
  pictureId,
  pictureNameWithExtension,
  pictureDescription
) => {
  const nonVisualPicturePropertiesFragment = fragment({
    namespaceAlias: { pic: namespaces.pic },
  }).ele('@pic', 'nvPicPr');
  // TODO: Handle picture attributes
  const nonVisualDrawingPropertiesFragment = buildNonVisualDrawingProperties(
    pictureId,
    pictureNameWithExtension,
    pictureDescription
  );
  nonVisualPicturePropertiesFragment.import(nonVisualDrawingPropertiesFragment);
  const nonVisualPictureDrawingPropertiesFragment = buildNonVisualPictureDrawingProperties();
  nonVisualPicturePropertiesFragment.import(nonVisualPictureDrawingPropertiesFragment);
  nonVisualPicturePropertiesFragment.up();

  return nonVisualPicturePropertiesFragment;
};

const buildPicture = ({
  id,
  fileNameWithExtension,
  description,
  relationshipId,
  width,
  height,
}) => {
  const pictureFragment = fragment({ namespaceAlias: { pic: namespaces.pic } }).ele('@pic', 'pic');
  const nonVisualPicturePropertiesFragment = buildNonVisualPictureProperties(
    id,
    fileNameWithExtension,
    description
  );
  pictureFragment.import(nonVisualPicturePropertiesFragment);
  const binaryLargeImageOrPictureFill = buildBinaryLargeImageOrPictureFill(relationshipId);
  pictureFragment.import(binaryLargeImageOrPictureFill);
  const shapeProperties = buildShapeProperties({ width, height });
  pictureFragment.import(shapeProperties);
  pictureFragment.up();

  return pictureFragment;
};

const buildGraphicData = (graphicType, attributes) => {
  const graphicDataFragment = fragment({ namespaceAlias: { a: namespaces.a } })
    .ele('@a', 'graphicData')
    .att('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture');
  if (graphicType === 'picture') {
    const pictureFragment = buildPicture(attributes);
    graphicDataFragment.import(pictureFragment);
  }
  graphicDataFragment.up();

  return graphicDataFragment;
};

const buildGraphic = (graphicType, attributes) => {
  const graphicFragment = fragment({ namespaceAlias: { a: namespaces.a } }).ele('@a', 'graphic');
  // TODO: Handle drawing type
  const graphicDataFragment = buildGraphicData(graphicType, attributes);
  graphicFragment.import(graphicDataFragment);
  graphicFragment.up();

  return graphicFragment;
};

const buildDrawingObjectNonVisualProperties = (pictureId, pictureName) =>
  fragment({ namespaceAlias: { wp: namespaces.wp } })
    .ele('@wp', 'docPr')
    .att('id', pictureId)
    .att('name', pictureName)
    .up();

const buildWrapSquare = () =>
  fragment({ namespaceAlias: { wp: namespaces.wp } })
    .ele('@wp', 'wrapSquare')
    .att('wrapText', 'bothSides')
    .att('distB', '228600')
    .att('distT', '228600')
    .att('distL', '228600')
    .att('distR', '228600')
    .up();

// eslint-disable-next-line no-unused-vars
const buildWrapNone = () =>
  fragment({ namespaceAlias: { wp: namespaces.wp } })
    .ele('@wp', 'wrapNone')
    .up();

const buildEffectExtentFragment = () =>
  fragment({ namespaceAlias: { wp: namespaces.wp } })
    .ele('@wp', 'effectExtent')
    .att('b', '0')
    .att('l', '0')
    .att('r', '0')
    .att('t', '0')
    .up();

const buildExtent = ({ width, height }) =>
  fragment({ namespaceAlias: { wp: namespaces.wp } })
    .ele('@wp', 'extent')
    .att('cx', width)
    .att('cy', height)
    .up();

const buildPositionV = () =>
  fragment({ namespaceAlias: { wp: namespaces.wp } })
    .ele('@wp', 'positionV')
    .att('relativeFrom', 'paragraph')
    .ele('@wp', 'posOffset')
    .txt('19050')
    .up()
    .up();

const buildPositionH = () =>
  fragment({ namespaceAlias: { wp: namespaces.wp } })
    .ele('@wp', 'positionH')
    .att('relativeFrom', 'column')
    .ele('@wp', 'posOffset')
    .txt('19050')
    .up()
    .up();

const buildSimplePos = () =>
  fragment({ namespaceAlias: { wp: namespaces.wp } })
    .ele('@wp', 'simplePos')
    .att('x', '0')
    .att('y', '0')
    .up();

const buildAnchoredDrawing = (graphicType, attributes) => {
  const anchoredDrawingFragment = fragment({ namespaceAlias: { wp: namespaces.wp } })
    .ele('@wp', 'anchor')
    .att('distB', '0')
    .att('distL', '0')
    .att('distR', '0')
    .att('distT', '0')
    .att('relativeHeight', '0')
    .att('behindDoc', 'false')
    .att('locked', 'true')
    .att('layoutInCell', 'true')
    .att('allowOverlap', 'false')
    .att('simplePos', 'false');
  // Even though simplePos isnt supported by Word 2007 simplePos is required.
  const simplePosFragment = buildSimplePos();
  anchoredDrawingFragment.import(simplePosFragment);
  const positionHFragment = buildPositionH();
  anchoredDrawingFragment.import(positionHFragment);
  const positionVFragment = buildPositionV();
  anchoredDrawingFragment.import(positionVFragment);
  const extentFragment = buildExtent({ width: attributes.width, height: attributes.height });
  anchoredDrawingFragment.import(extentFragment);
  const effectExtentFragment = buildEffectExtentFragment();
  anchoredDrawingFragment.import(effectExtentFragment);
  const wrapSquareFragment = buildWrapSquare();
  anchoredDrawingFragment.import(wrapSquareFragment);
  const drawingObjectNonVisualPropertiesFragment = buildDrawingObjectNonVisualProperties(
    attributes.id,
    attributes.fileNameWithExtension
  );
  anchoredDrawingFragment.import(drawingObjectNonVisualPropertiesFragment);
  const graphicFragment = buildGraphic(graphicType, attributes);
  anchoredDrawingFragment.import(graphicFragment);

  anchoredDrawingFragment.up();

  return anchoredDrawingFragment;
};

const buildInlineDrawing = (graphicType, attributes) => {
  const inlineDrawingFragment = fragment({ namespaceAlias: { wp: namespaces.wp } })
    .ele('@wp', 'inline')
    .att('distB', '0')
    .att('distL', '0')
    .att('distR', '0')
    .att('distT', '0');

  const extentFragment = buildExtent({ width: attributes.width, height: attributes.height });
  inlineDrawingFragment.import(extentFragment);
  const effectExtentFragment = buildEffectExtentFragment();
  inlineDrawingFragment.import(effectExtentFragment);
  const drawingObjectNonVisualPropertiesFragment = buildDrawingObjectNonVisualProperties(
    attributes.id,
    attributes.fileNameWithExtension
  );
  inlineDrawingFragment.import(drawingObjectNonVisualPropertiesFragment);
  const graphicFragment = buildGraphic(graphicType, attributes);
  inlineDrawingFragment.import(graphicFragment);

  inlineDrawingFragment.up();

  return inlineDrawingFragment;
};

const buildDrawing = (inlineOrAnchored = false, graphicType, attributes) => {
  const drawingFragment = fragment({ namespaceAlias: { w: namespaces.w } }).ele('@w', 'drawing');
  const inlineOrAnchoredDrawingFragment = inlineOrAnchored
    ? buildInlineDrawing(graphicType, attributes)
    : buildAnchoredDrawing(graphicType, attributes);
  drawingFragment.import(inlineOrAnchoredDrawingFragment);
  drawingFragment.up();

  return drawingFragment;
};

export {
  buildParagraph,
  buildTable,
  buildTableRow,
  buildNumberingInstances,
  buildLineBreak,
  buildIndentation,
  buildTextElement,
  buildBold,
  buildItalics,
  buildUnderline,
  buildDrawing,
  fixupLineHeight,
};
