import namespaces from '../namespaces';

const generateEmbeddedFontsXML = (fontName, fontData) => `
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <pkg:package xmlns:pkg="${namespaces.pkg}">
    <pkg:part pkg:name="/word/fonts/${fontName}.odttf" pkg:contentType="application/vnd.openxmlformats-officedocument.obfuscatedFont">
      <pkg:binaryData>${fontData}</pkg:binaryData>
    </pkg:part>
  </pkg:package>
`;

export default generateEmbeddedFontsXML;
