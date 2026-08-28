/**
 * Builds a minimal but realistic .pptx (raw OOXML, zipped) with:
 * - 16:9 slide size (12192000 x 6858000 EMU)
 * - slide background #1E2430
 * - title text box (44pt bold white, positioned at 914400,685800)
 * - body text box: bullet para (centered, 18pt #FFC000) + plain para
 * - filled rectangle #C43E1C (no text)
 * - embedded 1x1 PNG picture at a fixed position
 */
import JSZip from 'jszip';

const EMU = (inch) => Math.round(inch * 914400);

// 1x1 red PNG
const PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const XMLH = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function sp(id, name, xIn, yIn, wIn, hIn, inner) {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${EMU(xIn)}" y="${EMU(yIn)}"/><a:ext cx="${EMU(wIn)}" cy="${EMU(hIn)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${inner.spPr ?? ''}</p:spPr>` +
    (inner.txBody ? `<p:txBody><a:bodyPr/><a:lstStyle/>${inner.txBody}</p:txBody>` : '') +
    `</p:sp>`
  );
}

const titleSp = sp(2, 'Title', 1, 0.75, 9, 1.44, {
  txBody:
    '<a:p><a:r><a:rPr lang="en-US" sz="4400" b="1">' +
    '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>' +
    '<a:t>Quarterly Review</a:t></a:r></a:p>',
});

const bodySp = sp(3, 'Content', 1, 2.4, 8, 2.6, {
  txBody:
    '<a:p><a:pPr algn="ctr"><a:buChar char="•"/></a:pPr>' +
    '<a:r><a:rPr lang="en-US" sz="1800"><a:solidFill><a:srgbClr val="FFC000"/></a:solidFill></a:rPr>' +
    '<a:t>Growth in every region</a:t></a:r></a:p>' +
    '<a:p><a:pPr><a:buNone/></a:pPr>' +
    '<a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr>' +
    '<a:t>Plain paragraph with scheme color</a:t></a:r></a:p>',
});

const rectSp = sp(4, 'Band', 0, 5.0, 10, 0.63, {
  spPr: '<a:solidFill><a:srgbClr val="C43E1C"/></a:solidFill>',
});

const picXml =
  `<p:pic><p:nvPicPr><p:cNvPr id="5" name="Logo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="${EMU(8.4)}" y="${EMU(0.3)}"/><a:ext cx="${EMU(1.2)}" cy="${EMU(1.2)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;

const slide1 =
  XMLH +
  '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="1E2430"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>' +
  '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
  titleSp + bodySp + rectSp + picXml +
  '</p:spTree></p:cSld></p:sld>';

export function buildPptx() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    XMLH +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
      '</Types>',
  );
  zip.file(
    '_rels/.rels',
    XMLH +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
      '</Relationships>',
  );
  zip.file(
    'ppt/presentation.xml',
    XMLH +
      '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
  );
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    XMLH +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
      '</Relationships>',
  );
  zip.file('ppt/slides/slide1.xml', slide1);
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    XMLH +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>' +
      '</Relationships>',
  );
  zip.file('ppt/media/image1.png', PNG_B64, { base64: true });
  return zip.generateAsync({ type: 'arraybuffer' });
}
