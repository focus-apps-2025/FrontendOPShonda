import JSZip from "jszip";

export interface ImagePlacement {
  base64: string;
  type: "png" | "jpeg";
  fromCol: number;
  fromColOff: number;
  fromRow: number;
  fromRowOff: number;
  toCol: number;
  toColOff: number;
  toRow: number;
  toRowOff: number;
  cx: number; // Width in EMU (1 inch = 914400 EMU)
  cy: number; // Height in EMU
  anchor: "one" | "two";
  name?: string;
}

export async function injectImagesIntoXlsx(
  xlsxArrayBuffer: ArrayBuffer,
  images: ImagePlacement[]
): Promise<Blob> {
  const zip = await JSZip.loadAsync(xlsxArrayBuffer);

  const validImages = images.filter((img) => img.base64 && img.base64.length > 10);
  if (validImages.length === 0) {
    return zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
  }

  // 1. Add each image to xl/media/
  for (let i = 0; i < validImages.length; i++) {
    const img = validImages[i];
    const ext = img.type === "jpeg" ? "jpg" : "png";
    zip.file(`xl/media/image${i + 1}.${ext}`, base64ToUint8Array(img.base64));
  }

  // 2. Add drawing XML
  zip.file("xl/drawings/drawing1.xml", buildDrawingXml(validImages));

  // 3. Add drawing relationships
  zip.file("xl/drawings/_rels/drawing1.xml.rels", buildDrawingRelsXml(validImages));

  // 4. Patch sheet1.xml — add drawing reference
  const sheet1Path = "xl/worksheets/sheet1.xml";
  let sheet1 = await zip.file(sheet1Path)!.async("string");

  if (!sheet1.includes('xmlns:r=')) {
    sheet1 = sheet1.replace(
      "<worksheet ",
      '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    );
  }
  if (!sheet1.includes("<drawing ")) {
    sheet1 = sheet1.replace("</worksheet>", '<drawing r:id="rId_drawing"/></worksheet>');
  }
  zip.file(sheet1Path, sheet1);

  // 5. Add/patch sheet1.xml.rels
  const sheetRelsPath = "xl/worksheets/_rels/sheet1.xml.rels";
  const sheetRelsFile = zip.file(sheetRelsPath);
  let sheetRels: string;

  if (sheetRelsFile) {
    sheetRels = await sheetRelsFile.async("string");
    if (!sheetRels.includes("rId_drawing")) {
      sheetRels = sheetRels.replace(
        "</Relationships>",
        `  <Relationship Id="rId_drawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`
      );
    }
  } else {
    sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId_drawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
  }
  zip.file(sheetRelsPath, sheetRels);

  // 6. Patch [Content_Types].xml
  const contentTypesPath = "[Content_Types].xml";
  let contentTypes = await zip.file(contentTypesPath)!.async("string");

  if (!contentTypes.includes("drawing1.xml")) {
    contentTypes = contentTypes.replace(
      "</Types>",
      `  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`
    );
  }
  if (!contentTypes.includes('Extension="png"')) {
    contentTypes = contentTypes.replace("</Types>", `  <Default Extension="png" ContentType="image/png"/>\n</Types>`);
  }
  if (!contentTypes.includes('Extension="jpg"') && !contentTypes.includes('Extension="jpeg"')) {
    contentTypes = contentTypes.replace("</Types>", `  <Default Extension="jpg" ContentType="image/jpeg"/>\n</Types>`);
  }
  zip.file(contentTypesPath, contentTypes);

  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE",
  });
}

function base64ToUint8Array(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function buildDrawingXml(images: ImagePlacement[]): string {
  const anchors = images.map((img, i) => {
    const name = img.name || `Image ${i + 1}`;
    const rId = `rId${i + 1}`;
    const id = i + 2;

    if (img.anchor === "one") {
      return `<xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>${img.fromCol}</xdr:col><xdr:colOff>${img.fromColOff}</xdr:colOff>
      <xdr:row>${img.fromRow}</xdr:row><xdr:rowOff>${img.fromRowOff}</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="${img.cx}" cy="${img.cy}"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="${id}" name="${name}"/>
        <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}">
          <a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst>
        </a:blip>
        <a:srcRect/><a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr bwMode="auto">
        <a:xfrm><a:off x="${img.fromColOff}" y="${img.fromRowOff}"/><a:ext cx="${img.cx}" cy="${img.cy}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>`;
    } else {
      return `<xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>${img.fromCol}</xdr:col><xdr:colOff>${img.fromColOff}</xdr:colOff>
      <xdr:row>${img.fromRow}</xdr:row><xdr:rowOff>${img.fromRowOff}</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>${img.toCol}</xdr:col><xdr:colOff>${img.toColOff}</xdr:colOff>
      <xdr:row>${img.toRow}</xdr:row><xdr:rowOff>${img.toRowOff}</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="${id}" name="${name}"/>
        <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}">
          <a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst>
        </a:blip>
        <a:srcRect/><a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr bwMode="auto">
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${img.cx}" cy="${img.cy}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`;
    }
  }).join("\n");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
${anchors}
</xdr:wsDr>`;
}

function buildDrawingRelsXml(images: ImagePlacement[]): string {
  const rels = images.map((img, i) => {
    const ext = img.type === "jpeg" ? "jpg" : "png";
    return `  <Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${i + 1}.${ext}"/>`;
  }).join("\n");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels}
</Relationships>`;
}