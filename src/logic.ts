import JSZip from "jszip";
import {
  xpathNsResolver,
  CONTENT_TYPES,
  WORKBOOK,
  WORKBOOK_RELS,
  SHEET,
  XLSX_MIME_TYPE,
  ZIP_COMMENT,
} from "./constants.ts";
import {
  file,
  xml,
  getRelsPath,
  xmlToString,
  resolveRelsTarget,
} from "./helper.ts";

/** Helper function to remove the given XML node, and write it to the zipfile.
 * @param path The file path inside the zipfile.
 * @param xpath The XPath that selects the XML node to remove.
 * @returns `true` if node found and removed, `false` if node doesn't exist.
 */
async function removeNodeAndUpdate(
  path: string,
  xpath: string,
  zipfile: JSZip,
): Promise<boolean> {
  const documentFile = file(zipfile, path);
  const document = xml(await documentFile.async("string"));
  const protectionNode = document.evaluate(
    xpath,
    document,
    xpathNsResolver,
    XPathResult.FIRST_ORDERED_NODE_TYPE,
  ).singleNodeValue;
  if (!protectionNode) return false;
  protectionNode.parentNode?.removeChild(protectionNode);
  zipfile.file(path, xmlToString(document), {
    base64: false,
    binary: false,
    comment: ZIP_COMMENT,
  });
  return true;
}

/**
 * Unprotect the following in the given Excel file:
 * - Sheet protection
 * - Workbook structure protection (not the same as file encryption)
 * @param data The Excel file.
 */
export async function unprotectWorkbook(
  data: Parameters<JSZip["loadAsync"]>[0],
  compress?: boolean,
): Promise<Blob> {
  //#region load zip file
  const zipfile = await JSZip.loadAsync(data);
  // TODO: show progress
  console.log("Query: Parse zip file done.");
  //#endregion

  //#region get path to xl/workbook.xml
  const contentTypesFile = file(zipfile, CONTENT_TYPES.PATH);
  const contentTypes = xml(await contentTypesFile.async("string"));
  const workbookPath =
    contentTypes
      .evaluate(
        CONTENT_TYPES.XPATH_WORKBOOK_PATH,
        contentTypes,
        xpathNsResolver,
        XPathResult.STRING_TYPE,
      )
      .stringValue.replace(/^\//, "") || WORKBOOK.PATH_FALLBACK;
  // TODO: show progress
  console.log("Query: workbook.xml path retrieved.");
  //#endregion

  //#region get rIds of sheets
  const workbookFile = file(zipfile, workbookPath);
  const workbook = xml(await workbookFile.async("string"));
  const sheetRIdsIterator = workbook.evaluate(
    WORKBOOK.XPATH_SHEETS,
    workbook,
    xpathNsResolver,
    XPathResult.ORDERED_NODE_ITERATOR_TYPE,
  );
  const sheetNames: Map<string, string> = new Map();
  while (true) {
    const node = sheetRIdsIterator.iterateNext() as Element | null;
    if (!node) break;
    const rId = node.getAttributeNS(xpathNsResolver("r"), "id");
    const name = node.getAttribute("name");
    const sheetId = node.getAttribute("sheetId");
    if (rId && (name || sheetId)) sheetNames.set(rId, name || sheetId || "");
  }
  // TODO: show progress
  console.log("Query: sheet rIds retrieved.");
  //#endregion

  //#region get file paths of sheets
  const workbookRelsFile = file(zipfile, getRelsPath(workbookPath));
  const workbookRels = xml(await workbookRelsFile.async("string"));
  const sheetRelsIterator = workbookRels.evaluate(
    WORKBOOK_RELS.XPATH_SHEETS_RELS,
    workbookRels,
    xpathNsResolver,
    XPathResult.ORDERED_NODE_ITERATOR_TYPE,
  );
  const sheetPaths: Map<string, string> = new Map();
  while (true) {
    const node = sheetRelsIterator.iterateNext() as Element | null;
    if (!node) break;
    let [rId, sheetPath] = [
      node.getAttribute("Id"),
      node.getAttribute("Target"),
    ];
    if (rId && sheetPath && sheetNames.has(rId)) {
      sheetPaths.set(rId, resolveRelsTarget(sheetPath, workbookPath));
    }
  }
  // TODO: show progress
  console.log("Query: sheet xml paths retrieved.");
  //#endregion

  //#region remove protection
  if (
    await removeNodeAndUpdate(
      workbookPath,
      WORKBOOK.XPATH_PROTECTION_NODE,
      zipfile,
    )
  ) {
    // TODO: show progress
    console.log("Action: workbook protection removed.");
  } else {
    // TODO: show progress
    console.log("Action: workbook protection not found.");
  }

  for (const [rId, sheetPath] of sheetPaths) {
    const sheetName = sheetNames.get(rId) || `(rId: ${rId})`;
    if (
      await removeNodeAndUpdate(
        sheetPath,
        SHEET.XPATH_PROTECTION_NODE,
        zipfile,
      )
    ) {
      // TODO: show progress
      console.log(`Action: sheet protection for "${sheetName}" removed.`);
    } else {
      // TODO: show progress
      console.log(`Action: sheet protection for "${sheetName}" not found.`);
    }
  }
  //#endregion

  return await zipfile.generateAsync({
    type: "blob",
    platform: "DOS",
    mimeType: XLSX_MIME_TYPE,
    comment: ZIP_COMMENT,
    ...(compress
      ? { compression: "DEFLATE", compressionOptions: { level: 9 } }
      : { compression: "STORE" }),
  });
}
