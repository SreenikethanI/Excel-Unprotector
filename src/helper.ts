import type JSZip from "jszip";

/** Loads the path from the zipfile, and throws an error if it doesn't exist. */
export function file(zipFile: JSZip, path: string) {
  const file = zipFile.file(path);
  if (!file) throw new Error(`${path} doesn't exist.`);
  return file;
}

/** Helper function to parse the given XML string. */
export function xml(xmlString: string): XMLDocument {
  return new DOMParser().parseFromString(xmlString, "application/xml");
}

/** Helper function to stringify the given XML document. */
export function xmlToString(xml: XMLDocument): string {
  return new XMLSerializer().serializeToString(xml);
}

/** Generate the path to the rels file for the given file path. */
export function getRelsPath(path: string): string {
  path = "/" + path;
  const i = path.lastIndexOf("/");
  return `${path.substring(0, i)}/_rels${path.substring(i)}.rels`.substring(1);
}

/**
 * Resolve a path mentioned in a rels file.
 * @param targetPath The Target path mentioned in the rels file.
 * @param originalPath The original XML file's path to which the rels file belongs.
 * @example
 * resolveRelsTarget("worksheets/sheet1.xml", "xl/workbook.xml");
 * // returns "xl/worksheets/sheet1.xml"
 */
export function resolveRelsTarget(
  targetPath: string,
  originalPath: string,
): string {
  originalPath = "/" + originalPath;
  const i = originalPath.lastIndexOf("/");
  return `${originalPath.substring(0, i)}/${targetPath}`.substring(1);
}
