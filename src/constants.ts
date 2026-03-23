export const XLSX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export const ZIP_COMMENT =
  "Unprotected using https://sreenikethani.github.io/Excel-Unprotector/";

/** Values for `[Content_Types].xml` file. */
export const CONTENT_TYPES = {
  PATH: "[Content_Types].xml",
  NS_PREFIX: "contentTypesNs",
  NS_URI: "http://schemas.openxmlformats.org/package/2006/content-types",
  XPATH_WORKBOOK_PATH: `/contentTypesNs:Types/contentTypesNs:Override[@ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"]/@PartName`,
} as const;

/** Values for `xl/workbook.xml` file. */
export const WORKBOOK = {
  PATH_FALLBACK: "xl/workbook.xml",
  NS_PREFIX: "workbookNs",
  NS_URI: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  XPATH_SHEETS: `/workbookNs:workbook/workbookNs:sheets/workbookNs:sheet`,
  XPATH_PROTECTION_NODE: `/workbookNs:workbook/workbookNs:workbookProtection`,
} as const;

/** Values for `xl/_rels/workbook.xml.rels` file. */
export const WORKBOOK_RELS = {
  NS_PREFIX: "workbookRelsNs",
  NS_URI: "http://schemas.openxmlformats.org/package/2006/relationships",
  XPATH_SHEETS_RELS: `/workbookRelsNs:Relationships/workbookRelsNs:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"]`,
} as const;

/** Values for `xl/worksheets/sheet?.xml` file. */
export const SHEET = {
  NS_PREFIX: "sheetNs",
  NS_URI: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  XPATH_PROTECTION_NODE: `/sheetNs:worksheet/sheetNs:sheetProtection`,
} as const;

/** Callback to resolve XML Namespace. */
export function xpathNsResolver(prefix: string | null): string | null {
  switch (prefix) {
    case "r":
      return "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    case CONTENT_TYPES.NS_PREFIX:
      return CONTENT_TYPES.NS_URI;
    case WORKBOOK.NS_PREFIX:
      return WORKBOOK.NS_URI;
    case WORKBOOK_RELS.NS_PREFIX:
      return WORKBOOK_RELS.NS_URI;
    case SHEET.NS_PREFIX:
      return SHEET.NS_URI;
    default:
      return null;
  }
}
