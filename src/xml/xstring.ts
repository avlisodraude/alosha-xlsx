// ST_Xstring escaping is distinct from XML entity escaping. Apply it only to OOXML string fields,
// never to formulas, relationship URLs, or ordinary XML metadata.
export function encodeXString(value: string): string {
  return value.replace(/_(?=x[0-9a-fA-F]{4}_)|\r/g, (match) =>
    match === "_" ? "_x005F_" : "_x000D_",
  );
}

export function decodeXString(value: string): string {
  // One pass is essential: _x005F_x0041_ represents the literal _x0041_, not A.
  return value.replace(/_x([0-9a-fA-F]{4})_/g, (_match, hex: string) =>
    String.fromCharCode(Number.parseInt(hex, 16)),
  );
}

export function isXStringAttribute(tag: string, key: string): boolean {
  const local = tag.slice(tag.indexOf(":") + 1);
  return (
    key === "formatCode" ||
    key === "name" ||
    (key === "val" && (local === "name" || local === "rFont")) ||
    ["prompt", "promptTitle", "error", "errorTitle", "tooltip"].includes(key)
  );
}

export function isXStringElement(tag: string): boolean {
  const local = tag.slice(tag.indexOf(":") + 1);
  return local === "t" || local === "author" || local === "v";
}
