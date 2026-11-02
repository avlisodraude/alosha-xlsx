import type { XmlEvent } from "./xml-reader.js";

const ATTR_PREFIXES: Record<string, string> = {
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r",
  "http://www.w3.org/XML/1998/namespace": "xml",
  "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac": "x14ac",
};

/** Resolve namespace declarations by scope, then present canonical names to a part reader.
 * Unknown element namespaces are skipped as whole subtrees, including any familiar descendants.
 * Namespace-free documents remain supported for the existing lightweight part-reader API. */
export function* namespaceEvents(
  events: Iterable<XmlEvent>,
  allowed: readonly string[],
  onSkipped?: (name: string, namespace: string | undefined) => void,
): Generator<XmlEvent> {
  const scopes: Array<{ bindings: Record<string, string>; blocked: boolean; name: string }> = [];
  let legacy = false;
  let rootSeen = false;
  for (const event of events) {
    if (event.kind === "open") {
      const parent = scopes[scopes.length - 1];
      const bindings: Record<string, string> = Object.create(parent?.bindings ?? null);
      for (const [key, value] of Object.entries(event.attrs)) {
        if (key === "xmlns") bindings[""] = value;
        else if (key.startsWith("xmlns:")) bindings[key.slice(6)] = value;
      }
      const split = event.name.indexOf(":");
      const prefix = split < 0 ? "" : event.name.slice(0, split);
      const local = event.name.slice(split + 1);
      const uri = bindings[prefix];
      const first = !rootSeen;
      if (!rootSeen) {
        legacy = !uri && split < 0;
        rootSeen = true;
      }
      const blocked =
        parent?.blocked === true ||
        (uri ? !allowed.includes(uri) || (first && uri !== allowed[0]) : !legacy);
      const name = uri ? local : event.name;
      if (!event.selfClosing) scopes.push({ bindings, blocked, name });
      if (blocked) {
        if (!parent?.blocked) onSkipped?.(event.name, uri);
        continue;
      }
      const attrs: Record<string, string> = {};
      for (const [key, value] of Object.entries(event.attrs)) {
        if (key === "xmlns" || key.startsWith("xmlns:")) continue;
        const colon = key.indexOf(":");
        if (colon < 0) {
          attrs[key] = value;
          continue;
        }
        const attrUri = key.startsWith("xml:")
          ? "http://www.w3.org/XML/1998/namespace"
          : bindings[key.slice(0, colon)];
        const canonical =
          attrUri && Object.hasOwn(ATTR_PREFIXES, attrUri) ? ATTR_PREFIXES[attrUri] : undefined;
        if (canonical) attrs[`${canonical}:${key.slice(colon + 1)}`] = value;
        else if (!attrUri && legacy) attrs[key] = value;
        // Unrecognized namespaced attributes cannot masquerade as unqualified model fields.
      }
      yield { ...event, name, attrs };
    } else if (event.kind === "close") {
      const scope = scopes.pop();
      if (scope && !scope.blocked) yield { kind: "close", name: scope.name };
    } else if (!scopes[scopes.length - 1]?.blocked) {
      yield event;
    }
  }
}
