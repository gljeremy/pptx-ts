// Moving an element into a tree strips namespace declarations its new ancestors already make, on
// every element of the moved subtree, and remaps prefixes to the ancestors' (lxml's
// moveNodeToDocument). Expected strings and prefixes are lxml 5.3's output for the same steps.
import { expect, it } from "vitest";
import { parseXml } from "../src/oxml/parser.js";
import { tostring } from "../src/oxml/xml.js";

const cases: [string, string, string, "append" | "insert" | "addnext", string, (string | null)[]][] = [
  ["nested same prefix", '<a:root xmlns:a="http://a" xmlns:b="http://b"/>', '<a:x xmlns:a="http://a"><b:y xmlns:b="http://b" b:id="1"/></a:x>', "append",
    '<a:root xmlns:a="http://a" xmlns:b="http://b"><a:x><b:y b:id="1"/></a:x></a:root>', ["a", "b"]],
  ["nested other prefix", '<a:root xmlns:a="http://a" xmlns:b="http://b"/>', '<a:x xmlns:a="http://a"><q:y xmlns:q="http://b" q:id="1"><q:z/></q:y></a:x>', "append",
    '<a:root xmlns:a="http://a" xmlns:b="http://b"><a:x><b:y b:id="1"><b:z/></b:y></a:x></a:root>', ["a", "b", "b"]],
  ["not in scope kept", '<a:root xmlns:a="http://a"/>', '<a:x xmlns:a="http://a"><c:y xmlns:c="http://c"><c:z xmlns:c="http://c"/></c:y></a:x>', "append",
    '<a:root xmlns:a="http://a"><a:x><c:y xmlns:c="http://c"><c:z/></c:y></a:x></a:root>', ["a", "c", "c"]],
  ["default ns", '<root xmlns="http://a"/>', '<a:x xmlns:a="http://a"><a:y/></a:x>', "append",
    '<root xmlns="http://a"><x><y/></x></root>', [null, null]],
  ["insert", '<a:root xmlns:a="http://a" xmlns:b="http://b"><a:k/></a:root>', '<a:x xmlns:a="http://a"><b:y xmlns:b="http://b"/></a:x>', "insert",
    '<a:root xmlns:a="http://a" xmlns:b="http://b"><a:x><b:y/></a:x><a:k/></a:root>', ["a", "b"]],
  ["addnext", '<a:root xmlns:a="http://a" xmlns:b="http://b"><a:k/></a:root>', '<a:x xmlns:a="http://a"><b:y xmlns:b="http://b"/></a:x>', "addnext",
    '<a:root xmlns:a="http://a" xmlns:b="http://b"><a:k/><a:x><b:y/></a:x></a:root>', ["a", "b"]],
];

it.each(["append", "extend", "insert", "replace", "addprevious", "addnext"] as const)(
  "keeps a child's namespace shadowing out of sibling scope on %s (lxml contract)",
  (how) => {
    const parent = parseXml('<a:root xmlns:a="http://a"><a:k/></a:root>');
    const child = parseXml('<q:x xmlns:q="http://a"><q:y xmlns:q="http://b"><q:n/></q:y><q:z/></q:x>');
    const anchor = parent._children[0];
    if (how === "append") parent.append(child);
    else if (how === "extend") parent.extend([child]);
    else if (how === "insert") parent.insert(0, child);
    else if (how === "replace") parent.replace(anchor, child);
    else if (how === "addprevious") anchor.addprevious(child);
    else anchor.addnext(child);
    expect([...child.iter()].map((element) => element.prefix)).toEqual(["a", "q", "q", "a"]);
    expect(child.xpath("name(.)")).toBe("a:x");
    expect(child.xpath("name(./*[1])")).toBe("q:y");
    expect(child.xpath("name(./*[1]/*[1])")).toBe("q:n");
    expect(child.xpath("name(./*[2])")).toBe("a:z");
    const moved = '<a:x><q:y xmlns:q="http://b"><q:n/></q:y><a:z/></a:x>';
    const children = how === "replace" ? moved : how === "insert" || how === "addprevious" ? `${moved}<a:k/>` : `<a:k/>${moved}`;
    expect(tostring(parent)).toBe(`<a:root xmlns:a="http://a">${children}</a:root>`);
  },
);

for (const [name, parentXml, childXml, how, expected, prefixes] of cases) {
  it(`strips redundant namespace declarations on ${how}: ${name}`, () => {
    const parent = parseXml(parentXml);
    const child = parseXml(childXml);
    if (how === "append") parent.append(child);
    else if (how === "insert") parent.insert(0, child);
    else parent._children[0].addnext(child);
    expect(tostring(parent)).toBe(expected);
    expect([...child.iter()].map((e) => e.prefix)).toEqual(prefixes);
  });
}
