import { defineConfig, type Plugin } from "vitest/config";

/**
 * Test-only transform that makes every top-level binding of a src module patchable, the way
 * python's `mock.patch("pptx.opc.package.PartFactory")` patches a module global. Each module
 * gains `__pyBindings = { get(name), set(name, value) }`; ES live bindings carry a patched
 * value to every importer. `export const` becomes `export let` so it can be reassigned.
 */
function patchableModules(): Plugin {
  return {
    name: "patchable-modules",
    enforce: "post",
    transform(code, id) {
      if (!/[\\/]src[\\/].*\.ts$/.test(id) || id.includes("node_modules")) return null;
      const names = new Set<string>();
      const out = code.replace(/^(export\s+)?(const|let|var|class|function\*?|async function)\s+([A-Za-z_$][\w$]*)/gm, (m, exp, kind, name) => {
        names.add(name);
        return kind === "const" ? `${exp ?? ""}let ${name}` : m;
      });
      // where imported names come from: { localName: [specifier, importedName] }
      const imports: Record<string, [string, string]> = {};
      for (const m of code.matchAll(/^import\s*\{([^}]*)\}\s*from\s*["']([^"']+)["']/gm)) {
        for (const part of m[1].split(",")) {
          const [orig, alias] = part.trim().replace(/^type\s+/, "").split(/\s+as\s+/);
          if (orig) imports[(alias ?? orig).trim()] = [m[2], orig.trim()];
        }
      }
      if (!names.size && !Object.keys(imports).length) return null;
      const cases = [...names];
      const getter = cases.map((n) => `case ${JSON.stringify(n)}: return ${n};`).join(" ");
      const setter = cases.map((n) => `case ${JSON.stringify(n)}: ${n} = v; return true;`).join(" ");
      return {
        code:
          `${out}\nexport const __pyBindings = { id: ${JSON.stringify(id.replace(/\\/g, "/"))}, names: ${JSON.stringify(cases)}, imports: ${JSON.stringify(imports)}, ` +
          `get(n) { switch (n) { ${getter} } }, set(n, v) { switch (n) { ${setter} } return false; } };\n`,
        map: null,
      };
    },
  };
}

export default defineConfig({
  plugins: [patchableModules()],
  test: {
    include: ["test/**/*.test.ts"],
    setupFiles: ["test/setup.ts"],
  },
});
