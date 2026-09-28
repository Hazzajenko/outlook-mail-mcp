import { createRequire } from "node:module";

// Resolves to the package root from both src/ under tsx and dist/ after build.
// release-please bumps package.json, so this is the only place the version lives.
const pkg: unknown = createRequire(import.meta.url)("../package.json");

export const VERSION = (pkg as { version: string }).version;
