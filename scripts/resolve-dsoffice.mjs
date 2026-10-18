#!/usr/bin/env node
/**
 * dsh-office-skills engine path resolver.
 *
 * The office-render / office-recalculate / office-export-pdf skills invoke the
 * LibreOffice Kit CLI as "<node>" "<cli>" with absolute paths. This script
 * prints those paths as JSON so the model never guesses them:
 *
 *   { "libreofficeKit": { "node": "<absolute node>", "cli": "<absolute cli.js>" } }
 *
 * Exit codes: 0 = resolved and printed; 1 = dependency missing or unreadable
 * (stderr carries the recovery hint pointing at office/references/sidecar.md).
 */
import { createRequire } from "node:module";
import { statSync } from "node:fs";

const require = createRequire(import.meta.url);

let cli;
try {
	// The kit package restricts deep imports via its "exports" whitelist; the
	// public subpath is "./cli" (mapped to lib/cli.js). Older versions without
	// an "exports" map fall back to the direct file path.
	try {
		cli = require.resolve("@deepseek-ai/libreoffice-kit/cli");
	} catch {
		cli = require.resolve("@deepseek-ai/libreoffice-kit/lib/cli.js");
	}
} catch {
	process.stderr.write('resolve-dsoffice: LibreOffice Kit is not installed under this package. Follow office/references/sidecar.md: run "pi update --extensions" (npm/git package installs) or "npm install" in the package directory (local installs), then retry.\n');
	process.exit(1);
}
try {
	if (!statSync(cli).isFile()) throw new Error("not a file");
} catch {
	process.stderr.write(`resolve-dsoffice: CLI entry is missing or unreadable: ${cli}\n`);
	process.exit(1);
}
const [major, minor] = process.versions.node.split(".").map(Number);
if (major < 22 || (major === 22 && minor < 19)) {
	// The engine declares node >=22.19 and pi requires the same, so this is
	// informational only; never a silent pass on an unsupported runtime.
	process.stderr.write(`resolve-dsoffice: Node ${process.versions.node} is older than the engine's declared minimum (22.19); commands may fail.\n`);
}
process.stdout.write(JSON.stringify({ libreofficeKit: { node: process.execPath, cli } }, null, 2) + "\n");
