#!/usr/bin/env node
/**
 * sync-from-dsh.mjs — regenerate the seven SKILL.md files under skills/ from
 * upstream @deepseek-ai/dsh-skill-office assets, re-applying the pi adaptations.
 *
 * Usage:
 *   npm pack @deepseek-ai/dsh-skill-office@<version>
 *   tar -xzf deepseek-ai-dsh-skill-office-<version>.tgz
 *   node scripts/sync-from-dsh.mjs ./package/assets
 *
 * Every replacement is an exact-match assertion: if upstream rewords a
 * sentence, this script fails loudly with the offending text instead of
 * silently producing a half-adapted skill. Update the REPLACEMENTS table
 * below and commit the change together with the regenerated files.
 *
 * skills/office/references/sidecar.md is a pi-specific rewrite and is NOT
 * managed here — diff it by hand against upstream when syncing.
 */
import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const PKG_ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const SOURCE_COMMENT = "<!-- Skill text ported from @deepseek-ai/dsh-skill-office (MIT, Copyright (c) 2026 DeepSeek): https://github.com/deepseek-ai/deepseek-harness. Unofficial pi packaging in dsh-office-skills; see NOTICE. -->";

const REPLACEMENTS = [
	{
		"file": "office/SKILL.md",
		"replacements": [
			[
				"Call the `skill` tool with the format skill name before creating, reading, or editing Office files. Standalone rendering, PDF export, and recalculation need only their operation skill.",
				"Before creating, reading, or editing Office files, load the format skill for the task by reading its SKILL.md from the sibling directory next to this file. Standalone rendering, PDF export, and recalculation need only their operation skill."
			],
			[
				" Do not reload skills.",
				"\n\nThe sibling skill directories sit next to this file: `../office-docx`, `../office-pptx`, `../office-xlsx`, `../office-render`, `../office-recalculate`, `../office-export-pdf`. Read the chosen skill's SKILL.md before acting. Do not reload a skill you have already read in this session."
			]
		]
	},
	{
		"file": "office-docx/SKILL.md",
		"replacements": [
			[
				"Otherwise call `load_workspace_dependencies` and execute the returned Python path with its bundled libraries. Do not install packages or discover a system Python for the default workflow. If the tool is unavailable, use an already configured environment and report a missing dependency only when it prevents the requested operation.",
				"Otherwise follow `../../python/SETUP.md` (relative to this skill's directory) to select a Python environment with `python-docx` available; it defines the consent-first install for missing libraries. Do not install packages without the user's agreement. If no usable environment can be configured, report a missing dependency only when it prevents the requested operation."
			],
			[
				"Call `present({\"files\":[{\"path\":\"report.docx\"}]})` with the actual final DOCX path. It exposes the current source file without copying or preserving its bytes, so keep that file in place and do not present temporary QA reports unless requested. If `present` is unavailable, provide the final workspace path using the session's supported file delivery method.",
				"Report the actual final DOCX path in the reply so the user can open it from the workspace; keep the file in place and omit temporary QA reports unless requested. If the session provides a file-delivery or presentation mechanism, use it with the same final path."
			]
		]
	},
	{
		"file": "office-xlsx/SKILL.md",
		"replacements": [
			[
				"Otherwise call `load_workspace_dependencies` and run its Python executable with bundled `openpyxl` and `pandas`. Do not install packages or search for a system Python for the default workflow. If the tool is unavailable, use an already configured environment and report a missing dependency only when it prevents the requested operation.",
				"Otherwise follow `../../python/SETUP.md` (relative to this skill's directory) to select a Python environment with `openpyxl` (and `pandas` for analysis) available; it defines the consent-first install for missing libraries. Do not install packages without the user's agreement. If no usable environment can be configured, report a missing dependency only when it prevents the requested operation."
			],
			[
				"Call `present({\"files\":[{\"path\":\"report.xlsx\"}]})` with the actual final workbook path. It exposes the current source file without copying or preserving its bytes, so keep that file in place and omit intermediate scripts and reports unless requested. If `present` is unavailable, use the session's supported file delivery method.",
				"Report the actual final workbook path in the reply so the user can open it from the workspace; keep the file in place and omit intermediate scripts and reports unless requested. If the session provides a file-delivery or presentation mechanism, use it with the same final path."
			]
		]
	},
	{
		"file": "office-pptx/SKILL.md",
		"replacements": [
			[
				"Otherwise call `load_workspace_dependencies` and use its Python executable and bundled presentation libraries. Do not install npm or pip packages or locate a system interpreter for the default workflow. If the tool is unavailable, use an already configured environment and report a missing dependency only when it prevents the requested operation.",
				"Otherwise follow `../../python/SETUP.md` (relative to this skill's directory) to select a Python environment with `python-pptx` available; it defines the consent-first install for missing libraries. Do not install packages without the user's agreement. If no usable environment can be configured, report a missing dependency only when it prevents the requested operation."
			],
			[
				"Call `present({\"files\":[{\"path\":\"report-edited.pptx\"}]})` with the actual final PPTX path. It exposes the current source file without copying or preserving its bytes, so keep that file in place and leave intermediate images and reports out of delivery unless requested. If `present` is unavailable, use the session's supported file delivery method.",
				"Report the actual final PPTX path in the reply so the user can open it from the workspace; keep the file in place and leave intermediate images and reports out of delivery unless requested. If the session provides a file-delivery or presentation mechanism, use it with the same final path."
			]
		]
	},
	{
		"file": "office-render/SKILL.md",
		"replacements": [
			[
				"Use the installed LibreOffice Kit CLI through the shell with the absolute Node and CLI paths supplied below. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`.",
				"Use the installed LibreOffice Kit CLI through the shell. Resolve the absolute Node and CLI paths once per session by running, from this skill's directory, `node ../../scripts/resolve-dsoffice.mjs`; it prints `{ \"libreofficeKit\": { \"node\": \"<node>\", \"cli\": \"<cli>\" } }`. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`."
			],
			[
				"If the Office sidecar is missing, follow [sidecar setup](../office/references/sidecar.md): ask whether the user wants to download it and install only after they agree.",
				"If path resolution fails or the engine is missing, follow [sidecar setup](../office/references/sidecar.md): fix or fetch the engine only after the user agrees."
			],
			[
				"Deliver requested files through `present` when available; keep the files in place. Do not deliver inspection images or reports unless requested.",
				"Deliver requested files by reporting their final paths; keep the files in place. Do not deliver inspection images or reports unless requested."
			]
		]
	},
	{
		"file": "office-recalculate/SKILL.md",
		"replacements": [
			[
				"Use the installed LibreOffice Kit CLI through the shell with the absolute Node and CLI paths supplied below. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`.",
				"Use the installed LibreOffice Kit CLI through the shell. Resolve the absolute Node and CLI paths once per session by running, from this skill's directory, `node ../../scripts/resolve-dsoffice.mjs`; it prints `{ \"libreofficeKit\": { \"node\": \"<node>\", \"cli\": \"<cli>\" } }`. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`."
			],
			[
				"If the Office sidecar is missing, follow [sidecar setup](../office/references/sidecar.md): ask whether the user wants to download it and install only after they agree.",
				"If path resolution fails or the engine is missing, follow [sidecar setup](../office/references/sidecar.md): fix or fetch the engine only after the user agrees."
			],
			[
				"Verify the saved result with openpyxl, using `load_workspace_dependencies` for the Python path unless a project environment is required.",
				"Verify the saved result with openpyxl, using the Python environment from `../../python/SETUP.md` unless a project environment is required."
			],
			[
				"Deliver the verified output through `present` when available, keeping the file in place.",
				"Deliver the verified output by reporting its final path, keeping the file in place; use the session's file-delivery mechanism if it provides one."
			]
		]
	},
	{
		"file": "office-export-pdf/SKILL.md",
		"replacements": [
			[
				"Use the installed LibreOffice Kit CLI through the shell with the absolute Node and CLI paths supplied below. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`.",
				"Use the installed LibreOffice Kit CLI through the shell. Resolve the absolute Node and CLI paths once per session by running, from this skill's directory, `node ../../scripts/resolve-dsoffice.mjs`; it prints `{ \"libreofficeKit\": { \"node\": \"<node>\", \"cli\": \"<cli>\" } }`. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`."
			],
			[
				"If the Office sidecar is missing, follow [sidecar setup](../office/references/sidecar.md): ask whether the user wants to download it and install only after they agree.",
				"If path resolution fails or the engine is missing, follow [sidecar setup](../office/references/sidecar.md): fix or fetch the engine only after the user agrees."
			],
			[
				"Deliver the requested PDF through `present` when available, keeping the file in place.",
				"Deliver the requested PDF by reporting its final path, keeping the file in place; use the session's file-delivery mechanism if it provides one."
			]
		]
	}
];

function replaceOnce(text, find, replace, label) {
	const first = text.indexOf(find);
	if (first < 0) {
		throw new Error(`[${label}] upstream text no longer matches — upstream changed? Update the REPLACEMENTS table:\n  ${find.slice(0, 120)}...`);
	}
	if (text.indexOf(find, first + 1) >= 0) {
		throw new Error(`[${label}] multiple hits for replacement — table needs a more specific anchor:\n  ${find.slice(0, 120)}...`);
	}
	return text.slice(0, first) + replace + text.slice(first + find.length);
}

const srcRoot = process.argv[2];
if (!srcRoot) {
	console.error("usage: node scripts/sync-from-dsh.mjs <path-to-dsh-skill-office-assets>");
	process.exit(1);
}

for (const entry of REPLACEMENTS) {
	const label = entry.file;
	let raw = readFileSync(join(srcRoot, entry.file), "utf8");
	for (const [find, replace] of entry.replacements) {
		raw = replaceOnce(raw, find, replace, label);
	}
	// frontmatter: add the license field (upstream files carry name + description only)
	raw = replaceOnce(raw, "\n---\n", "\nlicense: MIT\n---\n", label + " (frontmatter)");
	const out = raw.trimEnd() + "\n\n" + SOURCE_COMMENT + "\n";
	writeFileSync(join(PKG_ROOT, "skills", entry.file), out);
	console.log("synced: skills/" + entry.file);
}
console.log("done. Remember: skills/office/references/sidecar.md is a manual, pi-specific rewrite — diff it against upstream by hand.");
