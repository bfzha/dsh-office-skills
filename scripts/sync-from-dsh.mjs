#!/usr/bin/env node
/**
 * Re-port the office skills from a published @deepseek-ai/dsh-skill-office version.
 *
 * Modes:
 *   node scripts/sync-from-dsh.mjs --check [--version X]          report baseline vs newest published
 *   node scripts/sync-from-dsh.mjs --apply [--version X]          re-port from newest published
 *   node scripts/sync-from-dsh.mjs <assets-dir> [--version X]     re-port from a local extraction
 *
 * The assets-dir form ports straight from an extracted tarball, e.g. the copy
 * bundled with DSH itself:
 *   <dsh>/node_modules/@deepseek-ai/dsh-skill-office/assets
 *
 * The port is a set of exact-text adaptations (RULES below). Every rule must match
 * exactly once; when upstream wording changes the script fails loudly instead of
 * silently dropping an adaptation. Only skills/ and upstream.json are written;
 * sidecar.md, python/ and scripts/ are hand-maintained.
 */
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const PACKAGE = "@deepseek-ai/dsh-skill-office";
const REGISTRY = "https://registry.npmjs.org/";
const ROOT = dirname(dirname(fileURLToPath(import.meta.url)));
const UPSTREAM_JSON = join(ROOT, "upstream.json");
const SKILLS_DIR = join(ROOT, "skills");

const SKILL_NAMES = ["office","office-docx","office-pptx","office-xlsx","office-render","office-recalculate","office-export-pdf"];

const RULES = {
  "office": [
    [
      "Call the `skill` tool with the format skill name before creating, reading, or editing Office files. Standalone rendering, PDF export, and recalculation need only their operation skill.",
      "Before creating, reading, or editing Office files, load the format skill for the task by reading its SKILL.md from the sibling directory next to this file. Standalone rendering, PDF export, and recalculation need only their operation skill.",
      "skill tool -> sibling SKILL.md loading"
    ],
    [
      "unless visual inspection is requested. Do not reload skills.",
      "unless visual inspection is requested.\n\nThe sibling skill directories sit next to this file: `../office-docx`, `../office-pptx`, `../office-xlsx`, `../office-render`, `../office-recalculate`, `../office-export-pdf`. Read the chosen skill's SKILL.md before acting. Do not reload a skill you have already read in this session.",
      "add sibling-directory map"
    ]
  ],
  "office-docx": [
    [
      "Otherwise call `load_workspace_dependencies` and execute the returned Python path with its bundled libraries. Do not install packages or discover a system Python for the default workflow. If the tool is unavailable, use an already configured environment and report a missing dependency only when it prevents the requested operation.",
      "Otherwise follow `../../python/SETUP.md` (relative to this skill's directory) to select a Python environment with `python-docx` available; it defines the consent-first install for missing libraries. Do not install packages without the user's agreement. If no usable environment can be configured, report a missing dependency only when it prevents the requested operation.",
      "load_workspace_dependencies -> python/SETUP.md"
    ],
    [
      "Call `present({\"files\":[{\"path\":\"report.docx\"}]})` with the actual final DOCX path. It exposes the current source file without copying or preserving its bytes, so keep that file in place and do not present temporary QA reports unless requested. If `present` is unavailable, provide the final workspace path using the session's supported file delivery method.",
      "Report the actual final DOCX path in the reply so the user can open it from the workspace; keep the file in place and omit temporary QA reports unless requested. If the session provides a file-delivery or presentation mechanism, use it with the same final path.",
      "present -> report final path"
    ]
  ],
  "office-xlsx": [
    [
      "Otherwise call `load_workspace_dependencies` and run its Python executable with bundled `openpyxl` and `pandas`. Do not install packages or search for a system Python for the default workflow. If the tool is unavailable, use an already configured environment and report a missing dependency only when it prevents the requested operation.",
      "Otherwise follow `../../python/SETUP.md` (relative to this skill's directory) to select a Python environment with `openpyxl` (and `pandas` for analysis) available; it defines the consent-first install for missing libraries. Do not install packages without the user's agreement. If no usable environment can be configured, report a missing dependency only when it prevents the requested operation.",
      "load_workspace_dependencies -> python/SETUP.md"
    ],
    [
      "Call `present({\"files\":[{\"path\":\"report.xlsx\"}]})` with the actual final workbook path. It exposes the current source file without copying or preserving its bytes, so keep that file in place and omit intermediate scripts and reports unless requested. If `present` is unavailable, use the session's supported file delivery method.",
      "Report the actual final workbook path in the reply so the user can open it from the workspace; keep the file in place and omit intermediate scripts and reports unless requested. If the session provides a file-delivery or presentation mechanism, use it with the same final path.",
      "present -> report final path"
    ]
  ],
  "office-pptx": [
    [
      "Otherwise call `load_workspace_dependencies` and use its Python executable and bundled presentation libraries. Do not install npm or pip packages or locate a system interpreter for the default workflow. If the tool is unavailable, use an already configured environment and report a missing dependency only when it prevents the requested operation.",
      "Otherwise follow `../../python/SETUP.md` (relative to this skill's directory) to select a Python environment with `python-pptx` available; it defines the consent-first install for missing libraries. Do not install packages without the user's agreement. If no usable environment can be configured, report a missing dependency only when it prevents the requested operation.",
      "load_workspace_dependencies -> python/SETUP.md"
    ],
    [
      "Call `present({\"files\":[{\"path\":\"report-edited.pptx\"}]})` with the actual final PPTX path. It exposes the current source file without copying or preserving its bytes, so keep that file in place and leave intermediate images and reports out of delivery unless requested. If `present` is unavailable, use the session's supported file delivery method.",
      "Report the actual final PPTX path in the reply so the user can open it from the workspace; keep the file in place and leave intermediate images and reports out of delivery unless requested. If the session provides a file-delivery or presentation mechanism, use it with the same final path.",
      "present -> report final path"
    ]
  ],
  "office-render": [
    [
      "Use the installed LibreOffice Kit CLI through the shell with the absolute Node and CLI paths supplied below. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`.",
      "Use the installed LibreOffice Kit CLI through the shell. Resolve the absolute Node and CLI paths once per session by running, from this skill's directory, `node ../../scripts/resolve-dsoffice.mjs`; it prints `{ \"libreofficeKit\": { \"node\": \"<node>\", \"cli\": \"<cli>\" } }`. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`.",
      "injected kit paths -> resolve-dsoffice.mjs"
    ],
    [
      "If the Office sidecar is missing, follow [sidecar setup](../office/references/sidecar.md): ask whether the user wants to download it and install only after they agree.",
      "If path resolution fails or the engine is missing, follow [sidecar setup](../office/references/sidecar.md): fix or fetch the engine only after the user agrees.",
      "sidecar wording -> npm dependency wording"
    ],
    [
      "Deliver requested files through `present` when available; keep the files in place. Do not deliver inspection images or reports unless requested.",
      "Deliver requested files by reporting their final paths; keep the files in place. Do not deliver inspection images or reports unless requested.",
      "present -> report final path"
    ]
  ],
  "office-recalculate": [
    [
      "Use the installed LibreOffice Kit CLI through the shell with the absolute Node and CLI paths supplied below. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`.",
      "Use the installed LibreOffice Kit CLI through the shell. Resolve the absolute Node and CLI paths once per session by running, from this skill's directory, `node ../../scripts/resolve-dsoffice.mjs`; it prints `{ \"libreofficeKit\": { \"node\": \"<node>\", \"cli\": \"<cli>\" } }`. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`.",
      "injected kit paths -> resolve-dsoffice.mjs"
    ],
    [
      "If the Office sidecar is missing, follow [sidecar setup](../office/references/sidecar.md): ask whether the user wants to download it and install only after they agree.",
      "If path resolution fails or the engine is missing, follow [sidecar setup](../office/references/sidecar.md): fix or fetch the engine only after the user agrees.",
      "sidecar wording -> npm dependency wording"
    ],
    [
      "Verify the saved result with openpyxl, using `load_workspace_dependencies` for the Python path unless a project environment is required.",
      "Verify the saved result with openpyxl, using the Python environment from `../../python/SETUP.md` unless a project environment is required.",
      "load_workspace_dependencies -> python/SETUP.md"
    ],
    [
      "Deliver the verified output through `present` when available, keeping the file in place.",
      "Deliver the verified output by reporting its final path, keeping the file in place; use the session's file-delivery mechanism if it provides one.",
      "present -> report final path"
    ]
  ],
  "office-export-pdf": [
    [
      "Use the installed LibreOffice Kit CLI through the shell with the absolute Node and CLI paths supplied below. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`.",
      "Use the installed LibreOffice Kit CLI through the shell. Resolve the absolute Node and CLI paths once per session by running, from this skill's directory, `node ../../scripts/resolve-dsoffice.mjs`; it prints `{ \"libreofficeKit\": { \"node\": \"<node>\", \"cli\": \"<cli>\" } }`. In Bash, invoke `\"<node>\" \"<cli>\"`; in PowerShell, invoke `& \"<node>\" \"<cli>\"`.",
      "injected kit paths -> resolve-dsoffice.mjs"
    ],
    [
      "If the Office sidecar is missing, follow [sidecar setup](../office/references/sidecar.md): ask whether the user wants to download it and install only after they agree.",
      "If path resolution fails or the engine is missing, follow [sidecar setup](../office/references/sidecar.md): fix or fetch the engine only after the user agrees.",
      "sidecar wording -> npm dependency wording"
    ],
    [
      "Deliver the requested PDF through `present` when available, keeping the file in place.",
      "Deliver the requested PDF by reporting its final path, keeping the file in place; use the session's file-delivery mechanism if it provides one.",
      "present -> report final path"
    ]
  ]
};

const ATTRIBUTION_LINE = "# Ported from DSH: @deepseek-ai/dsh-skill-office (MIT) — see NOTICE.";

function parseArgs(argv) {
  const args = { mode: null, version: null, assets: null };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--check") args.mode = "check";
    else if (a === "--apply") args.mode = "apply";
    else if (a === "--version") args.version = argv[++i];
    else if (a === "--assets") args.assets = argv[++i];
    else if (a.charAt(0) !== "-") args.assets = a;
  }
  return args;
}

function splitVersion(v) {
  const dash = v.indexOf("-");
  const core = dash === -1 ? v : v.slice(0, dash);
  const pre = dash === -1 ? "" : v.slice(dash + 1);
  return { core: core.split(".").map(Number), pre: pre ? pre.split(".") : [] };
}

/** Semver ordering that also handles prerelease identifiers (alpha < beta < rc < release). */
function compareVersions(a, b) {
  const A = splitVersion(a);
  const B = splitVersion(b);
  for (let i = 0; i < 3; i++) {
    const x = A.core[i] || 0;
    const y = B.core[i] || 0;
    if (x !== y) return x - y;
  }
  if (A.pre.length === 0 && B.pre.length === 0) return 0;
  if (A.pre.length === 0) return 1;
  if (B.pre.length === 0) return -1;
  const len = Math.max(A.pre.length, B.pre.length);
  for (let i = 0; i < len; i++) {
    const ap = A.pre[i];
    const bp = B.pre[i];
    if (ap === undefined) return -1;
    if (bp === undefined) return 1;
    const an = /^[0-9]+$/.test(ap);
    const bn = /^[0-9]+$/.test(bp);
    if (an && bn) {
      const x = Number(ap);
      const y = Number(bp);
      if (x !== y) return x - y;
      continue;
    }
    if (an) return -1;
    if (bn) return 1;
    if (ap !== bp) return ap < bp ? -1 : 1;
  }
  return 0;
}

async function fetchRegistryMeta() {
  const res = await fetch(REGISTRY + encodeURIComponent(PACKAGE));
  if (!res.ok) throw new Error("registry request failed with HTTP " + res.status);
  return res.json();
}

/** Highest published version. dist-tags.latest is not used: prerelease trains often leave it behind. */
function newestVersion(meta) {
  const versions = Object.keys(meta.versions || {});
  if (versions.length === 0) throw new Error("registry returned no versions");
  return versions.sort(compareVersions)[versions.length - 1];
}

async function downloadUpstream(version, meta) {
  const entry = meta.versions[version];
  if (!entry) throw new Error("upstream version " + version + " is not published");
  const tarball = entry.dist && entry.dist.tarball;
  if (!tarball) throw new Error("registry entry has no tarball URL");
  const staging = join(ROOT, ".upstream-tmp");
  mkdirSync(staging, { recursive: true });
  const work = mkdtempSync(join(staging, "run-"));
  const tgz = join(work, "upstream.tgz");
  const res = await fetch(tarball);
  if (!res.ok) throw new Error("tarball download failed with HTTP " + res.status);
  writeFileSync(tgz, Buffer.from(await res.arrayBuffer()));
  // Relative paths + cwd: on Windows an absolute "F:\\..." archive path reads as a
  // remote host spec to the GNU tar shipped with Git, and System32's bsdtar rejects
  // the --force-local workaround. Staying relative works for both.
  execFileSync("tar", ["-xzf", "upstream.tgz", "-C", "."], { cwd: work, stdio: "ignore" });
  const assets = join(work, "package", "assets");
  if (!existsSync(assets)) throw new Error("upstream tarball has no assets directory");
  const texts = {};
  for (const name of SKILL_NAMES) {
    const p = join(assets, name, "SKILL.md");
    if (!existsSync(p)) throw new Error("upstream is missing assets/" + name + "/SKILL.md");
    texts[name] = readFileSync(p, "utf8");
  }
  return { texts: texts, cleanup: function () { rmSync(work, { recursive: true, force: true }); } };
}

/** Port straight from a local extraction of the upstream package (assets/ dir). */
function readLocalAssets(dir) {
  if (!existsSync(dir)) throw new Error("assets directory not found: " + dir + " (plain Node cannot read paths inside an .asar archive; extract the package tarball first)");
  const texts = {};
  for (const name of SKILL_NAMES) {
    const p = join(dir, name, "SKILL.md");
    if (!existsSync(p)) throw new Error("local assets are missing " + name + "/SKILL.md");
    texts[name] = readFileSync(p, "utf8");
  }
  return { texts: texts, cleanup: function () {} };
}

/** Apply every adaptation rule; a missing or duplicated anchor is a hard error. */
function portSkill(name, raw) {
  let text = raw;
  const rules = RULES[name] || [];
  for (const rule of rules) {
    const find = rule[0];
    const replace = rule[1];
    const label = rule[2];
    const first = text.indexOf(find);
    if (first < 0) throw new Error(name + ": rule no longer matches upstream text - " + label + ". Update RULES in scripts/sync-from-dsh.mjs.");
    if (text.indexOf(find, first + 1) >= 0) throw new Error(name + ": rule matches more than once - " + label + ". Update RULES in scripts/sync-from-dsh.mjs.");
    text = text.slice(0, first) + replace + text.slice(first + find.length);
  }
  if (!text.startsWith("---\n")) throw new Error(name + ": upstream SKILL.md does not start with a frontmatter block");
  text = text.replace(/^---\n/, "---\n" + ATTRIBUTION_LINE + "\n");
  if (!/^license:/m.test(text)) text = text.replace(/\n---\n/, "\nlicense: MIT\n---\n");
  return text;
}

function sha256(text) {
  return createHash("sha256").update(text).digest("hex");
}

/** Port every skill into skills/ and report which files changed on disk. */
function portAll(texts) {
  const report = [];
  for (const name of SKILL_NAMES) {
    const ported = portSkill(name, texts[name]);
    const targetPath = join(SKILLS_DIR, name, "SKILL.md");
    const previous = existsSync(targetPath) ? readFileSync(targetPath, "utf8") : null;
    writeFileSync(targetPath, ported);
    report.push(name + ": " + (previous === ported ? "unchanged" : "rewritten"));
  }
  return report;
}

/** Record the ported version and per-file upstream hashes in upstream.json. */
function writeBaseline(version, texts) {
  const current = JSON.parse(readFileSync(UPSTREAM_JSON, "utf8"));
  const next = { package: current.package, repo: current.repo, version: version, portedAt: new Date().toISOString().slice(0, 10), files: {} };
  for (const name of SKILL_NAMES) next.files[name] = sha256(texts[name]);
  writeFileSync(UPSTREAM_JSON, JSON.stringify(next, null, "\t") + "\n");
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const baseline = JSON.parse(readFileSync(UPSTREAM_JSON, "utf8"));

  // Offline form: port from a local extraction, no registry access.
  if (args.assets) {
    console.log("source   : local " + args.assets);
    console.log("baseline : " + baseline.version + " (ported " + baseline.portedAt + ")");
    const local = readLocalAssets(args.assets);
    try {
      const target = args.version || baseline.version;
      const report = portAll(local.texts);
      console.log(report.join("\n"));
      if (args.version) {
        writeBaseline(target, local.texts);
        console.log("\nbaseline updated to " + target + ".");
      } else {
        console.log("\nbaseline left at " + baseline.version + ". Pass --version <x> to record the ported version.");
      }
      console.log("note: sidecar.md, python/ and scripts/ are hand-maintained and are not rewritten here.");
    } finally {
      local.cleanup();
    }
    return;
  }

  if (!args.mode) {
    console.error("usage: node scripts/sync-from-dsh.mjs (--check | --apply | <assets-dir>) [--version X]");
    process.exit(2);
  }
  const meta = await fetchRegistryMeta();
  const target = args.version || newestVersion(meta);
  console.log("upstream : " + PACKAGE);
  console.log("baseline : " + baseline.version + " (ported " + baseline.portedAt + ")");
  console.log("target   : " + target);

  if (args.mode === "check") {
    if (target === baseline.version) {
      console.log("result   : already up to date");
      return;
    }
    const up = await downloadUpstream(target, meta);
    try {
      const changed = SKILL_NAMES.filter(function (n) { return sha256(up.texts[n]) !== baseline.files[n]; });
      console.log("result   : " + target + " is newer; changed skills: " + (changed.length > 0 ? changed.join(", ") : "none"));
      console.log("next     : run npm run apply:upstream, then review git diff");
    } finally {
      up.cleanup();
    }
    return;
  }

  const up = await downloadUpstream(target, meta);
  try {
    const report = portAll(up.texts);
    writeBaseline(target, up.texts);
    console.log(report.join("\n"));
    console.log("");
    console.log("ported " + target + ". Review with: git diff -- skills upstream.json");
    console.log("note: skills/office/references/sidecar.md, python/ and scripts/ are hand-maintained and are not rewritten here.");
  } finally {
    up.cleanup();
  }
}

main().catch(function (error) {
  console.error("sync-from-dsh failed: " + (error instanceof Error ? error.message : String(error)));
  process.exit(1);
});
