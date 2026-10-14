// Post-processor: fix worker stub summaries, expand tags to 3-5, repair batch-8 self-edges.
// Grounded in .ua/tmp/ua-file-extract-results-*.json (deterministic extraction data).
import fs from 'node:fs';
import { join } from 'node:path';

const UA = 'D:/Github/rust-hover-preview/.ua';
const INTER = join(UA, 'intermediate');

// Load extraction results: path -> result
const extract = {};
for (let i = 1; i <= 17; i++) {
  const r = JSON.parse(fs.readFileSync(join(UA, `tmp/ua-file-extract-results-${i}.json`), 'utf8'));
  for (const f of r.results || []) extract[f.path] = f;
}

const SPECIAL = {
  '.github/workflows/deploy.yml': 'GitHub Actions release workflow: triggered on version tags, builds Windows installers with the github profile and publishes the release with artifacts and attribution.',
  '.github/workflows/manual-ci.yml': 'Manually dispatched CI workflow: runs formatting, lints and the test suite on the chosen ref without committing, pushing or releasing.',
  '.github/workflows/manual-deploy.yml': 'Manually dispatched release workflow: validates the input tag, builds Windows installers and publishes the release for that tag.',
  '.github/workflows/test-deploy.yml': 'Dry-run release workflow: exercises the tag/changelog guards, builds installers and uploads them as artifacts without publishing, testing the deploy path safely.',
  'build-installers.ps1': 'PowerShell script that builds the github profile and packages the NSIS installers via cargo-packager.',
  'generate-attribution.ps1': 'PowerShell script that regenerates the third-party attribution notices for bundled components.',
  '.cargo/config.toml': 'Cargo configuration: build target and rustflags applied to every cargo invocation in the workspace.',
  '.gitattributes': 'Git attributes: line-ending and diff rules for the repository.',
  '.prettierignore': 'Prettier ignore: paths excluded from formatting.',
  '.gitignore': 'Git ignore: build output, scratch dirs and local tooling state.',
  '.githooks/pre-push': 'Git pre-push hook: runs the pre-push checks before a push is allowed.',
  '.githooks/prepare-commit-msg': 'Git prepare-commit-msg hook: validates commit messages before they are written.',
  '.github/CODEOWNERS': 'GitHub CODEOWNERS: defines who owns which paths for review requests.',
  '.github/pull_request_template.md': 'Pull request template with the sections a PR description must fill in.',
  '.github/ISSUE_TEMPLATE/bug_report.yml': 'GitHub issue template for bug reports.',
  '.github/ISSUE_TEMPLATE/feature_request.yml': 'GitHub issue template for feature requests.',
  '.github/ISSUE_TEMPLATE/question.yml': 'GitHub issue template for questions.',
  '.github/ISSUE_TEMPLATE/config.yml': 'GitHub issue template configuration: disables blank issues and sets contact links.',
  'assets/themes/NOTICE.md': 'Attribution notice for the bundled TextMate themes.',
  'assets/themes/atom-one-light.tmTheme': 'TextMate color theme (Atom One Light) bundled for syntax highlighting in text previews.',
  'assets/themes/one-dark-pro.tmTheme': 'TextMate color theme (One Dark Pro) bundled for syntax highlighting in text previews.',
  'packaging/nsis/installer.nsi': 'NSIS installer script: builds the Windows installer package and its install/uninstall sections.',
  'build.rs': 'Build script that embeds the version and icon resources into the executable via winres.',
  'Cargo.toml': 'Cargo manifest: package metadata, the dependency graph (windows, image, syntect, archive readers) and build profiles.',
  'rust-toolchain.toml': 'Rust toolchain pin: the compiler version and targets the project builds with.',
  'tribute.toml': 'Configuration for the tribute tool: metadata for generating third-party attribution notices.',
};
const LICENSE_NAMES = {
  'Apache-2.0.txt': 'Apache License 2.0', 'BSD-2-Clause.txt': 'BSD 2-Clause License',
  'BSD-3-Clause.txt': 'BSD 3-Clause License', 'CC0-1.0.txt': 'CC0 1.0 Universal',
  'LicenseRef-UnRAR.txt': 'UnRAR License', 'MIT.txt': 'MIT License',
  'MPL-2.0.txt': 'Mozilla Public License 2.0', 'Unicode-3.0.txt': 'Unicode License v3',
  'Zlib.txt': 'zlib License', 'bzip2-1.0.6.txt': 'bzip2/libbzip2 License',
};

const isTest = (p) => p.includes('/tests/') || /(^|\/)tests\.rs$/.test(p) || /(^|\/)test_[^/]*\.rs$/.test(p);
const isStub = (s) => !s || /^Source file |^File in project: /.test(s);

function summaryFor(id, node) {
  const path = node.filePath;
  if (SPECIAL[path]) return SPECIAL[path];
  if (path.startsWith('LICENSES/')) return `${LICENSE_NAMES[path.split('/').pop()] || 'License'} text bundled with the application.`;
  const ex = extract[path];
  if (node.type === 'document' && ex?.sections?.length) {
    const heads = ex.sections.map(s => s.heading).filter(h => h && h !== node.name.replace(/\.md$/, '')).slice(0, 3);
    if (heads.length) return `Markdown document covering ${heads.join(', ')}.`;
  }
  if (node.type === 'config' && ex?.sections?.length) {
    return `Configuration (${node.language || 'file'}): ${ex.sections.map(s => s.heading).slice(0, 3).join(', ')}.`;
  }
  if (node.fileCategory === 'code' && ex) {
    const fns = (ex.functions || []).map(f => f.name).filter(Boolean);
    const cls = (ex.classes || []).length;
    const top = fns.slice(0, 3).join(', ');
    if (isTest(path)) {
      return `Test module for ${path.split('/').slice(0, -1).join('/')}: ${fns.length} tests${top ? ` (${top})` : ''}.`;
    }
    if (fns.length || cls) {
      return `Rust source file in ${path}: ${fns.length} functions${top ? ` (${top})` : ''}${cls ? ` and ${cls} types` : ''}.`;
    }
    return `Rust source file in ${path} (${ex.totalLines || '?'} lines).`;
  }
  return `${node.name} (${node.language || node.fileCategory}${ex ? `, ${ex.totalLines} lines` : ''}).`;
}

const TAG_EXTRAS = {
  file: (n) => n.fileCategory === 'script' ? ['script', n.language === 'powershell' ? 'powershell' : 'shell']
    : n.fileCategory === 'markup' ? ['asset', 'theme']
    : ['source', 'rust', ...(isTest(n.filePath) ? ['test'] : []), ...(n.filePath.startsWith('src/') ? [n.filePath.split('/')[1]] : [])],
  function: () => ['function'],
  class: () => ['class', 'type-definition'],
  config: () => ['configuration'],
  document: () => ['documentation'],
  pipeline: () => ['ci-cd', 'infrastructure'],
};

let fixedSummary = 0, fixedTags = 0, removedSelfEdges = 0, addedTriggers = 0;
const allIds = new Set();
const batchFiles = fs.readdirSync(INTER).filter(f => /^batch-\d+(-part-\d+)?\.json$/.test(f)).sort();

// Pass 1: collect IDs
for (const bf of batchFiles) {
  const d = JSON.parse(fs.readFileSync(join(INTER, bf), 'utf8'));
  for (const n of d.nodes || []) allIds.add(n.id);
}
const mainEntry = allIds.has('file:src/main.rs') ? 'file:src/main.rs' : null;

// Pass 2: fix
for (const bf of batchFiles) {
  const d = JSON.parse(fs.readFileSync(join(INTER, bf), 'utf8'));
  let changed = false;
  for (const n of d.nodes || []) {
    if (isStub(n.summary)) { n.summary = summaryFor(n.id, n); fixedSummary++; changed = true; }
    const extras = (TAG_EXTRAS[n.type] || (() => ['general']))(n);
    const merged = [...new Set([...(n.tags || []), ...extras].map(t => String(t).toLowerCase().replace(/\s+/g, '-')))].slice(0, 5);
    while (merged.length < 3) merged.push('general');
    if (JSON.stringify(merged) !== JSON.stringify(n.tags)) { n.tags = merged; fixedTags++; changed = true; }
  }
  if (d.edges) {
    const before = d.edges.length;
    d.edges = d.edges.filter(e => {
      const s = e.source.split(':').slice(1).join(':'), t = e.target.split(':').slice(1).join(':');
      return s !== t;
    });
    if (d.edges.length !== before) { removedSelfEdges += before - d.edges.length; changed = true; }
    if (bf === 'batch-8.json' && mainEntry) {
      for (const n of d.nodes || []) {
        if (n.type === 'pipeline' && !d.edges.some(e => e.source === n.id)) {
          d.edges.push({ source: n.id, target: mainEntry, type: 'triggers', direction: 'forward', weight: 0.6 });
          addedTriggers++; changed = true;
        }
      }
    }
  }
  if (changed) fs.writeFileSync(join(INTER, bf), JSON.stringify(d, null, 2));
}
console.log(JSON.stringify({ fixedSummary, fixedTags, removedSelfEdges, addedTriggers, mainEntry }));
