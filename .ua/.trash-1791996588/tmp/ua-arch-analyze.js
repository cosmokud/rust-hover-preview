const fs = require('fs');
function main() {
  const inputPath = process.argv[2];
  const outputPath = process.argv[3];
  if (!inputPath || !outputPath) { console.error('Usage: node script input.json output.json'); process.exit(1); }
  const data = JSON.parse(fs.readFileSync(inputPath, 'utf8'));
  const fileNodes = data.fileNodes || [];
  const importEdges = data.importEdges || [];
  const allEdges = data.allEdges || [];

  // helper: common prefix
  const paths = fileNodes.map(n => n.filePath);
  // common path prefix (directory prefix including trailing slash)
  let commonPrefix = '';
  if (paths.length) {
    const parts = paths.map(p => p.split('/'));
    const minLen = Math.min(...parts.map(p=>p.length));
    const shared = [];
    for (let i = 0; i < minLen; i++) {
      const seg = parts[0][i];
      if (parts.every(p => p[i] === seg)) shared.push(seg); else break;
    }
    // common prefix is shared dirs excluding file itself: if shared length == file path segments count? We want dir prefix.
    // Typical: if all start with src/, shared includes 'src' and then maybe deeper. We stop at first differing dir? 
    // Definition: "common path prefix shared by all files (e.g., if all paths start with src/, the common prefix is src/)"
    // We'll treat common prefix as all shared segments except last file segment? Simpler: shared segments where all files share; but if all same file? not here.
    // Use shared segments but if all files share full path identical? no. Keep shared[:-1] as dir prefix? 
    // Actually for src/foo.ts and src/bar.ts, shared=['src','foo.ts']? no second differs. shared=['src']. prefix='src/'.
    // For src/a/x.ts and src/b/y.ts, shared=['src']; prefix='src/'. good.
    // For a/b/c.ts and a/b/d.ts, shared=['a','b']; prefix='a/b/'? Then grouping first segment after prefix none -> root? 
    // Spec says group by first directory segment after prefix. We'll set prefix to shared dirs joined + slash, but ensure leaves at least? Use all shared segments except require at least one grouping possible? 
    // If shared includes file name for all? impossible unless same file. So shared last is dir? For src/x.ts and src/y.ts shared=['src']. ok.
    commonPrefix = shared.length ? shared.join('/') + '/' : '';
  }

  // group key for a file path
  function groupKey(filePath) {
    let rel = filePath;
    if (commonPrefix && rel.startsWith(commonPrefix)) rel = rel.slice(commonPrefix.length);
    const idx = rel.indexOf('/');
    if (idx === -1) {
      // flat? handle by extension pattern maybe; but just 'root'
      return 'root';
    }
    return rel.split('/')[0];
  }

  // A directory groups
  const directoryGroups = {};
  for (const n of fileNodes) {
    const g = groupKey(n.filePath);
    (directoryGroups[g] = directoryGroups[g] || []).push(n.id);
  }

  // B node type groups
  const nodeTypeGroups = {};
  for (const n of fileNodes) {
    const t = n.type || 'file';
    (nodeTypeGroups[t] = nodeTypeGroups[t] || []).push(n.id);
  }

  // C adjacency + fanin/fanout
  const fanOut = {}; const fanIn = {};
  for (const n of fileNodes) { fanOut[n.id]=0; fanIn[n.id]=0; }
  const idToNode = Object.fromEntries(fileNodes.map(n=>[n.id,n]));
  const groupOf = Object.fromEntries(fileNodes.map(n=>[n.id, groupKey(n.filePath)]));
  // adjacency list
  const adj = {}; for (const n of fileNodes) adj[n.id]=[];
  for (const e of importEdges) {
    if (adj[e.source]) adj[e.source].push(e.target);
  }
  for (const id of Object.keys(adj)) {
    fanOut[id] = adj[id].length;
    for (const t of adj[id]) if (fanIn[t] !== undefined) fanIn[t]++;
  }
  // group-level imports
  const groupImports = {}; // from -> set to
  for (const e of importEdges) {
    const gs = groupOf[e.source], gt = groupOf[e.target];
    if (gs === undefined || gt === undefined) continue;
    if (!groupImports[gs]) groupImports[gs] = {};
    if (gs !== gt) groupImports[gs][gt] = (groupImports[gs][gt]||0)+1;
  }

  // D cross-category edges from allEdges
  const crossMap = {}; // fromType->toType->edgeType->count
  const crossListArr = [];
  const crossAgg = {};
  for (const e of allEdges) {
    const s = idToNode[e.source]; const t = idToNode[e.target];
    if (!s || !t) continue;
    const st = s.type || 'file'; const tt = t.type || 'file';
    if (st === tt && st === 'file') {
      // still could record but spec wants cross-category? We'll record all type combos maybe filter? Spec example config->file. We'll aggregate all combos including file->file? Could be huge. Focus: group by (fromType,toType,edgeType). Include all.
    }
    const key = st + '->' + tt + ':' + e.type;
    crossAgg[key] = (crossAgg[key]||0)+1;
  }
  const crossCategoryEdges = Object.entries(crossAgg).map(([k,v]) => {
    const [left, et] = k.split(':');
    const [ft, tt] = left.split('->');
    return { fromType: ft, toType: tt, edgeType: et, count: v };
  }).sort((a,b)=>b.count-a.count);

  // E inter-group import frequency + F intra density
  const interCounts = {};
  const intraInternal = {};
  const groupTotalEdges = {};
  for (const g of Object.keys(directoryGroups)) { intraInternal[g]=0; groupTotalEdges[g]=0; }
  for (const e of importEdges) {
    const gs = groupOf[e.source], gt = groupOf[e.target];
    if (gs === undefined || gt === undefined) continue;
    groupTotalEdges[gs] = (groupTotalEdges[gs]||0)+1;
    if (gs !== gt) {
      const key = gs + '->' + gt;
      interCounts[key] = (interCounts[key]||0)+1;
    } else {
      intraInternal[gs]++;
    }
  }
  const interGroupImports = Object.entries(interCounts).map(([k,v])=>{
    const [from,to]=k.split('->');
    return {from,to,count:v};
  }).sort((a,b)=>b.count-a.count);
  const intraGroupDensity = {};
  for (const g of Object.keys(directoryGroups)) {
    const total = groupTotalEdges[g] || 0;
    const internal = intraInternal[g] || 0;
    intraGroupDensity[g] = { internalEdges: internal, totalEdges: total, density: total ? +(internal/total).toFixed(3) : 0 };
  }

  // G pattern matching
  const dirPatternMap = {
    routes:'api', api:'api', controllers:'api', endpoints:'api', handlers:'api', routers:'api', controller:'api', serializers:'api', blueprints:'api',
    services:'service', core:'service', lib:'service', domain:'service', logic:'service', internal:'service', signals:'service', composables:'service', mailers:'service', jobs:'service', channels:'service',
    models:'data', db:'data', data:'data', persistence:'data', repository:'data', entities:'data', entity:'data', migrations:'data', sql:'data', database:'data', schema:'data',
    components:'ui', views:'ui', pages:'ui', ui:'ui', layouts:'ui', screens:'ui',
    middleware:'middleware', plugins:'middleware', interceptors:'middleware', guards:'middleware',
    utils:'utility', helpers:'utility', common:'utility', shared:'utility', tools:'utility', templatetags:'utility', pkg:'utility',
    config:'config', constants:'config', env:'config', settings:'config', management:'config', commands:'config',
    tests:'test', test:'test', spec:'test', specs:'test', __tests__:'test',
    types:'types', interfaces:'types', schemas:'types', contracts:'types', dtos:'types', dto:'types', request:'types', response:'types',
    hooks:'hooks', store:'state', state:'state', reducers:'state', actions:'state', slices:'state',
    assets:'assets', static:'assets', public:'assets',
    docs:'documentation', documentation:'documentation', wiki:'documentation',
    deploy:'infrastructure', deployment:'infrastructure', infra:'infrastructure', infrastructure:'infrastructure', k8s:'infrastructure', kubernetes:'infrastructure', helm:'infrastructure', charts:'infrastructure', terraform:'infrastructure', tf:'infrastructure', docker:'infrastructure',
    bin:'entry', cmd:'entry',
  };
  const patternMatches = {};
  for (const g of Object.keys(directoryGroups)) {
    const gl = g.toLowerCase();
    if (dirPatternMap[gl]) patternMatches[g] = dirPatternMap[gl];
  }
  // file-level patterns applied to each file? patternMatches is per directory group per spec. But also group names no match for root/crates. We'll add fallbacks below? Keep only dir.

  // H deployment topology
  const infraFiles = [];
  let hasDockerfile=false,hasCompose=false,hasK8s=false,hasTerraform=false,hasCI=false;
  for (const n of fileNodes) {
    const p = n.filePath.toLowerCase(); const name = n.name.toLowerCase();
    if (name === 'dockerfile' || name.startsWith('dockerfile.')) { hasDockerfile=true; infraFiles.push(n.filePath); }
    if (name.startsWith('docker-compose') || name === 'compose.yaml' || name === 'compose.yml') { hasCompose=true; if(!infraFiles.includes(n.filePath)) infraFiles.push(n.filePath); }
    if (p.includes('k8s/')||p.includes('kubernetes/')||p.endsWith('.yaml')&&p.includes('chart')) {} 
    if (p.includes('/k8s/')||p.includes('/kubernetes/')||p.includes('/helm/')||p.includes('/charts/')) { hasK8s=true; }
    if (p.endsWith('.tf')||p.endsWith('.tfvars')) { hasTerraform=true; }
    if (p.includes('.github/workflows/')||name==='.gitlab-ci.yml'||name==='jenkinsfile') { hasCI=true; }
  }
  // k8s detection more simply: path contains k8s or helm
  for (const n of fileNodes) {
    const p=n.filePath.toLowerCase();
    if (p.includes('k8s')||p.includes('helm')||p.includes('charts')) hasK8s=true;
    if (p.endsWith('.tf')||p.endsWith('.tfvars')) { hasTerraform=true; if(!infraFiles.includes(n.filePath)) infraFiles.push(n.filePath);}
  }
  if (hasCI) {
    for (const n of fileNodes) {
      if (n.filePath.includes('.github/workflows/')) if(!infraFiles.includes(n.filePath)) infraFiles.push(n.filePath);
    }
  }
  const deploymentTopology = { hasDockerfile, hasCompose, hasK8s, hasTerraform, hasCI, infraFiles };

  // I data pipeline
  const schemaFiles=[], migrationFiles=[], dataModelFiles=[], apiHandlerFiles=[];
  for (const n of fileNodes) {
    const p=n.filePath.toLowerCase();
    if (p.endsWith('.sql') && p.includes('migration')) migrationFiles.push(n.filePath);
    else if (p.endsWith('.sql')) schemaFiles.push(n.filePath);
    if (p.endsWith('.graphql')||p.endsWith('.gql')||p.endsWith('.proto')||p.endsWith('.prisma')) schemaFiles.push(n.filePath);
    if (/model/i.test(p)) dataModelFiles.push(n.filePath);
    if (/route|controller|handler/i.test(p)) apiHandlerFiles.push(n.filePath);
  }
  const dataPipeline = { schemaFiles, migrationFiles, dataModelFiles, apiHandlerFiles };

  // J doc coverage
  const groupsWithDocs = Object.keys(directoryGroups).filter(g => directoryGroups[g].some(id => {
    const n=idToNode[id]; return n && n.type==='document';
  })).length;
  // also docs dir? count groups that contain a README
  const totalGroups = Object.keys(directoryGroups).length;
  const undocumentedGroups = Object.keys(directoryGroups).filter(g => !directoryGroups[g].some(id=>{const n=idToNode[id]; return n && (n.type==='document' || /readme/i.test(n.name));}));
  const docCoverage = { groupsWithDocs, totalGroups, coverageRatio: totalGroups? +(groupsWithDocs/totalGroups).toFixed(2):0, undocumentedGroups };

  // K dependency direction
  const dependencyDirection = [];
  const groupPairs = {};
  for (const e of importEdges) {
    const gs=groupOf[e.source], gt=groupOf[e.target];
    if (gs===undefined||gt===undefined||gs===gt) continue;
    const k=gs+'|'+gt;
    groupPairs[k]=(groupPairs[k]||0)+1;
  }
  const pairSet = new Set(Object.keys(groupPairs).map(k=>k.split('|').sort().join('|')));
  for (const pk of pairSet) {
    const [a,b]=pk.split('|');
    const ab=groupPairs[a+'|'+b]||0;
    const ba=groupPairs[b+'|'+a]||0;
    if (ab>ba && ab>0) dependencyDirection.push({dependent:a, dependsOn:b});
    else if (ba>ab && ba>0) dependencyDirection.push({dependent:b, dependsOn:a});
  }

  const result = {
    scriptCompleted: true,
    directoryGroups,
    nodeTypeGroups,
    crossCategoryEdges,
    interGroupImports,
    intraGroupDensity,
    patternMatches,
    deploymentTopology,
    dataPipeline,
    docCoverage,
    dependencyDirection,
    fileStats: {
      totalFileNodes: fileNodes.length,
      filesPerGroup: Object.fromEntries(Object.entries(directoryGroups).map(([k,v])=>[k,v.length])),
      nodeTypeCounts: Object.fromEntries(Object.entries(nodeTypeGroups).map(([k,v])=>[k,v.length]))
    },
    fileFanIn: fanIn,
    fileFanOut: fanOut,
    commonPrefix
  };
  fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
  process.exit(0);
}
main();