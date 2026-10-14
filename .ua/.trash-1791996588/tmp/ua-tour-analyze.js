const fs = require('fs');
const inputPath = process.argv[2];
const outputPath = process.argv[3];
if (!inputPath || !outputPath) { console.error('usage: node script input output'); process.exit(1); }
let data;
try { data = JSON.parse(fs.readFileSync(inputPath, 'utf8')); } catch(e){ console.error(e.message); process.exit(1);}
const nodes = data.nodes || [];
const edges = data.edges || [];
const layers = data.layers || [];
const idToNode = {};
nodes.forEach(n => { idToNode[n.id] = n; });
const isCode = (n) => n.type !== 'document' && n.type !== 'service' && n.type !== 'pipeline' && n.type !== 'resource' && n.type !== 'config' && n.type !== 'table' && n.type !== 'schema' && n.type !== 'endpoint';
// note: generic; only code entry detection below

const nodeIds = nodes.map(n=>n.id);
const fanIn = {}; const fanOut = {};
nodeIds.forEach(id=>{fanIn[id]=0;fanOut[id]=0;});
edges.forEach(e=>{
  if (fanOut[e.source]!==undefined) fanOut[e.source]++;
  if (fanIn[e.target]!==undefined) fanIn[e.target]++;
});

const fanInRanking = nodeIds.map(id=>({id, fanIn: fanIn[id], name: (idToNode[id].name)})).sort((a,b)=>b.fanIn-a.fanIn).slice(0,20).map(x=>({...x, summary: idToNode[x.id].summary}));
const fanOutRanking = nodeIds.map(id=>({id, fanOut: fanOut[id], name: (idToNode[id].name)})).sort((a,b)=>b.fanOut-a.fanOut).slice(0,20).map(x=>({...x, summary: idToNode[x.id].summary}));

const ENTRY_FILENAMES = new Set(['index.ts','index.js','main.ts','main.js','app.ts','app.js','server.ts','server.js','mod.rs','main.go','main.py','main.rs','main.cpp','main.c','manage.py','app.py','wsgi.py','asgi.py','run.py','__main__.py','Application.java','Main.java','Program.cs','config.ru','index.php','App.swift','Application.kt']);
const depthOfPath = (p) => !p?99 : p.split('/').length-1; // root file =0? file at root has 0 slashes? src/index.ts ->1
// Actually root or one level deep => depth <=1 segment? We'll score: path depth 0 or 1.
function scoreEntry(n){
  let s=0;
  const fp = n.filePath||'';
  const basename = fp.split('/').pop();
  if (n.type==='document') {
    if (fp==='README.md') return 5;
    if (fp.endsWith('.md') && !fp.includes('/')) return 2;
    return 0;
  }
  // code file heuristic: not non-code types
  const nonCodeTypes = new Set(['document','service','pipeline','resource','config','table','schema','endpoint']);
  if (nonCodeTypes.has(n.type)) return 0;
  if (ENTRY_FILENAMES.has(basename)) s+=3;
  const segs = fp.split('/').filter(Boolean);
  if (segs.length<=2) s+=1; // root or one level
  // high fanOut top10%
  return s; // add ranking conditions after
}
// compute thresholds
const fanOutVals = nodeIds.map(id=>fanOut[id]).sort((a,b)=>b-a);
const top10cut = fanOutVals[Math.floor(nodeIds.length*0.1)]||0;
const bottom25fanin = fanInValsSorted = nodeIds.map(id=>fanIn[id]).sort((a,b)=>a-b); const cut25 = fanInValsSorted[Math.floor(nodeIds.length*0.25)]||0;

const scored = nodes.map(n=>{
  let s = scoreEntry(n);
  if (s===0 && n.type!=='document') {
    const nonCodeTypes = new Set(['document','service','pipeline','resource','config','table','schema','endpoint']);
    if (nonCodeTypes.has(n.type)) s=0;
    else {
      // code file but not entryname etc; base 0, plus fanout/fanin
    }
  }
  // high fanout
  const nonCodeTypes = new Set(['document','service','pipeline','resource','config','table','schema','endpoint']);
  if (!nonCodeTypes.has(n.type)) {
    if (fanOut[n.id] >= top10cut && fanOut[n.id]>0) s+=1;
    if (fanIn[n.id] <= cut25) s+=1;
  }
  return {id:n.id, score:s, name:n.name, summary:n.summary};
});
const entryPointCandidates = scored.filter(x=>x.score>0).sort((a,b)=>b.score-a.score).slice(0,5);

// BFS from top code entry point (skip docs)
let startNode = null;
for (const c of scored.filter(x=>x.score>0).sort((a,b)=>b.score-a.score)) {
  const n = idToNode[c.id];
  const nonCodeTypes = new Set(['document','service','pipeline','resource','config','table','schema','endpoint']);
  if (!nonCodeTypes.has(n.type)) { startNode = c.id; break; }
}
// fallback to main.rs if exists
if (!startNode) {
  const mainish = nodeIds.find(id=>/main\.rs$/.test(idToNode[id].filePath||''));
  if (mainish) startNode = mainish;
}

const adj = {};
nodeIds.forEach(id=>{adj[id]=[];});
edges.forEach(e=>{
  if (e.type==='imports' || e.type==='calls') {
    if (adj[e.source]) adj[e.source].push(e.target);
  }
});

const visited = new Set(); const order=[]; const depthMap={}; const byDepth={};
if (startNode) {
  const queue=[[startNode,0]];
  visited.add(startNode);
  while(queue.length){
    const [id,d]=queue.shift();
    order.push(id); depthMap[id]=d; if(!byDepth[d])byDepth[d]=[]; byDepth[d].push(id);
    for(const t of (adj[id]||[])){
      if(!visited.has(t) && idToNode[t]){visited.add(t);queue.push([t,d+1]);}
    }
  }
}

// clusters: bidirectional pairs among imports/calls
const edgeSet = new Set(edges.map(e=>e.source+'->'+e.target+'|'+e.type));
function hasEdge(a,b,type){ return edgeSet.has(a+'->'+b+'|'+type); }
// build undirected adjacency for all edges? For clusters, use imports/calls both directions.
const pairEdges = new Map();
edges.forEach(e=>{
  if(e.type==='imports'||e.type==='calls'){
    const key = [e.source,e.target].sort().join('::');
    pairEdges.set(key,(pairEdges.get(key)||0)+1);
  }
});
// find bidirectional
const bidirPairs=[];
const seenPair=new Set();
edges.forEach(e=>{
  if((e.type==='imports'||e.type==='calls') && (hasEdge(e.target,e.source,'imports')||hasEdge(e.target,e.source,'calls'))){
    const key=[e.source,e.target].sort().join('::');
    if(!seenPair.has(key)){seenPair.add(key); bidirPairs.push([e.source,e.target]);}
  }
});
// group into clusters
const clusterMap = {}; // node -> set
bidirPairs.forEach(([a,b])=>{
  // find existing cluster containing a or b
  let target=null;
  for(const k in clusterMap){
    if(clusterMap[k].has(a)||clusterMap[k].has(b)){target=k;break;}
  }
  if(!target){target=a; clusterMap[target]=new Set([a,b]);}
  else { clusterMap[target].add(a); clusterMap[target].add(b); }
});
// expand: add nodes connecting to 2+ members
let changed=true;
while(changed){
  changed=false;
  nodes.forEach(n=>{
    Object.keys(clusterMap).forEach(k=>{
      const cs=clusterMap[k];
      if(cs.has(n.id)) return;
      let cnt=0;
      cs.forEach(m=>{
        // connection if any edge between n.id and m
        if(edges.some(e=>(e.source===n.id&&e.target===m)||(e.source===m&&e.target===n))) cnt++;
      });
      if(cnt>=2){ cs.add(n.id); changed=true; }
    });
  });
}
// final clusters with size 2-5
const clusters=[];
const used=new Set();
Object.values(clusterMap).forEach(cs=>{
  const arr=[...cs];
  if(arr.length<2||arr.length>5) return;
  if(arr.some(x=>used.has(x))) return;
  const edgeCount = edges.filter(e=>arr.includes(e.source)&&arr.includes(e.target)).length;
  arr.forEach(x=>used.add(x));
  clusters.push({nodes:arr, edgeCount});
});
clusters.sort((a,b)=>b.edgeCount-a.edgeCount);
const topClusters = clusters.slice(0,10);

// non-code files
const nonCode = {documentation:[], infrastructure:[], data:[], config:[]};
nodes.forEach(n=>{
  if (n.type==='document') nonCode.documentation.push({id:n.id,name:n.name,summary:n.summary});
  else if (n.type==='service'||n.type==='pipeline'||n.type==='resource') nonCode.infrastructure.push({id:n.id,name:n.name,type:n.type,summary:n.summary});
  else if (n.type==='table'||n.type==='schema'||n.type==='endpoint') nonCode.data.push({id:n.id,name:n.name,type:n.type,summary:n.summary});
  else if (n.type==='config') nonCode.config.push({id:n.id,name:n.name,summary:n.summary});
});

const nodeSummaryIndex = {};
nodes.forEach(n=>{ nodeSummaryIndex[n.id]={name:n.name,type:n.type,summary:n.summary}; });

const bfsTraversal = {
  startNode,
  order,
  depthMap,
  byDepth: Object.fromEntries(Object.entries(byDepth).map(([k,v])=>[k,v]))
};

const result = {
  scriptCompleted:true,
  entryPointCandidates,
  fanInRanking,
  fanOutRanking,
  bfsTraversal,
  nonCodeFiles: nonCode,
  clusters: topClusters,
  layers: {count:layers.length, list: layers.map(l=>({id:l.id,name:l.name,description:l.description}))},
  nodeSummaryIndex,
  totalNodes: nodes.length,
  totalEdges: edges.length
};
fs.writeFileSync(outputPath, JSON.stringify(result,null,2));
console.log('done', nodes.length, edges.length);
process.exit(0);
