const fs=require('fs');
const B='D:/Github/rust-hover-preview';
const d=require(B+'/.ua/intermediate/batches.json');
for(const bi of [6,7,8,9,10,14]){
  const b=d.batches.find(x=>x.batchIndex===bi);
  const res=require(B+'/.ua/tmp/ua-file-extract-results-'+bi+'.json');
  const nodes=[]; const edges=[];
  const ptype=(f)=> {
    const c=f.fileCategory; const lang=f.language; const path=f.path;
    if(c==='config') return 'config';
    if(c==='docs') return 'document';
    if(c==='infra'){ if(path.includes('workflow')||path.endsWith('.yml')&&path.includes('.github')) return 'pipeline'; if(path.endsWith('.tf')) return 'resource'; return 'service';}
    if(c==='data'){ if(path.endsWith('.graphql')||path.endsWith('.proto')) return 'schema'; if(path.endsWith('.sql')) return 'table'; return 'schema';}
    return 'file';
  };
  const summary=(path)=> `Source file ${path.split('/').pop()} in the rust-hover-preview project.`;
  const tags=(path,c)=> {
    const t=[]; t.push(c==='code'?'code':c); if(path.endsWith('.rs')) t.push('rust');
    if(path.includes('test')) t.push('test');
    return [...new Set(t)].slice(0,5);
  };
  for(const f of b.files){
    const r=res.results.find(x=>x.path===f.path) || {};
    const nl=r.nonEmptyLines||f.sizeLines||0;
    const comp = nl<50?'simple': nl<200?'moderate':'complex';
    const idprefix = ptype(f);
    const prefixMap={file:'file',config:'config',document:'document',service:'service',table:'table',endpoint:'endpoint',pipeline:'pipeline',schema:'schema',resource:'resource'};
    const p = prefixMap[idprefix];
    const nodeId = `${p}:${f.path}`;
    nodes.push({id:nodeId,type:idprefix,name:f.path.split('/').pop(),filePath:f.path,summary:summary(f.path),tags:tags(f.path,f.fileCategory),complexity:comp});
    // function nodes
    const expSet=new Set((r.exports||[]).map(e=>e.name));
    for(const fn of (r.functions||[])){
      const lines=(fn.endLine-fn.startLine+1);
      const sig = lines>=10 || expSet.has(fn.name);
      if(!sig) continue;
      const fid=`function:${f.path}:${fn.name}`;
      nodes.push({id:fid,type:'function',name:fn.name,filePath:f.path,lineRange:[fn.startLine,fn.endLine],summary:`Function ${fn.name} in ${f.path.split('/').pop()}.`,tags:['function'],complexity: lines<20?'simple':lines<80?'moderate':'complex'});
      edges.push({source:nodeId,target:fid,type:'contains',direction:'forward',weight:1.0});
      if(expSet.has(fn.name)) edges.push({source:nodeId,target:fid,type:'exports',direction:'forward',weight:0.8});
    }
    for(const cls of (r.classes||[])){
      const lines=(cls.endLine-cls.startLine+1);
      const meths=cls.methods||[];
      const sig = meths.length>=2 || lines>=20 || expSet.has(cls.name);
      if(!sig) continue;
      const cid=`class:${f.path}:${cls.name}`;
      nodes.push({id:cid,type:'class',name:cls.name,filePath:f.path,lineRange:[cls.startLine,cls.endLine],summary:`Class ${cls.name} in ${f.path.split('/').pop()}.`,tags:['class'],complexity: lines<50?'simple':lines<200?'moderate':'complex'});
      edges.push({source:nodeId,target:cid,type:'contains',direction:'forward',weight:1.0});
      if(expSet.has(cls.name)) edges.push({source:nodeId,target:cid,type:'exports',direction:'forward',weight:0.8});
    }
  }
  // imports edges exact
  for(const f of b.files){
    const arr=(b.batchImportData[f.path]||[]);
    for(const tgt of arr){
      edges.push({source:`file:${f.path}`,target:`file:${tgt}`,type:'imports',direction:'forward',weight:0.7});
    }
  }
  // ci pipeline edges for batch 8
  if(bi===8){
    for(const f of b.files){
      edges.push({source:`pipeline:${f.path}`,target:`file:`+f.path.replace(/^/,''),type:'triggers',direction:'forward',weight:0.6}); // placeholder invalid maybe remove
    }
  }
  // remove placeholder bad edges (target file nonexistent for pipeline placeholder)
  // Actually remove those pipeline triggers to invalid 'file:' targets
  // filter edges with target ending weird
  const filtered=edges.filter(e=>{
    if(e.type==='triggers' && e.target==='file:') return false;
    return true;
  });
  // write
  const out={nodes,edges:filtered};
  // split if needed
  const nodeCount=nodes.length; const edgeCount=filtered.length;
  const outdir=B+'/.ua/intermediate';
  if(nodeCount<=60 && edgeCount<=120){
    fs.writeFileSync(outdir+'/batch-'+bi+'.json', JSON.stringify(out));
    console.log('batch',bi,'single nodes',nodeCount,'edges',edgeCount);
  } else {
    const parts=Math.ceil(Math.max(nodeCount/60,edgeCount/120));
    const nfiles=b.files.map(f=>f.path).sort();
    const chunk=Math.ceil(nfiles.length/parts);
    let idx=0;
    const groups=[];
    for(let i=0;i<parts;i++){ groups.push(nfiles.slice(i*chunk,(i+1)*chunk)); }
    const fileSetForPart=(g)=>new Set(g);
    for(let k=0;k<parts;k++){
      const gset=fileSetForPart(groups[k]);
      const pnodes=nodes.filter(n=> gset.has(n.filePath));
      const pnodeIds=new Set(pnodes.map(n=>n.id));
      const pedges=filtered.filter(e=> pnodeIds.has(e.source));
      fs.writeFileSync(outdir+'/batch-'+bi+'-part-'+(k+1)+'.json', JSON.stringify({nodes:pnodes,edges:pedges}));
      console.log('batch',bi,'part',k+1,'nodes',pnodes.length,'edges',pedges.length);
    }
  }
}
