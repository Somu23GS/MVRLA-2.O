const fs=require('fs');
const src=fs.readFileSync('/home/user/lens_v56_section.js','utf8');
const a=src.indexOf('var _lensDCTCos');
const b=src.indexOf('function loadImgEl',a);
eval(src.slice(a,b));

function makeImage(kind, opt={}){
  const W=96,H=96,px=new Uint8ClampedArray(W*H*4);
  const bg=opt.bg||[215,218,222], fg=opt.fg||[110,120,130];
  const cx=opt.cx==null?48:opt.cx, cy=opt.cy==null?48:opt.cy, r=opt.r||28;
  for(let y=0;y<H;y++)for(let x=0;x<W;x++){
    const i=(y*W+x)*4;px[i]=bg[0];px[i+1]=bg[1];px[i+2]=bg[2];px[i+3]=255;
  }
  function set(x,y,c){if(x<0||y<0||x>=W||y>=H)return;const i=((y|0)*W+(x|0))*4;px[i]=c[0];px[i+1]=c[1];px[i+2]=c[2]}
  for(let y=0;y<H;y++)for(let x=0;x<W;x++){
    const dx=x-cx,dy=y-cy,d=Math.sqrt(dx*dx+dy*dy),ang=Math.atan2(dy,dx);
    let inside=false;
    if(kind==='ring') inside=d<r&&d>r*.43;
    else if(kind==='gear'){
      const edge=r*(.84+.16*(Math.cos(12*ang+(opt.phase||0))>0?1:0));
      inside=d<edge&&d>r*.25;
    }else if(kind==='sensor') inside=Math.abs(dx)<r*.62&&Math.abs(dy)<r*.92;
    else if(kind==='hex'){
      inside=d<r*(.86+.14*Math.cos(6*ang))&&d>r*.28;
    }
    if(inside){
      const shine=Math.max(-25,Math.min(25,(dx-dy)*.45));
      set(x,y,[fg[0]+shine,fg[1]+shine,fg[2]+shine]);
    }
  }
  return {px,W,H};
}
function desc(im){return lensDescFromRGBA(im.px,im.W,im.H)}
function score(a,b){return lensSim(a,b)}
const ring=desc(makeImage('ring'));
const gear=desc(makeImage('gear'));
const sensor=desc(makeImage('sensor'));
const hex=desc(makeImage('hex'));
const refs=[['ring',ring],['gear',gear],['sensor',sensor],['hex',hex]];
function rank(name,query){
  const rows=refs.map(([n,d])=>({n,s:score(query,d)})).sort((a,b)=>b.s.raw-a.s.raw);
  console.log('\n'+name);
  rows.forEach(r=>console.log(r.n,r.s.pct+'%', 'raw='+r.s.raw.toFixed(3), 'L='+r.s.luma,'E='+r.s.edges,'P='+r.s.pattern,'C='+r.s.colour));
  return rows[0].n;
}
const exact=score(ring,ring);
console.log('Exact duplicate:',exact.pct+'%');
if(exact.pct!==99)throw new Error('Exact duplicate calibration failed');
const top1=rank('Ring query — moved, larger, brighter, different background',desc(makeImage('ring',{cx:55,cy:43,r:33,bg:[178,190,200],fg:[155,160,165]})));
const top2=rank('Gear query — different size and phase',desc(makeImage('gear',{cx:43,cy:50,r:31,phase:.3,bg:[230,225,210],fg:[95,110,120]})));
const top3=rank('Sensor query — shifted and brighter',desc(makeImage('sensor',{cx:52,cy:45,r:30,bg:[190,195,190],fg:[150,155,150]})));
if(top1!=='ring'||top2!=='gear'||top3!=='sensor')throw new Error(`Ranking failed: ${top1}, ${top2}, ${top3}`);
const flatA=desc(makeImage('none',{bg:[200,200,200]}));
const flatB=desc(makeImage('none',{bg:[220,210,200]}));
const flatScore=score(flatA,flatB);
console.log('\nTwo empty/flat photos:',flatScore.pct+'% (must not look like an exact match)');
if(flatScore.pct>60)throw new Error('Flat-image false exact match');
console.log('\nALL V5.6 SYNTHETIC TESTS PASSED');
