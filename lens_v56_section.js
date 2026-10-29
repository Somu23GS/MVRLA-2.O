/* ═══════════════ 📸 MATERIAL LENS v5.6 — visual comparison + multi-feature matching ═══════════════
   Admin మాత్రమే reference photos upload చేస్తారు. User photo background నుండి componentని
   వేరుచేసి luma + HOG edges + pHash + colour + aspect తో compare చేస్తుంది.
   ప్రతి resultలో user/reference photos పక్కపక్కనే ఉంటాయి; full comparison తర్వాతే confirm. */
var LENS_VERSION='5.6',LENS_CACHE_PREFIX='v56_';
var lensStream=null,lensShotData='',_lensCamOK=false;
var _lensRows=[],_lensCompareRow=null,_lensScanCount=0,_lensFilterText='';
var _lensCache={};try{_lensCache=JSON.parse(localStorage.getItem('mvrla_lens')||'{}')}catch(e){_lensCache={}}
/* పాత v1/v2 descriptors వాడకూడదు — v5.6కి సరికొత్త cache మాత్రమే */
for(var _oldLensKey in _lensCache){if(_oldLensKey.indexOf(LENS_CACHE_PREFIX)!==0)delete _lensCache[_oldLensKey]}
function lensCacheGet(id){var d=_lensCache[LENS_CACHE_PREFIX+id];return d&&d.v===56?d:null}
function lensSaveCache(){
  try{localStorage.setItem('mvrla_lens',JSON.stringify(_lensCache));return}
  catch(e){
    /* Quota దగ్గరికి వస్తే పాత entries తొలగించి మళ్లీ ప్రయత్నించు */
    var ks=Object.keys(_lensCache).filter(function(k){return k.indexOf(LENS_CACHE_PREFIX)===0});
    ks.sort(function(a,b){return Number(_lensCache[a]._ts||0)-Number(_lensCache[b]._ts||0)});
    for(var i=0;i<Math.max(10,Math.ceil(ks.length/3));i++)if(ks[i])delete _lensCache[ks[i]];
    try{localStorage.setItem('mvrla_lens',JSON.stringify(_lensCache))}catch(e2){}
  }
}
function lensCachePut(id,d){
  if(!d)return;
  d._ts=Date.now();_lensCache[LENS_CACHE_PREFIX+id]=d;
  var ks=Object.keys(_lensCache).filter(function(k){return k.indexOf(LENS_CACHE_PREFIX)===0});
  if(ks.length>80){
    ks.sort(function(a,b){return Number(_lensCache[a]._ts||0)-Number(_lensCache[b]._ts||0)});
    for(var i=0;i<ks.length-80;i++)delete _lensCache[ks[i]];
  }
  lensSaveCache();
}

async function openLens(){
  $('lensRes').classList.remove('show');$('lensCmp').classList.remove('show');
  $('lensOv').classList.add('show');document.body.style.overflow='hidden';
  _lensCamOK=false;
  if(navigator.mediaDevices&&navigator.mediaDevices.getUserMedia){
    try{
      lensStream=await navigator.mediaDevices.getUserMedia({audio:false,video:{facingMode:{ideal:'environment'},width:{ideal:1280}}});
      if(lensStream){$('lensVid').srcObject=lensStream;_lensCamOK=true}
    }catch(err){_lensCamOK=false}
  }
  lensResetUI();
}
function lensApplyCamUI(){
  var nocam=!_lensCamOK;
  $('lensVid').style.display=nocam?'none':'block';
  $('lensNoCam').style.display=nocam?'flex':'none';
}
function lensResetUI(){
  var s=$('lensShotImg');s.style.display='none';s.removeAttribute('src');
  lensApplyCamUI();$('lensAfterBox').classList.remove('show');$('lensBar').style.display='flex';
  $('lensHint').textContent=_lensCamOK?'Partని centreలో, plain backgroundతో, పూర్తిగా frameలో ఉంచండి':'';
  lensShotData='';
}
function closeLens(){
  if(lensStream){var ts=lensStream.getTracks();for(var i=0;i<ts.length;i++)ts[i].stop();lensStream=null}
  $('lensOv').classList.remove('show');$('lensRes').classList.remove('show');$('lensCmp').classList.remove('show');
  document.body.style.overflow='';
}
function lensCapture(){
  var v=$('lensVid');
  if(_lensCamOK&&lensStream&&v.videoWidth){
    var c=$('lensCan'),s=Math.min(1,720/v.videoWidth);
    c.width=Math.max(1,Math.round(v.videoWidth*s));c.height=Math.max(1,Math.round(v.videoHeight*s));
    c.getContext('2d').drawImage(v,0,0,c.width,c.height);
    lensSetShot(c.toDataURL('image/jpeg',.84));
  }else{$('lensCam').click()}
}
function lensPick(inp){
  var f=inp.files&&inp.files[0];inp.value='';if(!f)return;
  var mime=String(f.type||'').toLowerCase(),name=String(f.name||'');
  var byName=/\.(jpe?g|png|webp|gif|bmp|heic|heif)$/i.test(name);
  if((!mime||mime.indexOf('image/')!==0)&&!byName){showToast('ఒక image ఎంచుకోండి','error');return}
  var r=new FileReader();
  r.onload=function(e){compress(e.target.result,900).then(lensSetShot)};
  r.onerror=function(){showToast('Photo చదవలేకపోయింది — మరో photo ప్రయత్నించండి','error')};
  r.readAsDataURL(f);
}
function lensSetShot(dataUrl){
  lensShotData=dataUrl;var img=$('lensShotImg');img.src=dataUrl;img.style.display='block';
  $('lensNoCam').style.display='none';$('lensVid').style.display='none';
  $('lensBar').style.display='none';$('lensAfterBox').classList.add('show');$('lensHint').textContent='';
}
function lensRetake(){lensResetUI()}
function lensAgain(){$('lensRes').classList.remove('show');$('lensCmp').classList.remove('show');openLens()}

/* Reference thumbnail కూడా descriptorతో save అవుతుంది — comparisonలో Drive thumbnail failure వచ్చినా photo ఉంటుంది */
function lensMakeThumb(img){
  var S=144,c=document.createElement('canvas');c.width=S;c.height=S;
  var ctx=c.getContext('2d');ctx.fillStyle='#0A1628';ctx.fillRect(0,0,S,S);
  var iw=img.naturalWidth||img.videoWidth||img.width,ih=img.naturalHeight||img.videoHeight||img.height;
  if(!iw||!ih)return '';
  var sc=Math.min((S-8)/iw,(S-8)/ih),w=Math.max(1,iw*sc),h=Math.max(1,ih*sc);
  ctx.drawImage(img,(S-w)/2,(S-h)/2,w,h);
  try{return c.toDataURL('image/jpeg',.72)}catch(e){return ''}
}

/* 8×8 DCT cosine table — 63-bit perceptual hash కోసం */
var _lensDCTCos=null;
function lensCosines(){
  if(_lensDCTCos)return _lensDCTCos;
  _lensDCTCos=[];
  for(var u=0;u<8;u++){var row=[];for(var x=0;x<32;x++)row.push(Math.cos((2*x+1)*u*Math.PI/64));_lensDCTCos.push(row)}
  return _lensDCTCos;
}
function lensSample(px,W,H,x,y,ch){
  x=Math.max(0,Math.min(W-1,x));y=Math.max(0,Math.min(H-1,y));
  var x0=Math.floor(x),y0=Math.floor(y),x1=Math.min(W-1,x0+1),y1=Math.min(H-1,y0+1),fx=x-x0,fy=y-y0;
  var a=px[(y0*W+x0)*4+ch],b=px[(y0*W+x1)*4+ch],c=px[(y1*W+x0)*4+ch],d=px[(y1*W+x1)*4+ch];
  return (a+(b-a)*fx)*(1-fy)+(c+(d-c)*fx)*fy;
}

/* Pure descriptor function — browser canvas లేకుండా కూడా unit-test చేయవచ్చు */
function lensDescFromRGBA(px,W,H){
  var N=W*H,i,x,y,R,G,B,lum=new Float64Array(N);
  var border=Math.max(2,Math.round(Math.min(W,H)*.06));
  var br=0,bg=0,bb=0,bc=0;
  for(y=0;y<H;y++)for(x=0;x<W;x++){
    i=y*W+x;R=px[i*4];G=px[i*4+1];B=px[i*4+2];lum[i]=.299*R+.587*G+.114*B;
    if(x<border||x>=W-border||y<border||y>=H-border){br+=R;bg+=G;bb+=B;bc++}
  }
  br/=bc||1;bg/=bc||1;bb/=bc||1;
  var dm=0,ds=0;
  for(y=0;y<H;y++)for(x=0;x<W;x++)if(x<border||x>=W-border||y<border||y>=H-border){
    i=y*W+x;var dr=(px[i*4]-br)/255,dg=(px[i*4+1]-bg)/255,db=(px[i*4+2]-bb)/255;
    var dd=Math.sqrt((dr*dr+dg*dg+db*db)/3);dm+=dd;ds+=dd*dd;
  }
  dm/=bc||1;var dv=Math.max(0,ds/(bc||1)-dm*dm),thr=Math.max(.11,Math.min(.38,dm+2*Math.sqrt(dv)));
  var mask=new Uint8Array(N);
  for(i=0;i<N;i++){
    var d0=(px[i*4]-br)/255,d1=(px[i*4+1]-bg)/255,d2=(px[i*4+2]-bb)/255;
    mask[i]=Math.sqrt((d0*d0+d1*d1+d2*d2)/3)>thr?1:0;
  }
  /* 3×3 morphological cleanup — isolated background noise తొలగిస్తుంది */
  var sm=new Uint8Array(N);
  for(y=0;y<H;y++)for(x=0;x<W;x++){
    i=y*W+x;
    if(x===0||y===0||x===W-1||y===H-1){sm[i]=mask[i];continue}
    var count=0;for(var yy=-1;yy<=1;yy++)for(var xx=-1;xx<=1;xx++)count+=mask[(y+yy)*W+x+xx];
    sm[i]=mask[i]?(count>=3?1:0):(count>=6?1:0);
  }
  /* Central connected component — componentకి దగ్గరలోని background/shadow పెద్దదైతే వదిలేయి */
  var labels=new Int32Array(N);for(i=0;i<N;i++)labels[i]=-1;
  var stack=new Int32Array(N),best=null,bestScore=0;
  for(var seed=0;seed<N;seed++)if(sm[seed]&&labels[seed]<0){
    var sp=0;stack[sp++]=seed;labels[seed]=seed;
    var area=0,central=0,minX=W,maxX=0,minY=H,maxY=0;
    while(sp){
      var id=stack[--sp];x=id%W;y=(id/W)|0;area++;
      if(x<minX)minX=x;if(x>maxX)maxX=x;if(y<minY)minY=y;if(y>maxY)maxY=y;
      if(x>W*.25&&x<W*.75&&y>H*.25&&y<H*.75)central++;
      var nb;
      if(x>0){nb=id-1;if(sm[nb]&&labels[nb]<0){labels[nb]=seed;stack[sp++]=nb}}
      if(x<W-1){nb=id+1;if(sm[nb]&&labels[nb]<0){labels[nb]=seed;stack[sp++]=nb}}
      if(y>0){nb=id-W;if(sm[nb]&&labels[nb]<0){labels[nb]=seed;stack[sp++]=nb}}
      if(y<H-1){nb=id+W;if(sm[nb]&&labels[nb]<0){labels[nb]=seed;stack[sp++]=nb}}
    }
    var score=area+central*1.35;
    if(area>=Math.max(10,N*.012)&&score>bestScore){bestScore=score;best={area:area,minX:minX,maxX:maxX,minY:minY,maxY:maxY}}
  }
  var x0,y0,x1,y1;
  if(best&&best.area>N*.90){x0=0;y0=0;x1=W-1;y1=H-1}
  else if(best){
    var padX=(best.maxX-best.minX+1)*.12,padY=(best.maxY-best.minY+1)*.12;
    x0=Math.max(0,Math.floor(best.minX-padX));x1=Math.min(W-1,Math.ceil(best.maxX+padX));
    y0=Math.max(0,Math.floor(best.minY-padY));y1=Math.min(H-1,Math.ceil(best.maxY+padY));
  }else{
    /* Component కనిపించకపోతే centre crop — user guidance part centreలో ఉంటుంది */
    x0=Math.floor(W*.075);x1=Math.ceil(W*.925);y0=Math.floor(H*.075);y1=Math.ceil(H*.925);
  }
  if(x1-x0<3||y1-y0<3){x0=0;y0=0;x1=W-1;y1=H-1}
  var cw=x1-x0+1,ch=y1-y0+1,S=32,M=S*S,r32=new Float64Array(M),g32=new Float64Array(M),b32=new Float64Array(M),l32=new Float64Array(M);
  for(y=0;y<S;y++)for(x=0;x<S;x++){
    var sx=x0+(x+.5)*cw/S-.5,sy=y0+(y+.5)*ch/S-.5,j=y*S+x;
    var rr=lensSample(px,W,H,sx,sy,0),gg=lensSample(px,W,H,sx,sy,1),bb2=lensSample(px,W,H,sx,sy,2);
    r32[j]=rr;g32[j]=gg;b32[j]=bb2;l32[j]=.299*rr+.587*gg+.114*bb2;
  }
  /* Zero-mean, unit-normalized luma — bright background వల్ల పెద్ద score రాకూడదు */
  var mean=0;for(i=0;i<M;i++)mean+=l32[i];mean/=M;
  var norm=0;for(i=0;i<M;i++)norm+=(l32[i]-mean)*(l32[i]-mean);norm=Math.sqrt(norm)||1;
  var g=[];for(i=0;i<M;i++)g.push(+((l32[i]-mean)/norm).toFixed(4));
  /* HOG: 4×4 cells, 9 unsigned-orientation bins, 2×2 block normalization */
  var cells=new Float64Array(16*9);
  for(y=1;y<S-1;y++)for(x=1;x<S-1;x++){
    var ii=y*S+x,gx=l32[ii+1]-l32[ii-1],gy=l32[ii+S]-l32[ii-S];
    var mag=Math.sqrt(gx*gx+gy*gy);if(mag<1e-5)continue;
    var ang=(Math.atan2(gy,gx)*180/Math.PI+180)%180,bin=Math.min(8,(ang/20)|0),cell=(y>>3)*4+(x>>3);
    cells[cell*9+bin]+=mag;
  }
  var hog=[];
  for(var by=0;by<3;by++)for(var bx=0;bx<3;bx++){
    var hn=0;
    for(var dy=0;dy<2;dy++)for(var dx=0;dx<2;dx++){var ci=(by+dy)*4+(bx+dx);for(var hb=0;hb<9;hb++)hn+=cells[ci*9+hb]*cells[ci*9+hb]}
    hn=Math.sqrt(hn)+1e-6;
    for(dy=0;dy<2;dy++)for(dx=0;dx<2;dx++){ci=(by+dy)*4+(bx+dx);for(hb=0;hb<9;hb++)hog.push(cells[ci*9+hb]/hn)}
  }
  norm=0;for(i=0;i<hog.length;i++)norm+=hog[i]*hog[i];norm=Math.sqrt(norm)||1;
  for(i=0;i<hog.length;i++)hog[i]=+(hog[i]/norm).toFixed(4);
  /* pHash — scale/brightness changesలో patternని స్థిరంగా compare చేయడానికి */
  var C=lensCosines(),coef=[];
  for(var v=0;v<8;v++)for(var u=0;u<8;u++){
    var sum=0;for(y=0;y<S;y++)for(x=0;x<S;x++)sum+=l32[y*S+x]*C[u][x]*C[v][y];coef.push(sum);
  }
  var ac=coef.slice(1).sort(function(a,b){return a-b}),med=ac[(ac.length/2)|0]||0,ph=[];
  for(i=1;i<coef.length;i++)ph.push(coef[i]>med?1:0);
  var hist=[];for(i=0;i<64;i++)hist.push(0);
  for(i=0;i<M;i++){
    var bi=(Math.min(3,(r32[i]/64)|0)<<4)|(Math.min(3,(g32[i]/64)|0)<<2)|Math.min(3,(b32[i]/64)|0);hist[bi]++;
  }
  for(i=0;i<64;i++)hist[i]=+(hist[i]/M).toFixed(5);
  var ar=cw/ch;ar=Math.min(ar,1/ar);
  return {v:56,g:g,hog:hog,ph:ph,h:hist,ar:+ar.toFixed(4)};
}
function imgDesc(img){
  var iw=img.naturalWidth||img.videoWidth||img.width,ih=img.naturalHeight||img.videoHeight||img.height;
  if(!iw||!ih)return null;
  var sc=Math.min(1,96/Math.max(iw,ih)),W=Math.max(8,Math.round(iw*sc)),H=Math.max(8,Math.round(ih*sc));
  var c=document.createElement('canvas');c.width=W;c.height=H;
  var ctx=c.getContext('2d',{willReadFrequently:true});ctx.fillStyle='#fff';ctx.fillRect(0,0,W,H);ctx.drawImage(img,0,0,W,H);
  try{var d=lensDescFromRGBA(ctx.getImageData(0,0,W,H).data,W,H);d.th=lensMakeThumb(img);return d}catch(e){return null}
}
function lensDot(a,b){var s=0,n=Math.min(a.length,b.length);for(var i=0;i<n;i++)s+=a[i]*b[i];return s}
function lensSim(a,b){
  var luma=Math.max(0,Math.min(1,lensDot(a.g,b.g)));
  var edges=Math.max(0,Math.min(1,lensDot(a.hog,b.hog)));
  var same=0;for(var i=0;i<a.ph.length;i++)if(a.ph[i]===b.ph[i])same++;
  var pattern=Math.max(0,Math.min(1,((same/a.ph.length)-.5)*2));
  var colour=0;for(i=0;i<64;i++)colour+=Math.min(a.h[i],b.h[i]);
  var aspect=Math.min(a.ar,b.ar)/(Math.max(a.ar,b.ar)||1);
  var raw=.30*luma+.35*edges+.20*pattern+.10*colour+.05*aspect;
  return {raw:raw,pct:Math.max(0,Math.min(99,Math.round(raw*100))),
    luma:Math.round(luma*100),edges:Math.round(edges*100),pattern:Math.round(pattern*100),
    colour:Math.round(colour*100),aspect:Math.round(aspect*100),rotation:0};
}
function loadImgEl(src){return new Promise(function(res){var im=new Image();im.onload=function(){res(im)};im.onerror=function(){res(null)};im.src=src})}
function lensRotateForDesc(img,deg){
  var iw=img.naturalWidth||img.width,ih=img.naturalHeight||img.height,sc=Math.min(1,160/Math.max(iw,ih));
  iw=Math.max(8,Math.round(iw*sc));ih=Math.max(8,Math.round(ih*sc));
  var c=document.createElement('canvas');c.width=deg===180?iw:ih;c.height=deg===180?ih:iw;
  var ctx=c.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,c.width,c.height);
  if(deg===90){ctx.translate(c.width,0);ctx.rotate(Math.PI/2)}
  else if(deg===180){ctx.translate(c.width,c.height);ctx.rotate(Math.PI)}
  else{ctx.translate(0,c.height);ctx.rotate(-Math.PI/2)}
  ctx.drawImage(img,0,0,iw,ih);return c;
}
function lensQueryVariants(img){
  var out=[{d:imgDesc(img),rot:0}];
  /* 90/180/270 alternatives — reference orientationకి user photo దగ్గరగా align అవుతుంది */
  var angles=[90,180,270];
  for(var i=0;i<angles.length;i++){var r=lensRotateForDesc(img,angles[i]);out.push({d:imgDesc(r),rot:angles[i]})}
  return out.filter(function(v){return !!v.d});
}
async function lensGetReference(cd){
  var cached=lensCacheGet(cd.fid);if(cached)return cached;
  var di='';
  try{
    var r=await fetch(APP_SCRIPT_URL+'?action=getimg&id='+encodeURIComponent(cd.fid)),j=await r.json();
    if(j&&j.ok&&j.dataUrl)di=j.dataUrl;
  }catch(e){}
  if(!di)return null;
  var el=await loadImgEl(di);if(!el)return null;
  var d=imgDesc(el);if(d)lensCachePut(cd.fid,d);return d;
}

async function findLensMatches(){
  if(!lensShotData){showToast('ముందు photo తీయండి 📸','error');return}
  var hint=$('lensHint');$('lensAfterBox').classList.remove('show');$('lensBar').style.display='none';
  hint.textContent='🔍 v5.6 · Component analysis…';
  var shotEl=await loadImgEl(lensShotData);
  if(!shotEl){hint.textContent='Photo చదవలేకపోయింది — మళ్లీ తీయండి';$('lensAfterBox').classList.add('show');return}
  var variants=lensQueryVariants(shotEl);
  if(!variants.length){hint.textContent='Photo process కాలేదు — మరో photo ప్రయత్నించండి';$('lensAfterBox').classList.add('show');return}
  var cands=[],seen={},k;
  for(k in _byPart){var x=_byPart[k],fid=driveId(x.img);if(fid&&!seen[x.p]){seen[x.p]=1;cands.push({x:x,fid:fid})}}
  for(var i=0;i<MATS.length;i++){var m2=MATS[i],f2=driveId(m2.img);if(f2&&!seen[m2.p]){seen[m2.p]=1;var xi=_byPart[m2.p];cands.push({x:xi||m2,fid:f2})}}
  var pre='';try{pre=($('lensPrefix').value||'').replace(/[^0-9]/g,'')}catch(e){}
  _lensFilterText=pre?'Part filter: #'+pre:'All admin references';
  if(pre&&cands.length){
    var filtered=cands.filter(function(cd){return String(cd.x.p).replace(/[^0-9]/g,'').indexOf(pre)>=0});
    if(filtered.length)cands=filtered;else _lensFilterText='#'+pre+' not found · all references used';
  }
  _lensScanCount=cands.length;
  if(!cands.length){showLensResults([]);return}
  var rows=[];
  /* మూడు references parallel — వేగం పెరుగుతుంది, phone మీద అధిక load పడదు */
  for(var start=0;start<cands.length;start+=3){
    var batch=cands.slice(start,start+3);
    hint.textContent='🔍 Comparing '+Math.min(start+3,cands.length)+' / '+cands.length+' references…';
    var got=await Promise.all(batch.map(async function(cd){
      try{var d=await lensGetReference(cd);return d?{cd:cd,d:d}:null}catch(e){return null}
    }));
    for(var gi=0;gi<got.length;gi++)if(got[gi]){
      var cd2=got[gi].cd,rd=got[gi].d,best=null;
      for(var vi=0;vi<variants.length;vi++){
        var sim=lensSim(variants[vi].d,rd),rank=sim.raw-(variants[vi].rot?.012:0);
        if(!best||rank>best._rank){best=sim;best._rank=rank;best.rotation=variants[vi].rot;best.pct=Math.max(0,Math.min(99,Math.round(rank*100)))}
      }
      rows.push({x:cd2.x,sim:best,ref:rd});
    }
    /* UI paintకి చిన్న break — matching సమయంలో interface freeze అవదు */
    await new Promise(function(rs){setTimeout(rs,0)});
  }
  rows.sort(function(a,b){return b.sim._rank-a.sim._rank});
  showLensResults(rows.slice(0,5));
}

function lensScoreClass(pct){return pct>=75?'g':(pct>=50?'a':'b')}
function lensImgTag(src,alt){
  return '<img src="'+esc(src||'')+'" alt="'+esc(alt||'')+'" referrerpolicy="no-referrer" onerror="this.style.opacity=0">';
}
function showLensResults(rows){
  $('lensOv').classList.remove('show');$('lensCmp').classList.remove('show');
  _lensRows=rows||[];var R=$('lensRes');R.classList.add('show');
  $('lensSummary').innerHTML='<div class="lq-img">'+lensImgTag(lensShotData,'Your photo')+'</div>'
    +'<div class="lq-info"><span>YOUR PHOTO</span><b>'+_lensScanCount+' admin references compared</b><small>'+esc(_lensFilterText)+' · v5.6 multi-feature analysis</small></div>';
  if(!rows.length){
    $('lensStatus').style.display='none';
    $('lensList').innerHTML='<div class="lres-empty">😕 '+(_lensScanCount?'Reference photos load కాలేవు. Internet / getimg deployment చెక్ చేయండి.':'Admin reference photos ఇంకా లేవు.')+'<br>Part number / nameతో search కూడా వాడవచ్చు.</div>';
    return;
  }
  var top=rows[0].sim.pct,margin=rows.length>1?top-rows[1].sim.pct:top;
  var st=$('lensStatus');st.style.display='flex';
  if(top>=75&&margin>=6){st.className='lres-status good';st.textContent='Strongest visual candidate is #1. Compare both photos and confirm the part number.'}
  else if(top<50){st.className='lres-status low';st.textContent='Low visual similarity. Retake with a plain background or enter a known part-start number.'}
  else{st.className='lres-status close';st.textContent='Close or moderate scores. Do not auto-select — open the comparison and verify carefully.'}
  var html='';
  for(var i=0;i<rows.length;i++){
    var row=rows[i],x=row.x,sim=row.sim,pct=sim.pct,cls=lensScoreClass(pct);
    var refSrc=row.ref&&row.ref.th?row.ref.th:driveSrc(x.img,240);
    html+='<article class="lm-card">'
      +'<div class="lm-cardtop"><span class="lm-rank">#'+(i+1)+(i===0?' · BEST CANDIDATE':' CANDIDATE')+'</span><span class="lm-pct '+cls+'">'+pct+'%<small>SIMILARITY</small></span></div>'
      +'<div class="lm-pair">'
        +'<div class="lm-side"><div class="lm-thumb">'+lensImgTag(lensShotData,'Your photo')+'</div><span>YOUR PHOTO</span></div>'
        +'<div class="lm-vs">VS</div>'
        +'<div class="lm-side"><div class="lm-thumb">'+lensImgTag(refSrc,'Admin reference')+'</div><span>ADMIN REFERENCE</span></div>'
      +'</div>'
      +'<div class="lm-meta"><b>'+esc(x.n)+'</b><span>'+esc(x.p)+(x.l?' · '+esc(x.l):'')+'</span></div>'
      +'<div class="lm-chips"><span>Luma '+sim.luma+'%</span><span>Edges '+sim.edges+'%</span><span>Pattern '+sim.pattern+'%</span><span>Colour '+sim.colour+'%</span></div>'
      +'<button class="lm-open" onclick="lensCompare('+i+')">Compare & Confirm <b>→</b></button>'
      +'</article>';
  }
  $('lensList').innerHTML=html;
}
function lensCompare(index){
  var row=_lensRows[index];if(!row)return;
  _lensCompareRow=row;var x=row.x,sim=row.sim,pct=sim.pct;
  $('cmpQuery').src=lensShotData;$('cmpQuery').style.opacity=1;
  var ref=row.ref&&row.ref.th?row.ref.th:driveSrc(x.img,600);$('cmpRef').src=ref;$('cmpRef').style.opacity=1;
  $('cmpRank').textContent='CANDIDATE #'+(index+1)+' OF '+_lensRows.length;
  $('cmpName').textContent=x.n||'Unnamed material';
  $('cmpPart').textContent=x.p+(x.l?' · '+x.l:'')+(x.m?' · '+x.m:'');
  var badge=$('cmpPct');badge.className='cmp-pct '+lensScoreClass(pct);badge.innerHTML=pct+'%<small>VISUAL SIMILARITY</small>';
  var metrics=[['Luma structure',sim.luma],['Edge geometry',sim.edges],['Perceptual pattern',sim.pattern],['Colour overlap',sim.colour],['Aspect ratio',sim.aspect]],html='';
  for(var i=0;i<metrics.length;i++){
    html+='<div class="cm-row"><span>'+metrics[i][0]+'</span><div class="cm-track"><i style="width:'+Math.max(0,Math.min(100,metrics[i][1]))+'%"></i></div><b>'+metrics[i][1]+'%</b></div>';
  }
  $('cmpBars').innerHTML=html;
  var note=pct>=75?'Strong visual similarity — part number మరియు component details కూడా verify చేయండి.':(pct>=50?'Possible match — shape, label మరియు part number జాగ్రత్తగా compare చేయండి.':'Low similarity — దీన్ని confirm చేయకముందు మరో photo / part-start filter ప్రయత్నించండి.');
  if(sim.rotation)note+=' Comparison కోసం '+sim.rotation+'° rotation alignment వాడబడిందi.';
  $('cmpNote').textContent=note;
  $('lensCmp').classList.add('show');
}
function lensCloseCompare(){$('lensCmp').classList.remove('show')}
function lensConfirmCompare(){
  if(!_lensCompareRow)return;var i=_lensCompareRow.x._i;
  $('lensCmp').classList.remove('show');closeLens();openDetail(i);
}

/* ⚡ v5.6 background warm-up — descriptor + reference thumbnail రెండూ pre-build అవుతాయి */
var _lensWarmT=null;
function lensWarm(){
  if(_lensWarmT)return;
  _lensWarmT=setTimeout(async function(){
    var done=0;
    for(var k in _byPart){
      var x=_byPart[k],fid=driveId(x.img);if(!fid||lensCacheGet(fid))continue;
      if(done>=80)break;
      var d=await lensGetReference({fid:fid});if(d)done++;
      await new Promise(function(rs){setTimeout(rs,450)});
      if(document.hidden)break;
    }
  },6000);
}

