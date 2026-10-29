const fs=require('fs');
class Classes{constructor(){this.s=new Set()}add(x){this.s.add(x)}remove(x){this.s.delete(x)}contains(x){return this.s.has(x)}}
const ids=['lensRes','lensCmp','lensOv','lensSummary','lensStatus','lensList','cmpQuery','cmpRef','cmpRank','cmpName','cmpPart','cmpPct','cmpBars','cmpNote','lensVid','lensNoCam','lensShotImg','lensAfterBox','lensBar','lensHint','lensCam'];
const els={};for(const id of ids)els[id]={id,classList:new Classes(),style:{},textContent:'',innerHTML:'',src:'',value:'',click(){},removeAttribute(){}};
global.document={getElementById:id=>els[id]||(els[id]={id,classList:new Classes(),style:{},textContent:'',innerHTML:'',click(){},removeAttribute(){}}),hidden:false,body:{style:{}},createElement:()=>{throw new Error('Canvas not expected in UI test')}};
global.localStorage={getItem:()=>null,setItem:()=>{}};global.navigator={};global.$=id=>document.getElementById(id);
global.esc=s=>String(s==null?'':s).replace(/[<>&"]/g,c=>({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[c]));
global.driveSrc=()=>'';global.showToast=()=>{};global.APP_SCRIPT_URL='';global._byPart={};global.MATS=[];
let opened=-1;global.openDetail=i=>{opened=i};
const full=fs.readFileSync('/home/user/uploads/index.html','utf8');
const start=full.indexOf('/* ═══════════════ 📸 MATERIAL LENS');
const end=full.indexOf('/* ═══════════════ SYNC',start);
const src=full.slice(start,end);eval(src);
lensShotData='data:image/jpeg;base64,QUERY';
const sim={raw:.82,pct:82,luma:84,edges:79,pattern:76,colour:71,aspect:95,rotation:90,_rank:.808};
const row={x:{_i:7,n:'Cam Follower Bearing KR22',p:'262-030-00010',l:'Bay 1',m:'BM Stacking',img:''},sim,ref:{th:'data:image/jpeg;base64,REFERENCE'}};
_lensScanCount=1;_lensFilterText='Part filter: #262';
showLensResults([row]);
const html=els.lensList.innerHTML;
for(const token of ['YOUR PHOTO','ADMIN REFERENCE','Compare &amp; Confirm','82%','Cam Follower Bearing KR22']){
  if(!html.includes(token))throw new Error('Results missing '+token+'\n'+html);
}
if(!els.lensRes.classList.contains('show'))throw new Error('Results did not open');
lensCompare(0);
if(els.cmpName.textContent!==row.x.n)throw new Error('Compare part name failed');
if(!els.cmpBars.innerHTML.includes('Perceptual pattern')||!els.cmpBars.innerHTML.includes('Aspect ratio'))throw new Error('Breakdown missing');
if(!els.cmpNote.textContent.includes('90°'))throw new Error('Rotation note missing');
if(!els.lensCmp.classList.contains('show'))throw new Error('Comparison did not open');
lensConfirmCompare();
if(opened!==7)throw new Error('Confirm did not open correct material index');
if(els.lensCmp.classList.contains('show')||els.lensRes.classList.contains('show'))throw new Error('Overlays did not close');
console.log('V5.6 UI FLOW TEST PASSED');
console.log('Results cards → side-by-side comparison → score breakdown → correct part detail index 7');
