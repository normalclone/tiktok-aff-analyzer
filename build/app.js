/* ============================================================
   TikTok Affiliate Analyzer — logic
   Đầu vào: file .xlsx export đơn hàng affiliate TikTok Shop
   ============================================================ */
'use strict';

/* ---------- Bộ icon SVG đồng bộ (Feather/Lucide style, 24x24 stroke) ---------- */
const ICONS={
  dollar:'<line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>',
  target:'<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
  check:'<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>',
  bar:'<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>',
  rotate:'<polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/>',
  activity:'<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>',
  clock:'<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
  bag:'<path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z"/><line x1="3" y1="6" x2="21" y2="6"/><path d="M16 10a4 4 0 0 1-8 0"/>',
  package:'<line x1="16.5" y1="9.4" x2="7.5" y2="4.21"/><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/>',
  alert:'<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>',
  search:'<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>',
  compare:'<circle cx="18" cy="18" r="3"/><circle cx="6" cy="6" r="3"/><path d="M13 6h3a2 2 0 0 1 2 2v7"/><path d="M11 18H8a2 2 0 0 1-2-2V9"/>',
  download:'<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>',
  upload:'<polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/>',
  refresh:'<polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>',
  filter:'<polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>',
  chevron:'<polyline points="6 9 12 15 18 9"/>',
  tick:'<polyline points="20 6 9 17 4 12"/>',
  clipboard:'<path d="M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2"/><rect x="9" y="3" width="6" height="4" rx="1"/><path d="m9 14 2 2 4-4"/>',
  bulb:'<path d="M9 18h6"/><path d="M10 22h4"/><path d="M15.09 14c.18-.98.65-1.74 1.41-2.5A4.65 4.65 0 0 0 18 8 6 6 0 0 0 6 8c0 1 .23 2.23 1.5 3.5A4.61 4.61 0 0 1 8.91 14"/>',
  video:'<polygon points="23 7 16 12 23 17 23 7"/><rect x="1" y="5" width="15" height="14" rx="2" ry="2"/>',
  link:'<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>'
};
function icon(name,size){const p=ICONS[name]||'';return `<svg class="ik" viewBox="0 0 24 24"${size?` style="width:${size}px;height:${size}px"`:''}>${p}</svg>`;}
function injectIcons(root=document){root.querySelectorAll('[data-icon]').forEach(el=>{el.innerHTML=icon(el.dataset.icon);});}

/* ---------- Tiện ích ---------- */
const VN=new Intl.NumberFormat('vi-VN');
const fmt   = n => (n==null||isNaN(n)) ? '—' : VN.format(Math.round(n));
const fmtD  = (n,d=1) => (n==null||isNaN(n)) ? '—' : VN.format(+(n).toFixed(d));
const pct   = (a,b) => (!b) ? '0%' : (a/b*100).toFixed(1)+'%';
const money = n => fmt(n)+' đ';
const moneyShort = n => {
  if(n==null||isNaN(n)) return '—';
  const a=Math.abs(n);
  if(a>=1e9) return fmtD(n/1e9,2)+' tỷ';
  if(a>=1e6) return fmtD(n/1e6,1)+' tr';
  if(a>=1e3) return fmtD(n/1e3,0)+'k';
  return fmt(n);
};
const esc = s => String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function rgba(hex,a){const h=hex.replace('#','');const n=parseInt(h,16);return `rgba(${(n>>16)&255},${(n>>8)&255},${n&255},${a})`;}
const clip=(s,n=40)=>{s=String(s||'');return s.length>n?s.slice(0,n)+'…':s;};
const fpct=x=>(x*100).toFixed(1)+'%';

/* Parse số kiểu Việt: "140.000"->140000, "1.234,56"->1234.56, "10%"->10 */
function toNum(v){
  if(v==null) return NaN;
  if(typeof v==='number') return v;
  let s=String(v).trim().replace(/\s/g,'').replace('%','');
  if(s===''||s==='nan'||s==='/'||s==='-'||s==='N/A') return NaN;
  const hasDot=s.includes('.'), hasComma=s.includes(',');
  if(hasDot&&hasComma) s=s.replace(/\./g,'').replace(',','.');
  else if(hasComma)    s=s.replace(/\./g,'').replace(',','.');
  else                 s=s.replace(/\./g,'');
  const n=parseFloat(s);
  return isNaN(n)?NaN:n;
}
const sum=(arr,f)=>arr.reduce((a,r)=>{const v=f(r);return a+(isNaN(v)?0:v);},0);

/* Parse ngày "30/06/2026 21:15:38" */
function parseDay(v){
  if(v==null) return null;
  const m=String(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(!m) return null;
  const [_,d,mo,y]=m;
  return {key:`${y}-${mo.padStart(2,'0')}-${d.padStart(2,'0')}`, label:`${d.padStart(2,'0')}/${mo.padStart(2,'0')}`};
}

/* ---------- Dò cột theo tên tiếng Việt ---------- */
const COLDEF={
  orderId:['ID đơn hàng'], product:['Tên sản phẩm'], productId:['ID sản phẩm'],
  price:['Giá'], sold:['Số món bán ra'], refund:['Số món đã hoàn tiền'], shop:['Tên cửa hàng'],
  orderType:['Loại đơn hàng'], status:['Trạng thái quyết toán đơn hàng'], content:['Loại nội dung'],
  gmv:['GMV'], estComm:['Hoa hồng tiêu chuẩn ước tính'], actComm:['Hoa hồng tiêu chuẩn'],
  finalRecv:['Tổng số tiền nhận được cuối cùng'], orderDate:['Ngày đặt hàng']
};
const norm=s=>String(s||'').toLowerCase().replace(/\s+/g,' ').trim();
function buildColMap(headers){
  const map={},H=headers.map(norm);
  for(const key in COLDEF){
    let idx=-1;
    for(const name of COLDEF[key]){
      const t=norm(name);
      idx=H.indexOf(t);
      if(idx<0) idx=H.findIndex(h=>h.includes(t)||t.includes(h));
      if(idx>=0) break;
    }
    map[key]=idx;
  }
  return map;
}

/* ---------- Trạng thái: màu TikTok ---------- */
const STATUS_META={
  'Đã quyết toán':     {cls:'ok',   color:'#25f4ee'},
  'Không đủ điều kiện': {cls:'bad',  color:'#fe2c55'},
  'Chờ xử lý':         {cls:'pend', color:'#5b9dff'},
  'AwaitingPayment':   {cls:'wait', color:'#ffb02e'}
};
const statusMeta=s=>STATUS_META[s]||{cls:'pend',color:'#a8a8b3'};
const PALETTE=['#25f4ee','#fe2c55','#ffb02e','#5b9dff','#b794ff','#39e29a','#ff8a5b','#f0e14a'];
const C_CYAN='#25f4ee',C_RED='#fe2c55',C_AMBER='#ffb02e',C_GREEN='#39e29a';

/* ---------- State ---------- */
let RAW=[], COLMAP=null, CHARTS={}, META={}, PICKERS={};

/* ============================================================
   Searchable Picker (full-width, có keyboard nav)
   ============================================================ */
const _openPickers=new Set();
function closeAllPickers(except){_openPickers.forEach(p=>{if(p!==except)p.close();});}
class Picker{
  constructor(mount,{onChange}){
    this.el=mount; this.onChange=onChange; this.options=[]; this.value='__ALL__'; this.allLabel='Tất cả'; this.hl=-1;
    this.render();
  }
  render(){
    this.el.classList.add('picker');
    this.el.innerHTML=
      `<button type="button" class="pk-trigger"><span class="pk-val"></span><span class="ico chev">${icon('chevron')}</span></button>`+
      `<div class="pk-panel hidden">`+
        `<div class="pk-search"><span class="ico">${icon('search')}</span><input type="text" placeholder="Tìm…" autocomplete="off"></div>`+
        `<div class="pk-list" role="listbox"></div>`+
      `</div>`;
    this.trigger=this.el.querySelector('.pk-trigger');
    this.panel=this.el.querySelector('.pk-panel');
    this.search=this.el.querySelector('input');
    this.list=this.el.querySelector('.pk-list');
    this.valEl=this.el.querySelector('.pk-val');
    this.trigger.addEventListener('click',e=>{e.stopPropagation();this.toggle();});
    this.search.addEventListener('input',()=>{this.hl=0;this.renderList();});
    this.search.addEventListener('keydown',e=>this.onKey(e));
  }
  setOptions(opts,allLabel){
    this.allLabel=allLabel; this.value='__ALL__';
    this.options=[{v:'__ALL__',t:allLabel},...opts.map(o=>({v:o,t:o}))];
    this.updateTrigger();
  }
  toggle(){this.panel.classList.contains('hidden')?this.open():this.close();}
  open(){
    closeAllPickers(this); _openPickers.add(this);
    this.panel.classList.remove('hidden'); this.el.classList.add('open');
    this.search.value=''; this.hl=this.options.findIndex(o=>o.v===this.value); this.renderList();
    this.search.focus();
  }
  close(){this.panel.classList.add('hidden'); this.el.classList.remove('open'); _openPickers.delete(this);}
  filtered(){const q=norm(this.search.value); return this.options.filter(o=>!q||norm(o.t).includes(q));}
  renderList(){
    const items=this.filtered();
    if(this.hl>=items.length) this.hl=items.length-1;
    if(this.hl<0 && items.length) this.hl=0;
    this.list.innerHTML=items.length
      ? items.map((o,i)=>`<div class="pk-opt ${o.v===this.value?'sel':''} ${i===this.hl?'hl':''}" data-v="${esc(o.v)}" role="option">`+
          `<span class="opt-t">${esc(o.t)}</span><span class="ico">${icon('tick')}</span></div>`).join('')
      : `<div class="pk-empty">Không tìm thấy</div>`;
    this.list.querySelectorAll('.pk-opt').forEach(el=>{
      el.addEventListener('click',()=>this.select(el.dataset.v));
      el.addEventListener('mousemove',()=>{const items=this.filtered();this.hl=items.findIndex(o=>o.v===el.dataset.v);this.highlight();});
    });
    this.scrollToHl();
  }
  highlight(){this.list.querySelectorAll('.pk-opt').forEach((el,i)=>el.classList.toggle('hl',i===this.hl));}
  scrollToHl(){const el=this.list.querySelector('.pk-opt.hl'); if(el&&el.scrollIntoView)el.scrollIntoView({block:'nearest'});}
  onKey(e){
    const items=this.filtered();
    if(e.key==='ArrowDown'){e.preventDefault();this.hl=Math.min(this.hl+1,items.length-1);this.highlight();this.scrollToHl();}
    else if(e.key==='ArrowUp'){e.preventDefault();this.hl=Math.max(this.hl-1,0);this.highlight();this.scrollToHl();}
    else if(e.key==='Enter'){e.preventDefault();if(items[this.hl])this.select(items[this.hl].v);}
    else if(e.key==='Escape'){e.preventDefault();this.close();this.trigger.focus();}
  }
  select(v){this.value=v;this.updateTrigger();this.close();this.trigger.focus();this.onChange&&this.onChange(v);}
  updateTrigger(){const o=this.options.find(o=>o.v===this.value);this.valEl.textContent=o?o.t:this.allLabel;this.valEl.classList.toggle('is-all',this.value==='__ALL__');}
  get(){return this.value;}
}
document.addEventListener('click',e=>{_openPickers.forEach(p=>{if(!p.el.contains(e.target))p.close();});});

/* ============================================================
   Đọc file
   ============================================================ */
// Parse + render từ một ArrayBuffer (dùng chung cho cả file tải lên lẫn link)
function processBuffer(buf, fileName){
  const fail=msg=>{hideLoading();showErr(msg);};
  /* Hoãn 1 nhịp để trình duyệt kịp vẽ loader (spinner xoay trên compositor)
     trước khi XLSX.read chiếm luồng chính → có cảm giác "đang tải", không treo đơ */
  setTimeout(()=>{
    try{
      const wb=XLSX.read(buf,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:''});
      if(!rows.length) return fail('File rỗng.');
      const headers=rows[0].map(h=>String(h).trim());
      COLMAP=buildColMap(headers);
      if(COLMAP.gmv<0 && COLMAP.price<0)
        return fail('Không nhận diện được cột GMV/Giá. File có đúng định dạng export đơn hàng affiliate TikTok không?');
      const C=COLMAP, get=(r,k)=> C[k]>=0 ? r[C[k]] : '';
      RAW=[];
      for(let i=1;i<rows.length;i++){
        const r=rows[i];
        if(!r || r.every(c=>c===''||c==null)) continue;
        const gmv=toNum(get(r,'gmv')), price=toNum(get(r,'price'));
        RAW.push({
          orderId:String(get(r,'orderId')||'').trim(),
          product:String(get(r,'product')||'(không tên)').trim()||'(không tên)',
          productId:String(get(r,'productId')||'').trim(),
          shop:String(get(r,'shop')||'(không rõ)').trim()||'(không rõ)',
          orderType:String(get(r,'orderType')||'(không rõ)').trim()||'(không rõ)',
          status:String(get(r,'status')||'(không rõ)').trim()||'(không rõ)',
          content:String(get(r,'content')||'(không rõ)').trim()||'(không rõ)',
          sold:toNum(get(r,'sold'))||0, refund:toNum(get(r,'refund'))||0,
          gmv:isNaN(gmv)?(isNaN(price)?0:price):gmv,
          estComm:toNum(get(r,'estComm'))||0, actComm:toNum(get(r,'actComm'))||0,
          finalRecv:toNum(get(r,'finalRecv'))||0, day:parseDay(get(r,'orderDate'))
        });
      }
      if(!RAW.length) return fail('Không có dòng dữ liệu nào sau tiêu đề.');
      META.fileName=fileName||'dữ liệu'; META.totalRows=RAW.length;
      initFilters(); render();
      document.getElementById('uploadArea').classList.add('hidden');
      document.getElementById('dash').classList.remove('hidden');
      document.getElementById('btnSample').classList.remove('hidden');
      document.getElementById('btnCsv').classList.remove('hidden');
      hideLoading();
    }catch(ex){ fail('Lỗi xử lý file: '+ex.message); console.error(ex); }
  },60);
}
function handleFile(file){
  hideErr(); showLoading('Đang đọc file…');
  const reader=new FileReader();
  reader.onerror=()=>{hideLoading();showErr('Không đọc được file.');};
  reader.onload=e=>{ setLoadingSub('Đang phân tích '+(file.name||'')+'…'); processBuffer(e.target.result, file.name); };
  reader.readAsArrayBuffer(file);
}
async function handleUrl(rawUrl){
  const url=(rawUrl||'').trim();
  if(!url){ showErr('Hãy dán link file .xlsx trước.'); return; }
  if(!/^https?:\/\//i.test(url)){ showErr('Link phải bắt đầu bằng http:// hoặc https://'); return; }
  hideErr(); showLoading('Đang tải từ link…');
  try{
    const res=await fetch(url);
    if(!res.ok) throw new Error('HTTP '+res.status);
    const buf=await res.arrayBuffer();
    let name=decodeURIComponent((url.split('?')[0].split('/').pop())||'').trim();
    if(!/\.(xlsx|xls)$/i.test(name)) name=(name||'tu-link')+'.xlsx';
    setLoadingSub('Đang phân tích '+name+'…');
    processBuffer(buf, name);
  }catch(ex){
    hideLoading();
    showErr('Không tải được từ link ('+ex.message+'). Link có thể đã hết hạn, hoặc nguồn chặn truy cập trực tiếp (CORS). Cách chắc chắn: tải file về máy rồi kéo-thả vào ô phía trên.');
  }
}
function showLoading(sub){document.getElementById('ldSub').textContent=sub||'Đọc & phân tích file';document.getElementById('loading').classList.add('show');}
function setLoadingSub(t){document.getElementById('ldSub').textContent=t;}
function hideLoading(){document.getElementById('loading').classList.remove('show');}

/* ============================================================
   Bộ lọc
   ============================================================ */
function uniq(key){return [...new Set(RAW.map(r=>r[key]))].filter(x=>x&&x!=='(không rõ)').sort((a,b)=>a.localeCompare(b,'vi'));}
function initFilters(){
  const mk=(id,onCh)=>{ if(!PICKERS[id]) PICKERS[id]=new Picker(document.getElementById(id),{onChange:onCh}); return PICKERS[id]; };
  mk('fShop',render).setOptions(uniq('shop'),'Tất cả cửa hàng');
  mk('fStatus',render).setOptions(uniq('status'),'Tất cả trạng thái');
  mk('fContent',render).setOptions(uniq('content'),'Tất cả nội dung');
  mk('fOrder',render).setOptions(uniq('orderType'),'Tất cả loại đơn');
}
function getFiltered(){
  const g=id=>PICKERS[id].get();
  const s=g('fShop'),st=g('fStatus'),c=g('fContent'),o=g('fOrder');
  return RAW.filter(r=>
    (s==='__ALL__'||r.shop===s)&&(st==='__ALL__'||r.status===st)&&
    (c==='__ALL__'||r.content===c)&&(o==='__ALL__'||r.orderType===o));
}

/* ============================================================
   Tổng hợp + render
   ============================================================ */
function groupBy(arr,keyFn){const m=new Map();for(const r of arr){const k=keyFn(r);if(!m.has(k))m.set(k,[]);m.get(k).push(r);}return m;}
function agg(rows){
  return {
    rows:rows.length, orders:new Set(rows.map(r=>r.orderId).filter(Boolean)).size,
    gmv:sum(rows,r=>r.gmv), sold:sum(rows,r=>r.sold), refund:sum(rows,r=>r.refund),
    est:sum(rows,r=>r.estComm), act:sum(rows,r=>r.actComm), final:sum(rows,r=>r.finalRecv),
    settled:rows.filter(r=>r.status==='Đã quyết toán').length
  };
}
function render(){
  const D=getFiltered(), A=agg(D);
  document.getElementById('metaInfo').innerHTML=`Đang phân tích <b>${fmt(A.rows)}</b> dòng · <b>${fmt(A.orders)}</b> đơn duy nhất`;
  renderKPIs(A); renderSummary(D,A); renderStatus(D); renderContent(D); renderDay(D);
  renderShops(D); renderProducts(D); renderIneligible(D); renderGap(D);
  document.getElementById('foot').textContent=
    `Nguồn: ${META.fileName} · ${fmt(META.totalRows)} dòng gốc · đang hiển thị ${fmt(D.length)} dòng theo bộ lọc.`;
}

function renderKPIs(A){
  const refundRate=A.sold? A.refund/A.sold*100:0;
  const settleRate=A.rows? A.settled/A.rows*100:0;
  const eff=A.est? A.act/A.est*100:0;
  const cards=[
    {ic:'dollar',col:C_CYAN, lab:'Tổng GMV', val:money(A.gmv), sub:`${fmt(A.orders)} đơn · ${fmt(A.sold)} món bán`},
    {ic:'target',col:'#ffffff',lab:'HH ước tính', val:money(A.est), sub:`bình quân ${pct(A.est,A.gmv)} GMV`},
    {ic:'check', col:C_CYAN, lab:'HH thực nhận', val:money(A.act), sub:`tiền nhận cuối: ${moneyShort(A.final)} đ`},
    {ic:'bar',   col:'#ffffff',lab:'Tỷ lệ quyết toán', val:settleRate.toFixed(1)+'%', sub:`${fmt(A.settled)}/${fmt(A.rows)} dòng đã quyết toán`},
    {ic:'rotate',col:C_RED,  lab:'Tỷ lệ hoàn', val:refundRate.toFixed(2)+'%', sub:`${fmt(A.refund)} món hoàn`, cls:refundRate>8?'neg':''},
    {ic:'activity',col:eff<70?C_AMBER:C_CYAN, lab:'Hiệu suất HH', val:eff.toFixed(1)+'%', sub:'thực nhận / ước tính', cls:eff<70?'warn':'pos'}
  ];
  document.getElementById('kpis').innerHTML=cards.map(k=>{
    const chip=k.col==='#ffffff'?'background:rgba(255,255,255,.1);color:#fff':`background:${rgba(k.col,.15)};color:${k.col}`;
    return `<div class="kpi">
      <div class="kpi-top"><span class="kpi-ico" style="${chip}">${icon(k.ic)}</span><span class="kpi-lab">${k.lab}</span></div>
      <div class="kpi-val ${k.cls||''}">${k.val}</div>
      <div class="kpi-sub">${k.sub}</div></div>`;
  }).join('');
}

/* ============================================================
   Tổng kết: đánh giá theo tiêu chí + khuyến nghị (tự sinh theo dữ liệu)
   ============================================================ */
function renderSummary(D,A){
  // ----- gom số liệu -----
  const shopAgg=[...groupBy(D,r=>r.shop).entries()].map(([k,v])=>({k,
    gmv:sum(v,x=>x.gmv),act:sum(v,x=>x.actComm),sold:sum(v,x=>x.sold),refund:sum(v,x=>x.refund)}))
    .sort((a,b)=>b.gmv-a.gmv);
  const nShop=shopAgg.length, topShop=shopAgg[0]||{k:'—',gmv:0};
  const top4Share=A.gmv? sum(shopAgg.slice(0,4),s=>s.gmv)/A.gmv : 0;
  const top1Share=A.gmv? topShop.gmv/A.gmv : 0;
  const liveShare=A.gmv? sum(D.filter(r=>r.content==='Live'),r=>r.gmv)/A.gmv : 0;
  const prodAgg=[...groupBy(D,r=>r.product+'§'+r.productId).values()].map(v=>({
    name:v[0].product,gmv:sum(v,x=>x.gmv)}));
  const nProd=prodAgg.length;
  const topProd=[...prodAgg].sort((a,b)=>b.gmv-a.gmv)[0];
  const topProdShare=(A.gmv&&topProd)? topProd.gmv/A.gmv : 0;
  const ine=D.filter(r=>r.status==='Không đủ điều kiện');
  const ineRate=A.rows? ine.length/A.rows : 0;
  const ineLost=sum(ine,r=>r.estComm);
  const ineTopShop=[...groupBy(ine,r=>r.shop).entries()].map(([k,v])=>({k,lost:sum(v,x=>x.estComm)})).sort((a,b)=>b.lost-a.lost)[0];
  const pending=D.filter(r=>r.status==='Chờ xử lý'||r.status==='AwaitingPayment');
  const pendEst=sum(pending,r=>r.estComm);
  const refundShop=[...shopAgg].filter(s=>s.sold>=20).map(s=>({...s,rate:s.sold?s.refund/s.sold:0})).sort((a,b)=>b.rate-a.rate)[0];
  const eff=A.est? A.act/A.est : 0, gap=A.est-A.act;
  const settleRate=A.rows? A.settled/A.rows : 0;
  const refundRate=A.sold? A.refund/A.sold : 0;
  const p1=x=>(x*100).toFixed(1)+'%';

  // ===== Đoạn mở: nói đơn giản cho người không chuyên =====
  let intro=`Tổng tiền hàng bán ra là <b>${money(A.gmv)}</b> từ <b>${fmt(A.orders)}</b> đơn. `;
  if(gap>0)
    intro+=`Theo tỷ lệ hoa hồng, lẽ ra được nhận khoảng <b>${money(A.est)}</b>, nhưng đến giờ mới thực nhận <b>${money(A.act)}</b>. `+
      `Còn khoảng <b>${money(gap)}</b> (gần ${p1(1-eff)}) chưa về — một phần đơn còn đang chờ duyệt, một phần bị loại nên không được tính. `;
  else
    intro+=`Hoa hồng lẽ ra được nhận (<b>${money(A.est)}</b>) gần như đã về đủ tài khoản (<b>${money(A.act)}</b>). `;
  if(ineRate>=0.2 || eff<0.6) intro+=`Nói gọn: việc bán hàng đang tốt, vấn đề là làm sao <b>nhận đủ phần hoa hồng đáng được hưởng</b>.`;
  else if(top4Share>0.8 || liveShare>0.95) intro+=`Nhìn chung khá tốt, chỉ là nguồn thu đang dồn vào một vài chỗ, nên cần trải rộng hơn cho an toàn.`;
  else intro+=`Nhìn chung mọi thứ đang chạy ổn, vẫn còn vài điểm có thể cải thiện thêm.`;

  // ===== Bức tranh hiện tại (gạch đầu dòng, theo trình tự kể chuyện) =====
  const finds=[];
  const F=(sev,html)=>finds.push({sev,html});

  if(settleRate<0.7)
    F(settleRate<0.5?'bad':'warn',
      `<b>Hoa hồng về chậm.</b> Mới khoảng <b>${p1(settleRate)}</b> đơn hàng được duyệt trả hoa hồng, còn <b>${fmt(A.rows-A.settled)}</b> đơn đang bị treo hoặc bị loại. Đây là lý do chính khiến tiền thực nhận thấp hơn dự kiến.`);
  else
    F('good',`<b>Hoa hồng về đều.</b> Phần lớn đơn đã được duyệt trả hoa hồng (${p1(settleRate)}).`);

  if(ineRate>=0.1)
    F(ineRate>=0.2?'bad':'warn',
      `<b>Nhiều đơn bị loại — đây là chỗ mất tiền nhiều nhất.</b> Có tới <b>${p1(ineRate)}</b> đơn không được tính hoa hồng, tương đương khoảng <b>${money(ineLost)}</b> bị mất${ineTopShop?`, nhiều nhất ở cửa hàng <b>${esc(ineTopShop.k)}</b>`:''}. Thường là do khách huỷ/trả hàng, hoặc cách gắn link và nội dung chưa đúng quy định của TikTok.`);

  if(refundRate>=0.05)
    F(refundRate>=0.1?'bad':'warn',
      `<b>Tỷ lệ trả hàng hơi cao.</b> Khoảng <b>${(refundRate*100).toFixed(1)}%</b> số món bị trả lại (${fmt(A.refund)} món)${refundShop&&refundShop.rate>=0.1?`, cao nhất ở <b>${esc(refundShop.k)}</b> (${(refundShop.rate*100).toFixed(1)}%)`:''}. Mỗi món bị trả là một lần mất khoản hoa hồng đã tính.`);
  else if(A.sold>0)
    F('good',`<b>Trả hàng ít</b> (${(refundRate*100).toFixed(1)}%) — rất tốt.`);

  if(nShop>=2 && top4Share>0.6)
    F(top4Share>0.8?'bad':'warn',
      `<b>Doanh số phụ thuộc vào ít cửa hàng.</b> ${Math.min(4,nShop)} cửa hàng lớn nhất chiếm tới <b>${p1(top4Share)}</b> doanh số, riêng <b>${esc(topShop.k)}</b> đã <b>${p1(top1Share)}</b>. Nếu một trong số này bán chậm lại, thu nhập sẽ giảm theo ngay.`);

  if(liveShare>0.9)
    F('warn',`<b>Gần như chỉ bán được khi Live.</b> ${p1(liveShare)} đơn đến từ livestream; video ngắn và gian hàng trưng bày gần như chưa tạo ra đơn. Nghĩa là ngoài giờ live thì hầu như không có doanh thu.`);

  F(topProdShare>0.4?'warn':'good',
    topProdShare>0.4
      ? `<b>Phụ thuộc một sản phẩm chính.</b> Riêng “${esc(clip(topProd.name,40))}” đã chiếm ${p1(topProdShare)} doanh số trong tổng ${fmt(nProd)} sản phẩm.`
      : `<b>Sản phẩm khá đa dạng.</b> Có ${fmt(nProd)} sản phẩm và không quá phụ thuộc một món (món bán chạy nhất chỉ chiếm ${p1(topProdShare)} doanh số) — đây là điểm tốt.`);

  // ===== Nên làm gì tiếp theo =====
  const recos=[];
  const R=(p,html)=>recos.push({p,html});
  if(ineRate>=0.1 && ineLost>0)
    R(1,`<b>Xử lý nhóm đơn bị loại trước tiên.</b> Xem lại vì sao đơn bị loại (khách huỷ/trả hàng, hay gắn link sai cách)${ineTopShop?`, bắt đầu từ cửa hàng <b>${esc(ineTopShop.k)}</b>`:''}. Giảm được phần này là có thể lấy lại tới <b>${money(ineLost)}</b> hoa hồng.`);
  if(refundShop && refundShop.rate>=0.1)
    R(1,`<b>Trao đổi với cửa hàng ${esc(refundShop.k)} để giảm trả hàng.</b> Kiểm tra lại size, hình ảnh và mô tả sản phẩm cho đúng — đây là cách dễ nhất để bớt mất hoa hồng.`);
  if(top4Share>0.7 && nShop>4)
    R(2,`<b>Tìm thêm cửa hàng mới để hợp tác.</b> Đừng chỉ dựa vào vài cửa hàng lớn — trải rộng ra thì thu nhập sẽ ổn định hơn.`);
  if(liveShare>0.9)
    R(2,`<b>Làm thêm video ngắn và sắp xếp gian hàng trưng bày</b> để có đơn cả khi không livestream.`);
  if(pending.length>0 && pendEst>0)
    R(2,`<b>Theo dõi ${fmt(pending.length)} đơn đang chờ duyệt</b> (khoảng ${money(pendEst)} hoa hồng). Chờ qua thời gian khách có thể trả hàng, nếu đơn ổn thì khoản này sẽ về tài khoản.`);
  if(topProdShare>0.4 && topProd)
    R(2,`<b>Bán kèm thêm sản phẩm khác</b> cùng với món bán chạy “${esc(clip(topProd.name,38))}” để không phụ thuộc vào một sản phẩm.`);
  if(topShop.gmv>0)
    R(3,`<b>Học theo cách làm của ${esc(topShop.k)}</b> (đang hiệu quả nhất) — cùng khung giờ live, cách dẫn dắt, nhóm sản phẩm — rồi áp dụng cho các cửa hàng khác.`);
  R(3,`<b>Ưu tiên đẩy mạnh sản phẩm hoa hồng cao mà ít bị trả lại</b>; giảm bớt những sản phẩm hay bị trả hàng.`);

  // ===== Render 1 card prose =====
  let html=`<p>${intro}</p>`;
  html+=`<div class="sub-h"><span class="ico">${icon('activity')}</span>Tình hình hiện tại</div>`+
    `<ul class="bullets">`+finds.map(f=>`<li class="b-${f.sev}">${f.html}</li>`).join('')+`</ul>`;
  html+=`<div class="sub-h"><span class="ico">${icon('bulb')}</span>Nên làm gì tiếp theo</div>`+
    `<ul class="bullets">`+recos.sort((a,b)=>a.p-b.p).slice(0,6).map(r=>`<li>${r.html}</li>`).join('')+`</ul>`;
  document.getElementById('summaryProse').innerHTML=html;
}

/* ---- charts ---- */
const GRID='rgba(255,255,255,.07)', TICK='#a8a8b3';
Chart.defaults.color=TICK; Chart.defaults.font.family='"TikTok Sans",Segoe UI,system-ui,sans-serif';
function destroy(id){if(CHARTS[id]){CHARTS[id].destroy();delete CHARTS[id];}}
function mkChart(id,cfg){destroy(id);CHARTS[id]=new Chart(document.getElementById(id),cfg);}
const axisMoney={ticks:{color:TICK,callback:v=>moneyShort(v)},grid:{color:GRID}};
const axisCat={ticks:{color:TICK},grid:{display:false}};
const isNarrow=()=>window.innerWidth<=760;
const doughnutOpts=tot=>({responsive:true,maintainAspectRatio:false,cutout:'60%',
  plugins:{legend:{position:isNarrow()?'bottom':'right',labels:{boxWidth:12,padding:10,color:'#e8e8ee'}},
    tooltip:{callbacks:{label:c=>`${c.label}: ${money(c.parsed)} (${pct(c.parsed,tot)})`}}}});

function renderStatus(D){
  const rows=[...groupBy(D,r=>r.status).entries()].map(([k,v])=>({k,...agg(v)})).sort((a,b)=>b.gmv-a.gmv);
  const tot=sum(rows,r=>r.gmv);
  mkChart('chStatus',{type:'doughnut',data:{labels:rows.map(r=>r.k),
    datasets:[{data:rows.map(r=>r.gmv),backgroundColor:rows.map(r=>statusMeta(r.k).color),borderColor:'#121214',borderWidth:3}]},
    options:doughnutOpts(tot)});
  const t=agg(D);
  document.getElementById('tblStatus').innerHTML=
    `<thead><tr><th>Trạng thái</th><th>Số dòng</th><th>% dòng</th><th>GMV</th><th>HH ước tính</th><th>HH thực nhận</th><th>Tiền nhận cuối</th></tr></thead><tbody>`+
    rows.map(r=>{const m=statusMeta(r.k);return `<tr>
      <td><span class="pill ${m.cls}">${esc(r.k)}</span></td>
      <td class="num">${fmt(r.rows)}</td><td class="num">${pct(r.rows,t.rows)}</td>
      <td class="num">${money(r.gmv)}</td><td class="num">${money(r.est)}</td>
      <td class="num">${money(r.act)}</td><td class="num">${money(r.final)}</td></tr>`;}).join('')+`</tbody>`;
}
function renderContent(D){
  const rows=[...groupBy(D,r=>r.content).entries()].map(([k,v])=>({k,gmv:sum(v,x=>x.gmv)})).sort((a,b)=>b.gmv-a.gmv);
  const tot=sum(rows,r=>r.gmv);
  mkChart('chContent',{type:'doughnut',data:{labels:rows.map(r=>r.k),
    datasets:[{data:rows.map(r=>r.gmv),backgroundColor:rows.map((_,i)=>PALETTE[i%PALETTE.length]),borderColor:'#121214',borderWidth:3}]},
    options:doughnutOpts(tot)});
}
function renderDay(D){
  const wd=D.filter(r=>r.day), m=groupBy(wd,r=>r.day.key), keys=[...m.keys()].sort();
  mkChart('chDay',{data:{labels:keys.map(k=>m.get(k)[0].day.label),datasets:[
    {type:'bar',label:'GMV',data:keys.map(k=>sum(m.get(k),r=>r.gmv)),backgroundColor:rgba(C_CYAN,.38),borderColor:C_CYAN,borderWidth:1,yAxisID:'y',order:3},
    {type:'line',label:'HH ước tính',data:keys.map(k=>sum(m.get(k),r=>r.estComm)),borderColor:C_AMBER,backgroundColor:C_AMBER,tension:.3,pointRadius:2,yAxisID:'y1',order:1},
    {type:'line',label:'HH thực nhận',data:keys.map(k=>sum(m.get(k),r=>r.actComm)),borderColor:C_GREEN,backgroundColor:C_GREEN,tension:.3,pointRadius:2,yAxisID:'y1',order:0}
  ]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
    plugins:{legend:{labels:{boxWidth:12,color:'#e8e8ee'}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${money(c.parsed.y)}`}}},
    scales:{x:axisCat,
      y:{position:'left',title:{display:true,text:'GMV',color:TICK},ticks:{color:TICK,callback:v=>moneyShort(v)},grid:{color:GRID}},
      y1:{position:'right',title:{display:true,text:'Hoa hồng',color:TICK},ticks:{color:TICK,callback:v=>moneyShort(v)},grid:{drawOnChartArea:false}}}}});
}
function renderShops(D){
  const rows=[...groupBy(D,r=>r.shop).entries()].map(([k,v])=>({k,...agg(v)})).sort((a,b)=>b.gmv-a.gmv);
  const top=rows.slice(0,10);
  mkChart('chShop',{type:'bar',data:{labels:top.map(r=>r.k),datasets:[
    {label:'GMV',data:top.map(r=>r.gmv),backgroundColor:C_CYAN},
    {label:'HH thực nhận',data:top.map(r=>r.act),backgroundColor:C_RED}
  ]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
    plugins:{legend:{labels:{boxWidth:12,color:'#e8e8ee'}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${money(c.parsed.x)}`}}},
    scales:{x:axisMoney,y:axisCat}}});
  const maxG=Math.max(...rows.map(r=>r.gmv),1);
  document.getElementById('tblShop').innerHTML=
    `<thead><tr><th>Cửa hàng</th><th>Đơn</th><th>GMV</th><th>HH thực nhận</th><th>Hoàn</th><th>% hoàn</th></tr></thead><tbody>`+
    rows.map(r=>`<tr><td>${esc(r.k)}<div class="minibar"><span style="width:${(r.gmv/maxG*100).toFixed(1)}%;background:${C_CYAN}"></span></div></td>
      <td class="num">${fmt(r.orders)}</td><td class="num">${money(r.gmv)}</td>
      <td class="num">${money(r.act)}</td><td class="num">${fmt(r.refund)}</td>
      <td class="num ${r.sold&&r.refund/r.sold>0.08?'neg':''}">${pct(r.refund,r.sold)}</td></tr>`).join('')+`</tbody>`;
}
function renderProducts(D){
  const rows=[...groupBy(D,r=>r.product+'§'+r.productId).values()].map(v=>({
    name:v[0].product, sold:sum(v,r=>r.sold), refund:sum(v,r=>r.refund), gmv:sum(v,r=>r.gmv), act:sum(v,r=>r.actComm)}));
  const clip=s=>s.length>46?s.slice(0,46)+'…':s;
  const top=[...rows].sort((a,b)=>b.sold-a.sold).slice(0,15);
  document.getElementById('tblProdTop').innerHTML=
    `<thead><tr><th>Sản phẩm</th><th>Bán</th><th>GMV</th><th>HH thực nhận</th></tr></thead><tbody>`+
    top.map(r=>`<tr><td title="${esc(r.name)}">${esc(clip(r.name))}</td><td class="num">${fmt(r.sold)}</td>
      <td class="num">${money(r.gmv)}</td><td class="num">${money(r.act)}</td></tr>`).join('')+`</tbody>`;
  const ref=[...rows].filter(r=>r.refund>0).sort((a,b)=>b.refund-a.refund).slice(0,15);
  document.getElementById('tblProdRefund').innerHTML=
    `<thead><tr><th>Sản phẩm</th><th>Hoàn</th><th>Bán</th><th>% hoàn</th></tr></thead><tbody>`+
    (ref.length?ref.map(r=>`<tr><td title="${esc(r.name)}">${esc(clip(r.name))}</td><td class="num neg">${fmt(r.refund)}</td>
      <td class="num">${fmt(r.sold)}</td><td class="num">${pct(r.refund,r.sold)}</td></tr>`).join('')
      :`<tr><td colspan="4" style="text-align:center;color:var(--mut2)">Không có sản phẩm bị hoàn trong phạm vi lọc</td></tr>`)+`</tbody>`;
}
function renderIneligible(D){
  const ine=D.filter(r=>r.status==='Không đủ điều kiện');
  const rows=[...groupBy(ine,r=>r.shop).entries()].map(([k,v])=>({k,rows:v.length,gmv:sum(v,x=>x.gmv),lost:sum(v,x=>x.estComm)})).sort((a,b)=>b.lost-a.lost);
  const top=rows.slice(0,10);
  mkChart('chIne',{type:'bar',data:{labels:top.map(r=>r.k),datasets:[{label:'HH ước tính bị mất',data:top.map(r=>r.lost),backgroundColor:C_RED}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},
      tooltip:{callbacks:{label:c=>money(c.parsed.x)}}},scales:{x:axisMoney,y:axisCat}}});
  document.getElementById('tblIne').innerHTML=
    `<thead><tr><th>Cửa hàng</th><th>Số đơn</th><th>GMV</th><th>HH ước tính mất</th></tr></thead><tbody>`+
    (rows.length?rows.map(r=>`<tr><td>${esc(r.k)}</td><td class="num">${fmt(r.rows)}</td>
      <td class="num">${money(r.gmv)}</td><td class="num neg">${money(r.lost)}</td></tr>`).join('')
      :`<tr><td colspan="4" style="text-align:center;color:var(--mut2)">Không có đơn không đủ điều kiện</td></tr>`)+`</tbody>`;
  const totLost=sum(ine,r=>r.estComm), totGmv=sum(ine,r=>r.gmv), allEst=sum(D,r=>r.estComm);
  document.getElementById('ineInsight').innerHTML= (ine.length
    ? `<span class="ico">${icon('alert')}</span><div>Có <b>${fmt(ine.length)}</b> dòng “Không đủ điều kiện” (GMV ${money(totGmv)}), làm <b class="neg">mất ~${money(totLost)}</b> hoa hồng ước tính — tương đương <b>${pct(totLost,allEst)}</b> tổng HH ước tính. Nguyên nhân phổ biến: khách hủy/trả hàng, đơn không hợp lệ, hoặc vi phạm chính sách affiliate.</div>`
    : `<span class="ico">${icon('check')}</span><div>Không có đơn “Không đủ điều kiện” trong phạm vi đang lọc.</div>`);
}
function renderGap(D){
  const rows=[...groupBy(D,r=>r.shop).entries()].map(([k,v])=>{const est=sum(v,x=>x.estComm),act=sum(v,x=>x.actComm);
    return {k,est,act,gap:est-act,eff:est?act/est*100:0};}).sort((a,b)=>b.est-a.est);
  const top=rows.slice(0,10);
  mkChart('chGap',{type:'bar',data:{labels:top.map(r=>r.k),datasets:[
    {label:'HH ước tính',data:top.map(r=>r.est),backgroundColor:C_AMBER},
    {label:'HH thực nhận',data:top.map(r=>r.act),backgroundColor:C_GREEN}
  ]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
    plugins:{legend:{labels:{boxWidth:12,color:'#e8e8ee'}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${money(c.parsed.x)}`}}},
    scales:{x:axisMoney,y:axisCat}}});
  document.getElementById('tblGap').innerHTML=
    `<thead><tr><th>Cửa hàng</th><th>HH ước tính</th><th>HH thực nhận</th><th>Hao hụt</th><th>Hiệu suất</th></tr></thead><tbody>`+
    rows.map(r=>`<tr><td>${esc(r.k)}</td><td class="num">${money(r.est)}</td><td class="num">${money(r.act)}</td>
      <td class="num neg">${money(r.gap)}</td><td class="num ${r.eff<70?'warn':'pos'}">${r.eff.toFixed(1)}%</td></tr>`).join('')+`</tbody>`;
}

/* ============================================================
   Xuất CSV
   ============================================================ */
function exportCSV(){
  const D=getFiltered(), lines=[], A=agg(D);
  const push=(...a)=>lines.push(a.map(x=>{const s=String(x==null?'':x);return /[",;\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;}).join(','));
  push('BÁO CÁO PHÂN TÍCH AFFILIATE TIKTOK');
  push('File',META.fileName); push('Số dòng',A.rows); push('Số đơn',A.orders);
  push('Tổng GMV',Math.round(A.gmv)); push('HH ước tính',Math.round(A.est));
  push('HH thực nhận',Math.round(A.act)); push('Tiền nhận cuối',Math.round(A.final));
  push('Tỷ lệ quyết toán (%)',(A.rows?A.settled/A.rows*100:0).toFixed(1));
  push('Tỷ lệ hoàn (%)',(A.sold?A.refund/A.sold*100:0).toFixed(2)); push('');
  push('THEO TRẠNG THÁI','Số dòng','GMV','HH ước tính','HH thực nhận');
  for(const [k,v] of groupBy(D,r=>r.status)){const a=agg(v);push(k,a.rows,Math.round(a.gmv),Math.round(a.est),Math.round(a.act));}
  push('');
  push('THEO CỬA HÀNG','Đơn','GMV','HH thực nhận','Món hoàn');
  for(const [k,v] of [...groupBy(D,r=>r.shop)].sort((a,b)=>sum(b[1],x=>x.gmv)-sum(a[1],x=>x.gmv))){const a=agg(v);push(k,a.orders,Math.round(a.gmv),Math.round(a.act),a.refund);}
  push('');
  push('THEO NGÀY','GMV','HH ước tính','HH thực nhận');
  const dm=groupBy(D.filter(r=>r.day),r=>r.day.key);
  for(const k of [...dm.keys()].sort()){const v=dm.get(k);push(v[0].day.label,Math.round(sum(v,r=>r.gmv)),Math.round(sum(v,r=>r.estComm)),Math.round(sum(v,r=>r.actComm)));}
  const blob=new Blob(['﻿'+lines.join('\n')],{type:'text/csv;charset=utf-8'});
  const url=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=url; a.download='bao_cao_affiliate_'+(META.fileName||'tiktok').replace(/\.xlsx?$/i,'')+'.csv';
  a.click(); URL.revokeObjectURL(url);
}

/* ============================================================
   UI wiring
   ============================================================ */
function showErr(m){document.getElementById('errMsg').textContent=m;document.getElementById('err').style.display='flex';}
function hideErr(){document.getElementById('err').style.display='none';}
injectIcons();
const drop=document.getElementById('drop'), fileInput=document.getElementById('file');
drop.onclick=()=>fileInput.click();
fileInput.onchange=e=>{if(e.target.files[0])handleFile(e.target.files[0]);};
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('hover');}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('hover');}));
drop.addEventListener('drop',e=>{const f=e.dataTransfer.files[0];if(f)handleFile(f);});
// nạp từ link
const urlInput=document.getElementById('urlInput'), urlBtn=document.getElementById('urlBtn');
urlBtn.onclick=()=>handleUrl(urlInput.value);
urlInput.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();handleUrl(urlInput.value);}});
document.getElementById('btnSample').onclick=()=>{
  document.getElementById('dash').classList.add('hidden');
  document.getElementById('uploadArea').classList.remove('hidden');
  document.getElementById('btnSample').classList.add('hidden');
  document.getElementById('btnCsv').classList.add('hidden');
  fileInput.value=''; urlInput.value='';
};
document.getElementById('btnCsv').onclick=exportCSV;

/* Re-render khi đổi khổ màn vượt ngưỡng hẹp (để legend doughnut đổi vị trí) */
let _narrow=isNarrow(), _rt=null;
window.addEventListener('resize',()=>{
  if(!RAW.length) return;
  clearTimeout(_rt);
  _rt=setTimeout(()=>{ const n=isNarrow(); if(n!==_narrow){_narrow=n; render();} },200);
});
