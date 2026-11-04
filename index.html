<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Dashboard Report AWT</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{--bg:#f5f6f8;--card:#fff;--ink:#17202e;--mut:#667085;--line:#e3e6eb;--ok:#0f8f7e;--l10:#e0a100;--l30:#d92d20;--brand:#1d3a8a;--sel:#2f6fed;--wh:#7a5af8;box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#0e131b;--card:#161d28;--ink:#e8ecf2;--mut:#98a2b3;--line:#263041;--ok:#2ec4ae;--l10:#f2b94b;--l30:#f4655a;--brand:#8fb0ff;--sel:#6c9bff;--wh:#a894ff}}
:root[data-theme="dark"]{--bg:#0e131b;--card:#161d28;--ink:#e8ecf2;--mut:#98a2b3;--line:#263041;--ok:#2ec4ae;--l10:#f2b94b;--l30:#f4655a;--brand:#8fb0ff;--sel:#6c9bff;--wh:#a894ff}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 "Plus Jakarta Sans","Segoe UI",system-ui,sans-serif}
main{max-width:1080px;margin:0 auto;padding:16px}
header{display:flex;flex-wrap:wrap;gap:12px;justify-content:space-between;align-items:flex-start;margin-bottom:14px}
h1{font-size:22px;margin:0;color:var(--brand);letter-spacing:-.01em}
.sub{color:var(--mut);font-size:13px;margin-top:2px}
.btn{background:var(--brand);color:#fff;border:0;border-radius:8px;padding:9px 14px;font:600 13px inherit;font-family:inherit;cursor:pointer}
:root[data-theme="dark"] .btn{color:#0e131b}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]) .btn{color:#0e131b}}
.btn:focus-visible,select:focus-visible,input:focus-visible{outline:2px solid var(--sel);outline-offset:2px}
.filters{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:14px}
.filters label,.filters .grp{font-size:12px;color:var(--mut);display:flex;flex-direction:column;gap:3px}
.seg{display:flex;border:1px solid var(--line);border-radius:8px;overflow:hidden;background:var(--card)}
.seg button{border:0;background:transparent;color:var(--ink);font:600 13px inherit;font-family:inherit;padding:8px 14px;cursor:pointer;border-right:1px solid var(--line)}
.seg button:last-child{border-right:0}
.seg button.on{background:var(--sel);color:#fff}
.seg button:focus-visible{outline:2px solid var(--sel);outline-offset:-2px}
select,input[type=date]{background:var(--card);color:var(--ink);border:1px solid var(--line);border-radius:8px;padding:7px 9px;font:14px inherit;font-family:inherit}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:14px}
.hero{display:grid;grid-template-columns:minmax(150px,220px) 1fr;gap:20px;align-items:center}
.big{font-size:52px;font-weight:700;line-height:1;color:var(--ok);letter-spacing:-.03em}
.big small{font-size:22px}
.lbl{color:var(--mut);font-size:13px}
.bar{display:flex;height:26px;border-radius:6px;overflow:hidden;background:var(--line)}
.bar i{display:block;min-width:0}
.legend{display:flex;flex-wrap:wrap;gap:14px;margin-top:10px;font-size:13px}
.legend b{font-weight:700}
.dot{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:14px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.kpi .v{font-size:26px;font-weight:700;letter-spacing:-.02em}
.kpi .s{font-size:12px;color:var(--mut)}
h2{font-size:15px;margin:0 0 10px}
.two{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.scroll{overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:13px;white-space:nowrap}
th,td{padding:7px 9px;text-align:right;border-bottom:1px solid var(--line)}
th{color:var(--mut);font-weight:600;background:var(--card)}
th:first-child,td:first-child,th.l,td.l{text-align:left}
tfoot td{font-weight:700;border-top:2px solid var(--line)}
.pill{padding:2px 8px;border-radius:99px;font-size:12px;font-weight:600}
.good{background:rgba(15,143,126,.14);color:var(--ok)}.warn{background:rgba(224,161,0,.16);color:var(--l10)}.bad{background:rgba(217,45,32,.14);color:var(--l30)}
.empty{padding:30px;text-align:center;color:var(--mut)}
.dh{display:flex;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;gap:8px;margin-bottom:10px}.dh h2{margin:0}.dsel{font-size:12px;color:var(--mut);display:flex;flex-direction:column;gap:3px}
#msg{color:var(--l30);font-size:13px;margin-bottom:10px}
svg text{fill:var(--mut);font-size:11px;font-family:inherit}
@media(max-width:700px){.hero,.two{grid-template-columns:1fr}.big{font-size:44px}}
</style>
</head>
<body>
<main>
<header>
 <div><h1>Dashboard Report AWT</h1><div class="sub" id="meta"></div></div>
 <div><input type="file" id="file" accept=".xlsx,.xls" hidden><button class="btn" id="up">Unggah Report AWT baru</button></div>
</header>
<div id="msg"></div>
<div class="filters">
 <label>Dari tanggal<input type="date" id="f1"></label>
 <label>Sampai tanggal<input type="date" id="f2"></label>
 <label>Tipe NP<select id="ft"></select></label>
 <div class="grp">Jenis barang<div class="seg" id="fn"><button type="button" data-v="">Semua</button><button type="button" data-v="YES">Tinting</button><button type="button" data-v="NO">Non Tinting</button></div></div>
</div>
<div id="out"></div>
</main>
<script>
// [tanggal, tipe, tinting, totalNP, sellTotal, sellOntime, sellL10, sellL30, whTotal, whOntime, whL10, whL30]
var D=[["2026-11-01","Cash&Carry via Gudang Store","NO",20,1,1,0,0,19,19,0,0],["2026-11-01","Cash&Carry via Pickup & Delivery","NO",47,45,45,0,0,2,2,0,0],["2026-11-01","Cash&Carry via Pickup & Delivery","YES",72,72,70,0,2,0,0,0,0],["2026-11-02","Cash&Carry via Gudang Store","NO",17,0,0,0,0,17,16,1,0],["2026-11-02","Cash&Carry via Pickup & Delivery","NO",29,29,29,0,0,0,0,0,0],["2026-11-02","Cash&Carry via Pickup & Delivery","YES",69,69,68,0,1,0,0,0,0],["2026-11-03","Cash&Carry via Gudang Store","NO",18,3,3,0,0,15,14,1,0],["2026-11-03","Cash&Carry via Pickup & Delivery","NO",46,45,45,0,0,1,1,0,0],["2026-11-03","Cash&Carry via Pickup & Delivery","YES",69,69,68,0,1,0,0,0,0],["2026-11-04","Cash&Carry via Gudang Store","NO",17,3,3,0,0,14,14,0,0],["2026-11-04","Cash&Carry via Pickup & Delivery","NO",33,30,30,0,0,3,3,0,0],["2026-11-04","Cash&Carry via Pickup & Delivery","YES",46,46,46,0,0,0,0,0,0]];
var META={site:"10005 – MITRA10 PERCETAKAN NEGARA",rm:"LANGKAYANA",info:"Order 1–4 Nov 2026 · Reg 1"};
var $=function(i){return document.getElementById(i)};
var TINT="",DT_TYPE="";
function segUI(){var b=$("fn").querySelectorAll("button");for(var i=0;i<b.length;i++){b[i].className=b[i].getAttribute("data-v")===TINT?"on":""}}
var MON=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
function fd(s){var p=s.split("-");return +p[2]+" "+MON[+p[1]-1]}
function n(x){return x.toLocaleString("id-ID")}
function pc(a,b){return b?(a/b*100):0}
function pf(a,b){return pc(a,b).toFixed(1).replace(".",",")+"%"}
function short(t){return t.replace("Cash&Carry via ","")}
function agg(rs){var o={t:0,sT:0,sO:0,s10:0,s30:0,wT:0,wO:0,w10:0,w30:0};rs.forEach(function(r){o.t+=r[3];o.sT+=r[4];o.sO+=r[5];o.s10+=r[6];o.s30+=r[7];o.wT+=r[8];o.wO+=r[9];o.w10+=r[10];o.w30+=r[11]});o.on=o.sO+o.wO;o.l10=o.s10+o.w10;o.l30=o.s30+o.w30;return o}
function pill(v,t){var c=v>=98?"good":v>=95?"warn":"bad";return '<span class="pill '+c+'">'+(t)+'</span>'}
function setup(){
 var ds=D.map(function(r){return r[0]}).sort();
 $("f1").value=ds[0];$("f2").value=ds[ds.length-1];
 var ts=[];D.forEach(function(r){if(ts.indexOf(r[1])<0)ts.push(r[1])});
 $("ft").innerHTML='<option value="">Semua</option>'+ts.map(function(t){return '<option value="'+t+'">'+short(t)+'</option>'}).join("");
 TINT="";segUI();
 $("meta").textContent=META.site+" · RM "+META.rm+" · "+META.info;
 render();
}
function render(){
 var a=$("f1").value,b=$("f2").value,t=$("ft").value,ti=TINT;
 var rs=D.filter(function(r){return(!a||r[0]>=a)&&(!b||r[0]<=b)&&(!t||r[1]===t)&&(!ti||r[2]===ti)});
 if(!rs.length){$("out").innerHTML='<div class="card empty">Tidak ada data untuk filter ini. Ubah rentang tanggal atau tipe NP.</div>';return}
 var g=agg(rs),h="";
 var po=pc(g.on,g.t),p10=pc(g.l10,g.t),p30=pc(g.l30,g.t);
 h+='<div class="card hero"><div><div class="big">'+po.toFixed(1).replace(".",",")+'<small>%</small></div><div class="lbl">NP selesai tepat waktu dari '+n(g.t)+' NP</div></div><div>'
  +'<div class="bar" role="img" aria-label="Komposisi ketepatan waktu"><i style="width:'+po+'%;background:var(--ok)"></i><i style="width:'+p10+'%;background:var(--l10)"></i><i style="width:'+p30+'%;background:var(--l30)"></i></div>'
  +'<div class="legend"><span><span class="dot" style="background:var(--ok)"></span>Tepat waktu <b>'+n(g.on)+'</b> ('+pf(g.on,g.t)+')</span><span><span class="dot" style="background:var(--l10)"></span>Telat &gt;10 menit <b>'+n(g.l10)+'</b></span><span><span class="dot" style="background:var(--l30)"></span>Telat &gt;30 menit <b>'+n(g.l30)+'</b></span></div></div></div>';
 var sO=pf(g.sO,g.sT),wO=pf(g.wO,g.wT);
 h+='<div class="kpis">'
  +'<div class="kpi"><div class="s">Total NP</div><div class="v">'+n(g.t)+'</div><div class="s">'+rs.length+' baris data</div></div>'
  +'<div class="kpi"><div class="s">Selling Area</div><div class="v" style="color:var(--sel)">'+n(g.sT)+'</div><div class="s">'+pf(g.sT,g.t)+' dari total · tepat waktu '+sO+'</div></div>'
  +'<div class="kpi"><div class="s">WH Area</div><div class="v" style="color:var(--wh)">'+n(g.wT)+'</div><div class="s">'+pf(g.wT,g.t)+' dari total · tepat waktu '+wO+'</div></div>'
  +'<div class="kpi"><div class="s">Total telat</div><div class="v" style="color:var(--l30)">'+n(g.l10+g.l30)+'</div><div class="s">'+pf(g.l10+g.l30,g.t)+' dari total NP</div></div></div>';
 // tren harian
 var days=[];rs.forEach(function(r){if(days.indexOf(r[0])<0)days.push(r[0])});days.sort();
 var dd=days.map(function(d){return{d:d,g:agg(rs.filter(function(r){return r[0]===d}))}});
 var mx=Math.max.apply(null,dd.map(function(x){return x.g.t}))||1;
 var W=Math.max(520,dd.length*70),H=230,bw=Math.min(40,W/dd.length*.55),svg='<svg viewBox="0 0 '+W+' '+H+'" width="100%" style="min-width:'+Math.min(W,520)+'px" role="img" aria-label="Tren NP harian">';
 dd.forEach(function(x,i){
  var cx=(i+.5)*W/dd.length,x0=cx-bw/2,y=H-30,sc=(H-70)/mx;
  var segs=[[x.g.on,"var(--ok)"],[x.g.l10,"var(--l10)"],[x.g.l30,"var(--l30)"]];
  segs.forEach(function(s){var hh=s[0]*sc;if(hh>0){y-=hh;svg+='<rect x="'+x0+'" y="'+y+'" width="'+bw+'" height="'+hh+'" fill="'+s[1]+'" rx="2"/>'}});
  svg+='<text x="'+cx+'" y="'+(y-18)+'" text-anchor="middle" style="font-weight:700;fill:var(--ink)">'+x.g.t+'</text>';
  svg+='<text x="'+cx+'" y="'+(y-6)+'" text-anchor="middle">'+pc(x.g.on,x.g.t).toFixed(1).replace(".",",")+'%</text>';
  svg+='<text x="'+cx+'" y="'+(H-10)+'" text-anchor="middle">'+fd(x.d)+'</text>';
 });
 svg+='<line x1="0" x2="'+W+'" y1="'+(H-30)+'" y2="'+(H-30)+'" stroke="var(--line)"/></svg>';
 h+='<div class="card"><h2>Tren harian: jumlah NP dan % tepat waktu</h2><div class="scroll">'+svg+'</div></div>';
 // per area + per tipe
 function arow(nm,c,T,O,L1,L3){return '<tr><td class="l"><span class="dot" style="background:'+c+'"></span>'+nm+'</td><td>'+n(T)+'</td><td>'+n(O)+'</td><td>'+n(L1)+'</td><td>'+n(L3)+'</td><td>'+pill(pc(O,T),pf(O,T))+'</td></tr>'}
 h+='<div class="two"><div class="card"><h2>Perbandingan area picking</h2><div class="scroll"><table><thead><tr><th>Area</th><th>NP</th><th>Tepat</th><th>&gt;10 mnt</th><th>&gt;30 mnt</th><th>% Tepat</th></tr></thead><tbody>'
  +arow("Selling Area","var(--sel)",g.sT,g.sO,g.s10,g.s30)+arow("WH Area","var(--wh)",g.wT,g.wO,g.w10,g.w30)+'</tbody></table></div></div>';
 var keys=[];rs.forEach(function(r){var k=short(r[1])+"|"+r[2];if(keys.indexOf(k)<0)keys.push(k)});
 var tb='';keys.forEach(function(k){var p=k.split("|");var x=agg(rs.filter(function(r){return short(r[1])===p[0]&&r[2]===p[1]}));
  tb+='<tr><td class="l">'+p[0]+'</td><td class="l">'+(p[1]==="YES"?"Tinting":"Non Tinting")+'</td><td>'+n(x.t)+'</td><td>'+n(x.l10+x.l30)+'</td><td>'+pill(pc(x.on,x.t),pf(x.on,x.t))+'</td></tr>'});
 h+='<div class="card"><h2>Per tipe NP dan tinting</h2><div class="scroll"><table><thead><tr><th>Tipe NP</th><th class="l">Tinting</th><th>NP</th><th>Telat</th><th>% Tepat</th></tr></thead><tbody>'+tb+'</tbody></table></div></div></div>';
 if(!ti){var gt=agg(rs.filter(function(r){return r[2]==="YES"})),gn=agg(rs.filter(function(r){return r[2]==="NO"}));
  var crow=function(nm,c,x){return '<tr><td class="l"><span class="dot" style="background:'+c+'"></span>'+nm+'</td><td>'+n(x.t)+'</td><td>'+pf(x.t,g.t)+'</td><td>'+n(x.sT)+'</td><td>'+n(x.wT)+'</td><td>'+n(x.l10+x.l30)+'</td><td>'+(x.t?pill(pc(x.on,x.t),pf(x.on,x.t)):'-')+'</td></tr>'};
  h+='<div class="card"><h2>Tinting vs Non Tinting</h2><div class="scroll"><table><thead><tr><th>Jenis</th><th>NP</th><th>Porsi</th><th>Selling</th><th>WH</th><th>Telat</th><th>% Tepat</th></tr></thead><tbody>'+crow('Tinting','var(--wh)',gt)+crow('Non Tinting','var(--sel)',gn)+'</tbody></table></div></div>'}
 // detail
 var dts=[];rs.forEach(function(r){if(dts.indexOf(r[1])<0)dts.push(r[1])});var dtype=dts.indexOf(DT_TYPE)<0?"":DT_TYPE,drs=rs.filter(function(r){return!dtype||r[1]===dtype}),dg=agg(drs),dopts='<option value="">Semua</option>'+dts.map(function(t){return '<option value="'+t+'"'+(t===dtype?' selected':'')+'>'+short(t)+'</option>'}).join(""),dt='';drs.slice().sort(function(a,b){return a[0]<b[0]?-1:a[0]>b[0]?1:0}).forEach(function(r){var x=agg([r]);
  dt+='<tr><td class="l">'+fd(r[0])+'</td><td class="l">'+short(r[1])+'</td><td>'+(r[2]==="YES"?"Tinting":"Non Tinting")+'</td><td>'+r[3]+'</td><td>'+r[4]+'</td><td>'+r[8]+'</td><td>'+x.on+'</td><td>'+x.l10+'</td><td>'+x.l30+'</td><td>'+pill(pc(x.on,x.t),pf(x.on,x.t))+'</td></tr>'});
 h+='<div class="card"><div class="dh"><h2>Detail data</h2><label class="dsel">Tipe NP<select id="fdt">'+dopts+'</select></label></div><div class="scroll"><table><thead><tr><th>Tanggal</th><th class="l">Tipe NP</th><th>Tinting</th><th>Total</th><th>Selling</th><th>WH</th><th>Tepat</th><th>&gt;10</th><th>&gt;30</th><th>% Tepat</th></tr></thead><tbody>'+dt+'</tbody><tfoot><tr><td colspan="3" class="l">Total</td><td>'+n(dg.t)+'</td><td>'+n(dg.sT)+'</td><td>'+n(dg.wT)+'</td><td>'+n(dg.on)+'</td><td>'+n(dg.l10)+'</td><td>'+n(dg.l30)+'</td><td>'+pf(dg.on,dg.t)+'</td></tr></tfoot></table></div></div>';
 $("out").innerHTML=h;
}
["f1","f2","ft"].forEach(function(i){$(i).addEventListener("change",render)});
$("fn").addEventListener("click",function(e){var b=e.target.closest("button");if(!b)return;TINT=b.getAttribute("data-v");segUI();render()});
$("out").addEventListener("change",function(e){if(e.target.id==="fdt"){DT_TYPE=e.target.value;render()}});
$("up").addEventListener("click",function(){$("file").click()});
$("file").addEventListener("change",function(e){
 var f=e.target.files[0];if(!f)return;$("msg").textContent="";
 var rd=new FileReader();
 rd.onload=function(){try{
  var wb=XLSX.read(rd.result,{type:"array",cellDates:true});
  var rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:null});
  var hi=-1;rows.forEach(function(r,i){if(hi<0&&r.indexOf("Site ID")>=0)hi=i});
  if(hi<0)throw new Error("Kolom “Site ID” tidak ditemukan.");
  var nd=[],site="",rm="";
  for(var i=hi+1;i<rows.length;i++){var r=rows[i];if(!r[1]||!(r[4] instanceof Date))continue;
   var d=new Date(r[4].getTime()+12*3600*1000).toISOString().slice(0,10);
   site=r[1]+" – "+r[2];rm=r[3];
   nd.push([d,String(r[5]),String(r[7]),+r[8]||0,+r[9]||0,+r[11]||0,+r[13]||0,+r[15]||0,+r[17]||0,+r[19]||0,+r[21]||0,+r[23]||0]);
  }
  if(!nd.length)throw new Error("Tidak ada baris data pada file.");
  var info=String(rows[2]&&rows[2][1]||"").split("\n")[0];
  D=nd;META={site:site,rm:rm,info:info};setup();
 }catch(err){$("msg").textContent="Gagal membaca file: "+err.message+" Pastikan formatnya sama dengan Report AWT."}};
 rd.readAsArrayBuffer(f);e.target.value="";
});
setup();
</script>
</body>
</html>
