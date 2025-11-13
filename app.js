/* =========================
 * Point Sync - app.js
 * (giữ nguyên logic, thêm CD-Key: miễn phí 2 lần, từ lần 3 mới hỏi)
 * + Cải tiến làm tròn điểm kiểu ROUND_HALF_UP (2.25 -> 2.3)
 * + Đồng bộ cho TẤT CẢ sheet lớp trong file đích
 * ========================= */

/* ===== Helpers ===== */
const A2Z = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));
function setStatus(t){ const el=document.getElementById("status"); if(el) el.textContent=t; }
function addOptions(id, def){
  const s=document.getElementById(id); if(!s) return;
  s.innerHTML=""; A2Z.forEach(L=>{const o=document.createElement("option");o.value=L;o.textContent=L;s.appendChild(o);});
  s.value=def;
}
function clearMissing(){ const tb=document.getElementById("missing_tbody"); if(tb) tb.innerHTML=""; }
function addMissing(stt,hoten,lop){
  const tb=document.getElementById("missing_tbody"); if(!tb) return;
  const tr=document.createElement("tr");
  [stt,hoten,lop].forEach(v=>{const td=document.createElement("td"); td.textContent=v??""; tr.appendChild(td);});
  tb.appendChild(tr);
}
function colToIndex(L){ L=(L||"").toUpperCase(); let idx=0; for(const ch of L){const n=ch.charCodeAt(0)-64; if(n>=1&&n<=26) idx=idx*26+n;} return idx; }
function norm(s){ if(s==null) return ""; s=String(s).trim().toLowerCase().replaceAll("đ","d"); try{s=s.normalize("NFD").replace(/[\u0300-\u036f]/g,"");}catch(e){} return s; }
function last8(v){
  if(v==null) return "";
  let s=String(v).trim();
  if(/^[0-9]+\.0$/.test(s)) s=s.slice(0,-2);
  const d=(s.match(/[0-9]/g)||[]).join("");
  if(!d) return "";
  return d.padStart(8,"0").slice(-8);
}
function colLetter(n){ let s=""; while(n>0){const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26);} return s; }
const MAX_R=4000, MAX_C=128;

/* ===== Làm tròn kiểu ROUND_HALF_UP (2.25 -> 2.3) ===== */
function roundHalfUp(value, digits=0){
  if(value==null) return null;
  let s = String(value).trim().replace(",", ".");
  if(s==="") return null;

  // Chỉ xử lý số; nếu chuỗi lạ thì ép sang Number rồi stringify lại
  if(!/^-?\d+(\.\d+)?$/.test(s)){
    const num = Number(s);
    if(!isFinite(num)) return null;
    s = String(num);
  }

  let negative = false;
  if(s.startsWith("-")){
    negative = true;
    s = s.slice(1);
  }

  let [intPart, fracPart=""] = s.split(".");
  fracPart = fracPart.replace(/[^0-9]/g,"");

  if(digits <= 0){
    // làm tròn về số nguyên
    const first = fracPart[0] || "0";
    if(first >= "5"){
      let carry = 1, res = "";
      for(let i=intPart.length-1;i>=0;i--){
        let d = intPart.charCodeAt(i)-48 + carry;
        if(d>=10){ d-=10; carry=1; } else carry=0;
        res = String.fromCharCode(48+d) + res;
      }
      if(carry) res = "1" + res;
      intPart = res;
    }
    const outStr = (negative?"-":"") + intPart;
    return Number(outStr);
  }

  // digits > 0: làm tròn đến digits chữ số sau dấu phẩy
  while(fracPart.length < digits+1){
    fracPart += "0";
  }
  const cut   = fracPart.slice(0, digits);
  const next  = fracPart[digits] || "0";

  let fracRounded = cut;
  let carryInt = 0;

  if(next >= "5"){
    // +1 vào phần thập phân
    let carry = 1, res = "";
    for(let i=cut.length-1;i>=0;i--){
      let d = cut.charCodeAt(i)-48 + carry;
      if(d>=10){ d-=10; carry=1; } else carry=0;
      res = String.fromCharCode(48+d) + res;
    }
    fracRounded = res;
    if(carry){
      // tràn sang phần nguyên
      carryInt = 1;
    }
  }

  if(carryInt){
    let carry=1, res="";
    for(let i=intPart.length-1;i>=0;i--){
      let d = intPart.charCodeAt(i)-48 + carry;
      if(d>=10){ d-=10; carry=1; } else carry=0;
      res = String.fromCharCode(48+d) + res;
    }
    if(carry) res="1"+res;
    intPart = res;
  }

  fracRounded = fracRounded.padStart(digits,"0").slice(0,digits);
  let outStr = (negative?"-":"") + intPart;
  if(digits>0) outStr += "." + fracRounded;
  return Number(outStr);
}

/* Điểm: trả về float trong [0,10], làm sạch đến 2 chữ số bằng ROUND_HALF_UP */
function parseScore(val){
  if(val==null) return null;
  let s = String(val).trim();
  if(s==="") return null;
  s = s.replace(",", ".");

  const f = Number(s);
  if(!isFinite(f) || f < 0 || f > 10) return null;

  // làm sạch nguồn đến 2 chữ số (2.345 -> 2.35) theo half-up
  return roundHalfUp(f, 2);
}

/* ===== Detect rows/cols (giữ nguyên) ===== */
function detectRows(sheet){
  let header=null;
  for(let r=1;r<=50;r++){
    const a=sheet.cell(r,1).value();
    let found=false;
    for(let c=1;c<=MAX_C;c++){
      const v=sheet.cell(r,c).value();
      if(v==null||v==="") continue;
      if(norm(v).includes("ma dinh danh")){found=true;break;}
    }
    if(String(a).trim()==="STT"&&found){ header=r; break; }
  }
  if(!header) header=6;
  let first=null;
  for(let r=header+1;r<=header+200;r++){
    const a=sheet.cell(r,1).value();
    if(a!=null && /^\d+$/.test(String(a).trim())){ first=r; break; }
  }
  if(!first) first=header+3;
  let last=first-1, empty=0;
  for(let r=first;r<=Math.min(first+2000,MAX_R);r++){
    const a=sheet.cell(r,1).value();
    if(a==null||String(a).trim()===""){ empty++; if(empty>=3) break; }
    else { empty=0; last=r; }
  }
  return {header,first,last};
}

function autodetectCols(sheet, header){
  let maxCol=0, empties=0;
  for(let c=1;c<=MAX_C;c++){
    const v=sheet.cell(header,c).value();
    if(v==null||String(v).trim()===""){ empties++; if(empties>=12) break; }
    else { empties=0; maxCol=c; }
  }
  let txRoot=null;
  for(let c=1;c<=maxCol;c++){
    const s=norm(sheet.cell(header,c).value());
    if(s.includes("ddg")&&s.includes("tx")){ txRoot=c; break; }
  }
  let subR=null,best=-1;
  if(txRoot){
    for(let r=header+1;r<=header+4;r++){
      let cnt=0;
      for(let c=txRoot;c<=maxCol;c++){
        const v=sheet.cell(r,c).value(); if(v==null) break;
        const sv=String(v).trim();
        if(/^\d+$/.test(sv)) cnt++; else break;
      }
      if(cnt>best){ best=cnt; subR=cnt>0? r:null; }
    }
  }
  const mapping={}, txCols=[];
  if(txRoot&&subR){
    let c=txRoot, idx=1;
    while(c<=maxCol){
      const v=sheet.cell(subR,c).value(); if(v==null) break;
      const sv=String(v).trim();
      if(/^\d+$/.test(sv) && Number(sv)===idx){
        mapping[`Thuong xuyen ${idx}`]=c; txCols.push(c); idx++; c++;
      } else break;
    }
  }
  let gk=null, ck=null;
  for(let c=1;c<=maxCol;c++){
    const s=norm(sheet.cell(header,c).value());
    if(gk==null && ((s.includes("ddg")&&s.includes("gk"))||s.includes("giua"))) gk=c;
    if(ck==null && ((s.includes("ddg")&&s.includes("ck"))||s.includes("cuoi"))) ck=c;
  }
  if(gk) mapping["Giua ky"]=gk;
  if(ck) mapping["Cuoi ky"]=ck;

  if(Object.keys(mapping).length===0){
    mapping["Thuong xuyen 1"]=5;
    mapping["Thuong xuyen 2"]=6;
    mapping["Thuong xuyen 3"]=7;
    mapping["Giua ky"]=10;
    mapping["Cuoi ky"]=11;
  }

  const txLetters=txCols.map(colLetter);
  const gkL=gk?colLetter(gk):"-";
  const ckL=ck?colLetter(ck):"-";
  const txPart=txCols.length?`TX=${txCols.length} [${txLetters.join(", ")}]`:"TX=0";
  return {mapping, summary:`${txPart}; GK=${gkL}; CK=${ckL}`};
}

/* ===== State ===== */
let destArrayBuffer=null, srcArrayBuffer=null;
addOptions("col_sbd","A");
addOptions("col_score","B");

/* ===== CD-Key gate ===== */
const LS_KEY_RUNS="ps_runs";
const LS_KEY_LICENSED="ps_licensed";
const LS_KEY_DEVICE="ps_device";

async function sha256Hex(str){
  const enc=new TextEncoder().encode(str);
  const buf=await crypto.subtle.digest("SHA-256", enc);
  const arr=Array.from(new Uint8Array(buf));
  return arr.map(b=>b.toString(16).padStart(2,"0")).join("").toUpperCase();
}
/* Sinh/đọc deviceId ổn định */
async function getDeviceId(){
  let id=localStorage.getItem(LS_KEY_DEVICE);
  if(id) return id;
  const seed=[
    navigator.userAgent||"",
    navigator.platform||"",
    navigator.language||"",
    screen.width+"x"+screen.height+"@"+(window.devicePixelRatio||1),
  ].join("|");
  id=(await sha256Hex(seed)).slice(0,16).toUpperCase();
  localStorage.setItem(LS_KEY_DEVICE, id);
  return id;
}
/* CD-Key = 16 ký tự đầu của SHA-256(DeviceID), format 4-4-4-4 */
async function expectedKeyFromDeviceId(deviceId){
  const digest = await sha256Hex(deviceId);
  const first16 = digest.slice(0,16);
  return first16.match(/.{1,4}/g).join("-");
}
/* So khớp CD-Key người dùng nhập */
async function cdkeyMatches(input, deviceId){
  const raw = String(input||"").toUpperCase().replace(/[^A-Z0-9]/g,"");
  const expectedRaw = (await expectedKeyFromDeviceId(deviceId)).replace(/-/g,"").toUpperCase();
  return raw === expectedRaw;
}

async function isLicensedOrPrompt(){
  // đã kích hoạt?
  if(localStorage.getItem(LS_KEY_LICENSED)==="1") return true;

  // miễn phí 8 lần (bạn có thể chỉnh lại số lần nếu muốn)
  const runs = parseInt(localStorage.getItem(LS_KEY_RUNS)||"0",10);
  if(runs < 4){
    localStorage.setItem(LS_KEY_RUNS, String(runs+1));
    return true;
  }

  // hiện modal
  const deviceId = await getDeviceId();
  const modal = document.getElementById("cdkey_modal");
  const backdrop = document.getElementById("cdkey_backdrop");
  const inpDev = document.getElementById("cdkey_device");
  const inpKey = document.getElementById("cdkey_input");
  const errBox = document.getElementById("cdkey_error");
  const btnOk = document.getElementById("cdkey_submit");
  const btnCancel = document.getElementById("cdkey_cancel");

  inpDev.value = deviceId;
  inpKey.value = "";
  errBox.style.display="none";
  modal.hidden=false; backdrop.hidden=false;
  setTimeout(()=>inpKey.focus(), 50);

  function closeModal(){
    modal.hidden=true; backdrop.hidden=true;
    btnOk.removeEventListener("click", onSubmit);
    btnCancel.removeEventListener("click", onCancel);
    inpKey.removeEventListener("keydown", onEnter);
  }
  function onCancel(){ closeModal(); }
  async function onSubmit(){
    btnOk.disabled=true;
    const ok = await cdkeyMatches(inpKey.value, deviceId);
    btnOk.disabled=false;
    if(ok){
      localStorage.setItem(LS_KEY_LICENSED,"1");
      errBox.style.display="none";
      closeModal();
      setStatus("Kích hoạt thành công — bạn có thể tiếp tục đồng bộ điểm.");
      // cho chạy tiếp hành động đồng bộ
      document.getElementById("run_btn").click();
    }else{
      errBox.style.display="block";
    }
  }
  function onEnter(e){ if(e.key==="Enter"){ e.preventDefault(); onSubmit(); } }

  btnOk.addEventListener("click", onSubmit);
  btnCancel.addEventListener("click", onCancel);
  inpKey.addEventListener("keydown", onEnter);

  return false;
}

/* ===== File events ===== */
document.getElementById("dst_xlsx")?.addEventListener("change", async (e)=>{
  const f=e.target.files&&e.target.files[0];
  const nameEl=document.getElementById("dst_name");
  if(nameEl) nameEl.textContent=f?f.name:"";
  destArrayBuffer = f? await f.arrayBuffer(): null;
  const badge=document.getElementById("detect_info"); if(badge) badge.textContent="Chưa đọc file lớp.";
  const destSel=document.getElementById("dest_label"); if(destSel) destSel.innerHTML="";
});
document.getElementById("src_xlsx")?.addEventListener("change", async (e)=>{
  const f=e.target.files&&e.target.files[0];
  srcArrayBuffer = f? await f.arrayBuffer(): null;
});

/* ===== Analyze columns (lấy cấu trúc từ 1 sheet đại diện) ===== */
document.getElementById("analyze_btn")?.addEventListener("click", async ()=>{
  try{
    if(!destArrayBuffer){ setStatus("Chưa chọn file lớp (.xlsx)."); return; }
    setStatus("Đang phân tích cấu trúc file lớp...");

    const wb=await XlsxPopulate.fromDataAsync(destArrayBuffer);
    const skip=new Set(["huongdan","readme","guide","hdsd"]);
    let pick=null;
    wb.sheets().forEach(sh=>{
      if(pick) return;
      const nm=String(sh.name()).toLowerCase();
      if(skip.has(nm)) return;
      try{ if(String(sh.cell(6,1).value()).trim()==="STT") pick=sh.name(); }catch(e){}
    });
    if(!pick){
      wb.sheets().forEach(sh=>{ const nm=String(sh.name()).toLowerCase(); if(!pick && !skip.has(nm)) pick=sh.name(); });
    }
    if(!pick){
      document.getElementById("detect_info").textContent="Không tìm thấy sheet lớp phù hợp.";
      setStatus("Vui lòng kiểm tra lại mẫu file lớp."); return;
    }

    const ws=wb.sheet(pick);
    const {header}=detectRows(ws);
    const {mapping, summary}=autodetectCols(ws, header);

    const destSel=document.getElementById("dest_label");
    destSel.innerHTML="";
    Object.keys(mapping).forEach(k=>{ const o=document.createElement("option"); o.value=k; o.textContent=k; destSel.appendChild(o); });

    document.getElementById("detect_info").textContent=`[${pick}] ${summary}`;
    setStatus("Đã phân tích file lớp, chọn cột điểm đích để đồng bộ.");
  }catch(err){
    console.error(err);
    setStatus("Lỗi khi phân tích cột: "+(err&&err.message?err.message:String(err)));
  }
});

/* ===== Ctrl+R phím tắt ===== */
document.addEventListener("keydown",(ev)=>{
  if((ev.ctrlKey||ev.metaKey) && (ev.key==="r"||ev.key==="R")){
    ev.preventDefault();
    document.getElementById("run_btn").click();
  }
});

/* ===== RUN SYNC (có kiểm tra CD-Key, chạy cho TOÀN BỘ sheet lớp) ===== */
document.getElementById("run_btn").addEventListener("click", async ()=>{
  if(!(await isLicensedOrPrompt())) return;  /* chặn nếu chưa hợp lệ (modal sẽ hiện) */

  try{
    clearMissing();
    if(!srcArrayBuffer||!destArrayBuffer){ setStatus("Vui lòng chọn đủ 2 tệp .xlsx (nguồn & lớp)."); return; }

    const destLabel=document.getElementById("dest_label").value;
    if(!destLabel){ setStatus("Chưa có lựa chọn cột đích. Hãy bấm '3. Tìm kiếm cột cần chép' trước."); return; }

    setStatus("Đang đồng bộ điểm... Vui lòng chờ.");

    // ===== 1) File nguồn =====
    const swb=await XlsxPopulate.fromDataAsync(srcArrayBuffer);
    const sst=swb.sheet(0);
    const idIdx=colToIndex(document.getElementById("col_sbd").value||"A");
    const scIdx=colToIndex(document.getElementById("col_score").value||"B");

    let sMaxRow=2, empty=0;
    for(let r=2;r<=MAX_R;r++){
      const v1=sst.cell(r,idIdx).value();
      const v2=sst.cell(r,scIdx).value();
      if((v1==null||String(v1).trim()==="")&&(v2==null||String(v2).trim()==="")){
        empty++; if(empty>=50){ sMaxRow=r-empty; break; }
      } else { empty=0; sMaxRow=r; }
    }

    const sourceMap={}; let invalidScores=0;
    for(let r=2;r<=sMaxRow;r++){
      const sid=sst.cell(r,idIdx).value();
      const sc=sst.cell(r,scIdx).value();
      const k=last8(sid); if(!k) continue;
      const score=parseScore(sc);
      if(score==null){ invalidScores++; continue; }
      sourceMap[k]=score;
    }

    // ===== 2) File lớp: duyệt TOÀN BỘ sheet lớp =====
    const dwb=await XlsxPopulate.fromDataAsync(destArrayBuffer);
    const names=dwb.sheets().map(sh=>sh.name());
    const skip=new Set(["huongdan","readme","guide","hdsd"]);
    const classSheets=names.filter(n=>!skip.has(String(n).toLowerCase()));

    let totalStudentsAll=0, copiedAll=0, missingAll=0, updatedSheets=0;

    for(const sname of classSheets){
      const ws=dwb.sheet(sname);
      const {header,first,last}=detectRows(ws);
      const {mapping}=autodetectCols(ws, header);
      if(!(destLabel in mapping)) continue;
      const destCol=mapping[destLabel];

      let maxCol=0, empt=0;
      for(let c=1;c<=MAX_C;c++){
        const v=ws.cell(header,c).value();
        if(v==null||String(v).trim()===""){ empt++; if(empt>=12) break; } else { empt=0; maxCol=c; }
      }

      // cột Mã định danh
      let mddCol=null;
      for(let c=1;c<=maxCol;c++){
        const s=norm(ws.cell(header,c).value());
        let scr=0;
        if(s.includes("ma dinh danh")) scr+=1;
        if(s.includes("bo gd")||s.includes("bo gd&dt")||s.includes("bo gd dt")) scr+=2;
        if(s.includes("ma dd")||s.includes("mdd")) scr+=1;
        if(scr>0){ mddCol=c; break; }
      }
      if(!mddCol) continue;

      // Họ tên + Lớp
      let nameCol=null, classCol=null;
      for(let c=1;c<=maxCol;c++){
        const s=norm(ws.cell(header,c).value());
        if(!nameCol && (s.includes("ho va ten")||s.includes("hovaten")||s.includes("ho ten")||(s.includes("ho")&&s.includes("ten")))) nameCol=c;
        if(!classCol && s.includes("lop")) classCol=c;
        if(nameCol && classCol) break;
      }

      const presentRows=[];
      for(let r=first;r<=last;r++){
        const a=ws.cell(r,1).value();
        if(!(a==null || String(a).trim()==="")) presentRows.push(r);
      }
      if(presentRows.length===0) continue;

      // Ghi điểm cho từng HS trong sheet
      for(const r of presentRows){
        const mddVal=ws.cell(r,mddCol).value();
        const k=last8(mddVal);
        if(k && (k in sourceMap)){
          // LÀM TRÒN 1 CHỮ SỐ BẰNG ROUND_HALF_UP (2.25 -> 2.3)
          const val = roundHalfUp(sourceMap[k], 1);
          ws.cell(r,destCol).value(val).style("numberFormat","0.0");
        }
      }

      // tô vàng những ô vẫn chưa có điểm + gom danh sách thiếu
      const yellow="ffff00";
      let copiedSheet=0;
      const missingRows=[];
      for(const r of presentRows){
        const v=ws.cell(r,destCol).value();
        if(v==null||String(v)===""){ missingRows.push(r); ws.cell(r,destCol).style("fill",yellow); }
        else { copiedSheet++; }
      }
      for(const r of missingRows){
        const stt=ws.cell(r,1).value();
        const hoten=nameCol? ws.cell(r,nameCol).value() : "";
        const lop=classCol? ws.cell(r,classCol).value() : sname;
        addMissing(String(stt??"").trim(), String(hoten??"").trim(), String(lop??"").trim());
      }

      totalStudentsAll+=presentRows.length;
      copiedAll+=copiedSheet;
      missingAll+=missingRows.length;
      updatedSheets++;
    }

    // ===== 3) Xuất file kết quả =====
    const outName=`Dongbo_cot_${destLabel}_toan_bo_lop.xlsx`;
    const blob=await dwb.outputAsync();
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a"); a.href=url; a.download=outName; a.click();
    URL.revokeObjectURL(url);

    const extra = invalidScores ? ` (Bỏ qua ${invalidScores} giá trị điểm không hợp lệ ở file nguồn.)` : "";
    setStatus(`Hoàn tất: ${copiedAll}/${totalStudentsAll} HS · Lớp cập nhật: ${updatedSheets} · Thiếu điểm: ${missingAll} · File xuất: ${outName}${extra}`);
  }catch(err){
    console.error(err);
    setStatus("Lỗi khi đồng bộ: "+(err && err.message ? err.message : String(err)));
  }
});

/* Re-ensure selects (an toàn) */
addOptions("col_sbd","A");
addOptions("col_score","B");

/* Dành cho tác giả: xem nhanh deviceId & CD-Key đúng
   Mở DevTools gõ: await window.PointSync.authorInfo()  */
window.PointSync = {
  async authorInfo(){
    const id = await getDeviceId();
    const key = await expectedKeyFromDeviceId(id);
    console.log("Device ID:", id, "CD-Key hợp lệ:", key);
    return { deviceId:id, cdkey:key };
  }
};
