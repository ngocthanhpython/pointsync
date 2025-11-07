// ===== Helpers =====
const A2Z = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));

function setStatus(t) {
  document.getElementById("status").textContent = t;
}

function addOptions(id, def) {
  const s = document.getElementById(id);
  s.innerHTML = "";
  A2Z.forEach((L) => {
    const o = document.createElement("option");
    o.value = L;
    o.textContent = L;
    s.appendChild(o);
  });
  s.value = def;
}

function clearMissing() {
  document.getElementById("missing_tbody").innerHTML = "";
}

function addMissing(stt, hoten, lop) {
  const tr = document.createElement("tr");
  [stt, hoten, lop].forEach((v) => {
    const td = document.createElement("td");
    td.textContent = v ?? "";
    tr.appendChild(td);
  });
  document.getElementById("missing_tbody").appendChild(tr);
}

function colToIndex(L) {
  L = (L || "").toUpperCase();
  let idx = 0;
  for (const ch of L) {
    const n = ch.charCodeAt(0) - 64;
    if (n >= 1 && n <= 26) idx = idx * 26 + n;
  }
  return idx;
}

function norm(s) {
  if (s == null) return "";
  s = String(s).trim().toLowerCase().replaceAll("đ", "d");
  try {
    s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  } catch (e) {}
  return s;
}

function last8(v) {
  if (v == null) return "";
  let s = String(v).trim();
  if (/^[0-9]+\.0$/.test(s)) s = s.slice(0, -2);
  const d = (s.match(/[0-9]/g) || []).join("");
  if (!d) return "";
  return d.padStart(8, "0").slice(-8);
}

function parseScore(val) {
  if (val == null) return null;
  const f = Number(String(val).trim().replace(",", "."));
  if (!isFinite(f) || f < 0 || f > 10) return null;
  return Math.round(f * 100) / 100;
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

const MAX_R = 4000;
const MAX_C = 128;

function detectRows(sheet) {
  let header = null;
  for (let r = 1; r <= 50; r++) {
    const a = sheet.cell(r, 1).value();
    let found = false;
    for (let c = 1; c <= MAX_C; c++) {
      const v = sheet.cell(r, c).value();
      if (v == null || v === "") continue;
      if (norm(v).includes("ma dinh danh")) {
        found = true;
        break;
      }
    }
    if (String(a).trim() === "STT" && found) {
      header = r;
      break;
    }
  }
  if (!header) header = 6;

  let first = null;
  for (let r = header + 1; r <= header + 200; r++) {
    const a = sheet.cell(r, 1).value();
    if (a != null && /^\d+$/.test(String(a).trim())) {
      first = r;
      break;
    }
  }
  if (!first) first = header + 3;

  let last = first - 1;
  let empty = 0;
  for (let r = first; r <= Math.min(first + 2000, MAX_R); r++) {
    const a = sheet.cell(r, 1).value();
    if (a == null || String(a).trim() === "") {
      empty++;
      if (empty >= 3) break;
    } else {
      empty = 0;
      last = r;
    }
  }
  return { header, first, last };
}

function autodetectCols(sheet, header) {
  let maxCol = 0, empties = 0;
  for (let c = 1; c <= MAX_C; c++) {
    const v = sheet.cell(header, c).value();
    if (v == null || String(v).trim() === "") {
      empties++;
      if (empties >= 12) break;
    } else {
      empties = 0;
      maxCol = c;
    }
  }

  let txRoot = null;
  for (let c = 1; c <= maxCol; c++) {
    const s = norm(sheet.cell(header, c).value());
    if (s.includes("ddg") && s.includes("tx")) {
      txRoot = c;
      break;
    }
  }

  let subR = null, best = -1;
  if (txRoot) {
    for (let r = header + 1; r <= header + 4; r++) {
      let cnt = 0;
      for (let c = txRoot; c <= maxCol; c++) {
        const v = sheet.cell(r, c).value();
        if (v == null) break;
        const sv = String(v).trim();
        if (/^\d+$/.test(sv)) cnt++;
        else break;
      }
      if (cnt > best) {
        best = cnt;
        subR = cnt > 0 ? r : null;
      }
    }
  }

  const mapping = {};
  const txCols = [];

  if (txRoot && subR) {
    let c = txRoot, idx = 1;
    while (c <= maxCol) {
      const v = sheet.cell(subR, c).value();
      if (v == null) break;
      const sv = String(v).trim();
      if (/^\d+$/.test(sv) && Number(sv) === idx) {
        mapping[`Thuong xuyen ${idx}`] = c;
        txCols.push(c);
        idx++; c++;
      } else break;
    }
  }

  let gk = null, ck = null;
  for (let c = 1; c <= maxCol; c++) {
    const s = norm(sheet.cell(header, c).value());
    if (gk == null && ((s.includes("ddg") && s.includes("gk")) || s.includes("giua"))) gk = c;
    if (ck == null && ((s.includes("ddg") && s.includes("ck")) || s.includes("cuoi"))) ck = c;
  }

  if (gk) mapping["Giua ky"] = gk;
  if (ck) mapping["Cuoi ky"] = ck;

  if (Object.keys(mapping).length === 0) {
    mapping["Thuong xuyen 1"] = 5;
    mapping["Thuong xuyen 2"] = 6;
    mapping["Thuong xuyen 3"] = 7;
    mapping["Giua ky"] = 10;
    mapping["Cuoi ky"] = 11;
  }

  const txLetters = txCols.map(colLetter);
  const gkL = gk ? colLetter(gk) : "-";
  const ckL = ck ? colLetter(ck) : "-";
  const txPart = txCols.length ? `TX=${txCols.length} [${txLetters.join(", ")}]` : "TX=0";

  return {
    mapping,
    summary: `${txPart}; GK=${gkL}; CK=${ckL}`,
  };
}

// ===== State =====
let destArrayBuffer = null;
let srcArrayBuffer = null;

// Init selects (sau khi DOM sẵn sàng)
document.addEventListener('DOMContentLoaded', () => {
  addOptions("col_sbd", "A");
  addOptions("col_score", "B");
});

// File events
document.getElementById("dst_xlsx").addEventListener("change", async (e) => {
  const f = e.target.files && e.target.files[0];
  document.getElementById("dst_name").textContent = f ? f.name : "";
  destArrayBuffer = f ? await f.arrayBuffer() : null;
  document.getElementById("detect_info").textContent = "Chưa đọc file lớp.";
  document.getElementById("dest_label").innerHTML = "";
});

document.getElementById("src_xlsx").addEventListener("change", async (e) => {
  const f = e.target.files && e.target.files[0];
  srcArrayBuffer = f ? await f.arrayBuffer() : null;
});

// Analyze columns
document.getElementById("analyze_btn").addEventListener("click", async () => {
  try {
    if (!destArrayBuffer) {
      setStatus("Chưa chọn file lớp (.xlsx).");
      return;
    }
    setStatus("Đang phân tích cấu trúc file lớp...");

    const wb = await XlsxPopulate.fromDataAsync(destArrayBuffer);
    const skip = new Set(["huongdan", "readme", "guide", "hdsd"]);

    let pick = null;
    wb.sheets().forEach((sh) => {
      if (pick) return;
      const nm = String(sh.name()).toLowerCase();
      if (skip.has(nm)) return;
      try {
        if (String(sh.cell(6, 1).value()).trim() === "STT") pick = sh.name();
      } catch (e) {}
    });
    if (!pick) {
      wb.sheets().forEach((sh) => {
        const nm = String(sh.name()).toLowerCase();
        if (!pick && !skip.has(nm)) pick = sh.name();
      });
    }

    if (!pick) {
      document.getElementById("detect_info").textContent = "Không tìm thấy sheet lớp phù hợp.";
      setStatus("Vui lòng kiểm tra lại mẫu file lớp.");
      return;
    }

    const ws = wb.sheet(pick);
    const { header } = detectRows(ws);
    const { mapping, summary } = autodetectCols(ws, header);

    const destSel = document.getElementById("dest_label");
    destSel.innerHTML = "";
    Object.keys(mapping).forEach((k) => {
      const o = document.createElement("option");
      o.value = k;
      o.textContent = k;
      destSel.appendChild(o);
    });

    document.getElementById("detect_info").textContent = `[${pick}] ${summary}`;
    setStatus("Đã phân tích file lớp, chọn cột điểm đích để đồng bộ.");
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi phân tích cột: " + (err && err.message ? err.message : String(err)));
  }
});

// Ctrl+R
document.addEventListener("keydown", (ev) => {
  if ((ev.ctrlKey || ev.metaKey) && (ev.key === "r" || ev.key === "R")) {
    ev.preventDefault();
    document.getElementById("run_btn").click();
  }
});

// Run sync
document.getElementById("run_btn").addEventListener("click", async () => {
  try {
    clearMissing();

    if (!srcArrayBuffer || !destArrayBuffer) {
      setStatus("Vui lòng chọn đủ 2 tệp .xlsx (nguồn & lớp).");
      return;
    }

    const destLabel = document.getElementById("dest_label").value;
    if (!destLabel) {
      setStatus("Chưa có lựa chọn cột đích. Hãy bấm 'Phân tích cột' trước.");
      return;
    }

    setStatus("Đang đồng bộ điểm... Vui lòng chờ.");

    // Source
    const swb = await XlsxPopulate.fromDataAsync(srcArrayBuffer);
    const sst = swb.sheet(0);
    const idIdx = colToIndex(document.getElementById("col_sbd").value || "A");
    const scIdx = colToIndex(document.getElementById("col_score").value || "B");

    let sMaxRow = 2, empty = 0;
    for (let r = 2; r <= MAX_R; r++) {
      const v1 = sst.cell(r, idIdx).value();
      const v2 = sst.cell(r, scIdx).value();
      if ((v1 == null || String(v1).trim() === "") && (v2 == null || String(v2).trim() === "")) {
        empty++;
        if (empty >= 50) { sMaxRow = r - empty; break; }
      } else {
        empty = 0; sMaxRow = r;
      }
    }

    const sourceMap = {};
    let invalidScores = 0;
    for (let r = 2; r <= sMaxRow; r++) {
      const sid = sst.cell(r, idIdx).value();
      const sc = sst.cell(r, scIdx).value();
      const k = last8(sid);
      if (!k) continue;
      const score = parseScore(sc);
      if (score == null) { invalidScores++; continue; }
      sourceMap[k] = score;
    }

    // Dest
    const dwb = await XlsxPopulate.fromDataAsync(destArrayBuffer);
    const names = dwb.sheets().map((sh) => sh.name());
    const skip = new Set(["huongdan", "readme", "guide", "hdsd"]);
    const classSheets = names.filter((n) => !skip.has(String(n).toLowerCase()));

    let totalStudentsAll = 0;
    let copiedAll = 0;
    let missingAll = 0;
    let updatedSheets = 0;

    for (const sname of classSheets) {
      const ws = dwb.sheet(sname);
      const { header, first, last } = detectRows(ws);
      const { mapping } = autodetectCols(ws, header);
      if (!(destLabel in mapping)) continue;
      const destCol = mapping[destLabel];

      // maxCol
      let maxCol = 0, empt = 0;
      for (let c = 1; c <= MAX_C; c++) {
        const v = ws.cell(header, c).value();
        if (v == null || String(v).trim() === "") {
          empt++;
          if (empt >= 12) break;
        } else {
          empt = 0; maxCol = c;
        }
      }

      // Mã định danh
      let mddCol = null;
      for (let c = 1; c <= maxCol; c++) {
        const s = norm(ws.cell(header, c).value());
        let scr = 0;
        if (s.includes("ma dinh danh")) scr += 1;
        if (s.includes("bo gd") || s.includes("bo gd&dt") || s.includes("bo gd dt")) scr += 2;
        if (s.includes("ma dd") || s.includes("mdd")) scr += 1;
        if (scr > 0) { mddCol = c; break; }
      }
      if (!mddCol) continue;

      // Họ tên + Lớp
      let nameCol = null, classCol = null;
      for (let c = 1; c <= maxCol; c++) {
        const s = norm(ws.cell(header, c).value());
        if (!nameCol && (s.includes("ho va ten") || s.includes("hovaten") || s.includes("ho ten") || (s.includes("ho") && s.includes("ten")))) nameCol = c;
        if (!classCol && s.includes("lop")) classCol = c;
        if (nameCol && classCol) break;
      }

      const presentRows = [];
      for (let r = first; r <= last; r++) {
        const a = ws.cell(r, 1).value();
        if (!(a == null || String(a).trim() === "")) presentRows.push(r);
      }
      if (presentRows.length === 0) continue;

      // Ghi điểm
      for (const r of presentRows) {
        const mddVal = ws.cell(r, mddCol).value();
        const k = last8(mddVal);
        if (k && k in sourceMap) {
          const val = Math.round(Number(sourceMap[k]) * 10) / 10;
          ws.cell(r, destCol).value(val).style("numberFormat", "0.0");
        }
      }

      const yellow = "ffff00";
      let copiedSheet = 0;
      const missingRows = [];

      for (const r of presentRows) {
        const v = ws.cell(r, destCol).value();
        if (v == null || String(v) === "") {
          missingRows.push(r);
          ws.cell(r, destCol).style("fill", yellow);
        } else {
          copiedSheet++;
        }
      }

      for (const r of missingRows) {
        const stt = ws.cell(r, 1).value();
        const hoten = nameCol ? ws.cell(r, nameCol).value() : "";
        const lop = classCol ? ws.cell(r, classCol).value() : sname;
        addMissing(String(stt ?? "").trim(), String(hoten ?? "").trim(), String(lop ?? "").trim());
      }

      totalStudentsAll += presentRows.length;
      copiedAll += copiedSheet;
      missingAll += missingRows.length;
      updatedSheets++;
    }

    // Xuất file
    const outName = `Dongbo_cot_${destLabel}_toan_bo_lop.xlsx`;
    const blob = await dwb.outputAsync();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = outName; a.click();
    URL.revokeObjectURL(url);

    const extra = invalidScores ? ` (Bỏ qua ${invalidScores} giá trị điểm không hợp lệ ở file nguồn.)` : "";
    setStatus(`Hoàn tất: ${copiedAll}/${totalStudentsAll} HS · Lớp cập nhật: ${updatedSheets} · Thiếu điểm: ${missingAll} · File xuất: ${outName}${extra}`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi đồng bộ: " + (err && err.message ? err.message : String(err)));
  }
});
