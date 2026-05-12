/* ── CLOCK ── */
const DAYS = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
const MONS = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "Mei",
  "Jun",
  "Jul",
  "Agu",
  "Sep",
  "Okt",
  "Nov",
  "Des",
];

function tick() {
  const n = new Date();
  const hms = [n.getHours(), n.getMinutes(), n.getSeconds()]
    .map((x) => String(x).padStart(2, "0"))
    .join(":");
  document.getElementById("clockTime").textContent = hms;
  document.getElementById("clockDate").textContent = `${
    DAYS[n.getDay()]
  }, ${n.getDate()} ${MONS[n.getMonth()]} ${n.getFullYear()}`;
}
tick();
setInterval(tick, 1000);

/* ── COUNT ANIMATION ── */
function animCount(el, to, dur = 650) {
  const s = performance.now(),
    from = parseInt(el.textContent) || 0;
  (function step(now) {
    const p = Math.min((now - s) / dur, 1),
      e = 1 - Math.pow(1 - p, 3);
    el.textContent = Math.round(from + (to - from) * e);
    if (p < 1) requestAnimationFrame(step);
  })(s);
}

const DEPT_MAP = [
  // ── ADMINITRASI group ──────────────────────────────────────
  {
    patterns: [
      "mekanik cmy scf",
      "mekonik cmy scf",
      "mekanik cmyscf",
      "mekanik cmy",
    ],
    sheetKeyword: "Mekanik CMY SCF", // D6
    displayName: "Mekanik CMY SCF 工務",
  },
  {
    patterns: ["mekanik", "kabag mekanik", "adm mekanik"],
    sheetKeyword: "Mekanik 工務", // D5
    displayName: "Mekanik 工務",
  },
  {
    patterns: [
      "kabag hr-ga",
      "kabag hr ga",
      "kabag hr_ga",

      "it",
      "office boy",
      "ob",
      "accounting",
      "finance",
      "general affair",
      "ga",
      "exim",
      "scrap",
      "hrd",
      "driver",
      "purchasing umum",
      "purchasing administrasi",
      "purchasing adm",

      "kabag exim",
    ],
    sheetKeyword: "Administrasi", // D4
    displayName: "Administrasi 管理部",
  },

  // ── MPQ group ──────────────────────────────────────────────
  {
    patterns: ["ppic delivery", "ppic-delivery"],
    sheetKeyword: "PPIC Delivery", // D9
    displayName: "PPIC Delivery 生管 出貨",
  },
  {
    patterns: ["ppic planning", "ppic plannnig", "ppic plan", "ppic"],
    sheetKeyword: "PPIC Planning", // D8
    displayName: "PPIC Planning 生管 文件與",
  },
  {
    patterns: ["marketing"],
    sheetKeyword: "Marketing", // D7
    displayName: "Marketing 業務",
  },
  {
    patterns: ["purchasing"],
    sheetKeyword: "Purchasing", // D10
    displayName: "Purchasing 採購",
  },
  {
    patterns: [
      "warehouse",
      "gudang",
      "wh",
      "kepala wh",
      "kepala warehouse",
      "kabag wh",
      "kabag warehouse",
      "adm wh",
      "adm warehouse",
      "admin warehouse",
    ],
    sheetKeyword: "Warehouse", // D11
    displayName: "Warehouse 倉庫",
  },
  {
    patterns: [
      "qc",
      "quality control",

      "qc finish good",
      "assistant leader qc",
      "qc belah kecil",
      "qc cutting",
      "qc printing",
      "qc incoming",
      "qc molded",
      "qc laminating",
      "qc laboratorium",
      "qc belah besar",
      "qc check claim",
      "adm qc",
    ],
    sheetKeyword: "QC", // D12
    displayName: "QC 品管",
  },

  // ── Production ─────────────────────────────────────────────
  {
    patterns: [
      "laminating",
      "lamintaing",
      "kepala shift laminating",
      "kepala shift lamintaing",
    ],
    sheetKeyword: "Laminating", // D13
    displayName: "Laminating 貼合",
  },
  {
    patterns: [
      "printing",
      "printong",
      "leader line printing",
      "leader line printong",
      "adm printing",
      "adm printong",
    ],
    sheetKeyword: "Printing", // D14
    displayName: "Printing 印刷",
  },
  {
    patterns: ["mixing", "mixong", "leader mixing", "leader mixong"],
    sheetKeyword: "Mixing", // D15
    displayName: "Mixing 配色房",
  },
  {
    patterns: ["belah kecil"],
    sheetKeyword: "Belah Kecil", // D16
    displayName: "Belah Kecil 小剖台",
  },
  {
    patterns: ["cutting molded", "leader line cutting molded"],
    sheetKeyword: "Cutting Molded", // D21
    displayName: "Cutting Molded 大料斬台",
  },
  {
    patterns: [
      "cutting",

      "assistant kepala shift cutting",
      "kepala shift cutting",
      "leader line cutting",
      "adm cutting",
    ],
    sheetKeyword: "Cutting 斬台", // D17
    displayName: "Cutting 斬台",
  },
  {
    patterns: ["buffing"],
    sheetKeyword: "Buffing", // D18
    displayName: "Buffing 打磨",
  },
  {
    patterns: [
      "press molded",
      "pres molded",
      "leader line press molded",
      "leader line pres molded",
    ],
    sheetKeyword: "Press Molded", // D22
    displayName: "Press Molded 模壓",
  },

  // ── D23: Trimming + Embos Automatic + Packing Molded ──────
  // PENTING: Semua pattern D23 harus pakai sheetKeyword YANG SAMA
  // supaya tidak saling timpa saat dikirim ke spreadsheet.
  // Gunakan "Packing Molded" sebagai keyword tunggal karena cell D23
  // pasti mengandung teks tersebut.
  {
    patterns: ["embos automatic", "emboss automatic"],
    sheetKeyword: "Packing Molded", // D23 ← FIX: dulu "Embos Automatic"
    displayName:
      "Trimming, Embos Automatic, Packing Molded ( 修邊, 轉印, 包裝)",
  },
  {
    patterns: ["embos", "emboss"],
    sheetKeyword: "Embos 轉印", // D19
    displayName: "Embos 轉印",
  },
  {
    patterns: ["packing molded"],
    sheetKeyword: "Packing Molded", // D23
    displayName:
      "Trimming, Embos Automatic, Packing Molded ( 修邊, 轉印, 包裝)",
  },
  {
    patterns: ["packing"],
    sheetKeyword: "Packing 包裝", // D20
    displayName: "Packing 包裝",
  },
  {
    patterns: ["trimming", "triming"],
    sheetKeyword: "Packing Molded", // D23
    displayName:
      "Trimming, Embos Automatic, Packing Molded ( 修邊, 轉印, 包裝)",
  },
  {
    patterns: ["leader line trimming", "leader line triming"],
    sheetKeyword: "Packing Molded", // D23
    displayName:
      "Trimming, Embos Automatic, Packing Molded ( 修邊, 轉印, 包裝)",
  },
  {
    patterns: ["molded"],
    sheetKeyword: "Packing Molded", // D23
    displayName:
      "Trimming, Embos Automatic, Packing Molded ( 修邊, 轉印, 包裝)",
  },
  {
    patterns: ["leader line bottom"],
    sheetKeyword: "Packing Molded", // D23
    displayName:
      "Trimming, Embos Automatic, Packing Molded ( 修邊, 轉印, 包裝)",
  },

  // ── CMY & Development ──────────────────────────────────────
  {
    patterns: ["cmy scf", "cmyscf", "cmy"],
    sheetKeyword: "CMY SCF", // D24
    displayName: "CMY SCF 超臨界泡沫",
  },
  {
    patterns: ["development", "pengembangan", "dev"],
    sheetKeyword: "Development", // D25
    displayName: "Development 開發",
  },
];

function getDeptEntry(deptName) {
  if (!deptName) return null;
  const lower = deptName.toLowerCase().trim();
  for (const entry of DEPT_MAP) {
    for (const pat of entry.patterns) {
      if (lower.includes(pat)) return entry;
    }
  }
  return null;
}

/** Cari displayName berdasarkan sheetKeyword (untuk result modal) */
function getDisplayName(keyword) {
  for (const entry of DEPT_MAP) {
    if (entry.sheetKeyword === keyword) return entry.displayName;
  }
  return keyword;
}

/* ── DATE HELPER ── */
function getTodayDateStr() {
  const d = new Date();
  return [
    String(d.getDate()).padStart(2, "0"),
    String(d.getMonth() + 1).padStart(2, "0"),
    d.getFullYear(),
  ].join("/");
}

/* ── PARSER ── */
let allRows = [];

function normStatus(raw) {
  let r = raw
    .toLowerCase()
    .replace(/\bizim\b/g, "izin")
    .replace(/\s+/g, " ")
    .trim();
  if (/izin.*pulang cepat|pulang cepat/.test(r)) return "IZIN_PULANG_CEPAT";
  if (/\b(sudah masuk|sudah hadir|sudah datang)\b/.test(r)) {
    if (/izin.*siang/.test(r)) return "IZIN_SIANG";
    if (/terlambat/.test(r)) return "IZIN_TERLAMBAT";
    return "MASUK";
  }
  if (/\balfa\b/.test(r)) return "ALFA";
  if (/\bsakit\b/.test(r)) return "SAKIT";
  if (/\bcuti\b/.test(r)) return "CUTI";
  if (/izin.*siang/.test(r)) return "IZIN_SIANG";
  if (/izin.*terlambat|terlambat/.test(r)) return "IZIN_TERLAMBAT";
  if (/\bizin\b/.test(r) || /tidak masuk/.test(r)) return "IZIN_TIDAK_MASUK";
  return "LAIN";
}

function counted(s) {
  return ["SAKIT", "ALFA", "CUTI", "IZIN_TIDAK_MASUK"].includes(s);
}

function parse(text) {
  const rows = [];
  text = text.replace(/\r/g, "");
  const lines = text.split("\n");
  let buffer = [];

  function processBuffer(raw) {
    if (!raw) return;
    raw = raw
      .replace(/^\s*\d+\.\s*/, "")
      .replace(/\s+/g, " ")
      .trim();
    const match = raw.match(/\((.*?)\)\s*$/i);
    if (!match) return;
    const sRaw = match[1].trim();
    let before = raw.slice(0, match.index).trim();
    let name = "",
      dept = "";
    const commaIndex = before.indexOf(",");
    if (commaIndex !== -1) {
      name = before.slice(0, commaIndex).trim();
      dept = before.slice(commaIndex + 1).trim();
    } else {
      name = before.trim();
    }
    if (!name) return;
    const status = normStatus(sRaw);
    if (status === "MASUK") return;
    rows.push({ name, dept, sRaw, status });
  }

  for (let line of lines) {
    let t = line.trim();
    if (!t) {
      if (buffer.length) {
        processBuffer(buffer.join(" "));
        buffer = [];
      }
      continue;
    }
    if (/^\*.*\*$/.test(t)) {
      if (buffer.length) {
        processBuffer(buffer.join(" "));
        buffer = [];
      }
      continue;
    }
    if (/^\s*\d+\./.test(t)) {
      if (buffer.length) processBuffer(buffer.join(" "));
      buffer = [t];
    } else {
      if (buffer.length) buffer.push(t);
    }
  }
  if (buffer.length) processBuffer(buffer.join(" "));
  return rows;
}

/* ── BADGE ── */
function badge(status, raw) {
  const m = {
    SAKIT: ["bdg-sakit", "bi-thermometer-half", "SAKIT"],
    ALFA: ["bdg-alfa", "bi-x-circle", "ALFA"],
    CUTI: ["bdg-cuti", "bi-calendar-check", "CUTI"],
    IZIN_SIANG: ["bdg-izin", "bi-clock", "IZIN MASUK SIANG"],
    IZIN_TERLAMBAT: ["bdg-other", "bi-clock-history", "IZIN TERLAMBAT"],
    IZIN_TIDAK_MASUK: ["bdg-other", "bi-person-x", "IZIN TIDAK MASUK"],
    IZIN_PULANG_CEPAT: ["bdg-other", "bi-box-arrow-right", "IZIN PULANG CEPAT"],
    IZIN_LAIN: ["bdg-other", "bi-info-circle", "IZIN"],
    LAIN: ["bdg-other", "bi-question-circle", raw.toUpperCase()],
  };
  const [cls, ico, lbl] = m[status] || m.LAIN;
  return `<span class="bdg ${cls}"><i class="bi ${ico}"></i>${lbl}</span>`;
}

/* ── RENDER ── */
function renderSummary(rows) {
  const cnt = rows.filter((r) => counted(r.status));
  const dm = {};
  cnt.forEach((r) => {
    const d = r.dept || "Lainnya";
    dm[d] = (dm[d] || 0) + 1;
  });
  animCount(document.getElementById("statAbsen"), cnt.length);
  animCount(document.getElementById("statBagian"), Object.keys(dm).length);
  animCount(document.getElementById("statTotal"), rows.length);
  const g = document.getElementById("deptGrid");
  if (!Object.keys(dm).length) {
    g.innerHTML = `<div class="empty-ph w-100"><i class="bi bi-bar-chart-line"></i><p>Belum ada data.<br/>Klik PROSES DATA untuk memulai analisis.</p></div>`;
    return;
  }
  g.innerHTML = Object.entries(dm)
    .sort((a, b) => b[1] - a[1])
    .map(
      ([d, n], i) => `
      <div class="dept-card" style="animation-delay:${i * 45}ms">
        <div class="dept-name">${d}</div>
        <div class="dept-num">${n}</div>
        <div class="dept-lbl">Tidak Hadir</div>
      </div>`
    )
    .join("");
}

function searchLabel(status, raw) {
  const labels = {
    SAKIT: "sakit",
    ALFA: "alfa",
    CUTI: "cuti",
    IZIN_SIANG: "izin masuk siang",
    IZIN_TERLAMBAT: "izin terlambat",
    IZIN_LAIN: "izin",
    IZIN_TIDAK_MASUK: "izin tidak masuk",
    LAIN: raw,
  };
  return (labels[status] || raw).toLowerCase();
}

function renderTable(rows, q = "") {
  const body = document.getElementById("tblBody");
  const f = q
    ? rows.filter((r) => {
        const kw = q.toLowerCase();
        return (
          r.name.toLowerCase().includes(kw) ||
          r.dept.toLowerCase().includes(kw) ||
          r.status.toLowerCase().includes(kw) ||
          r.sRaw.toLowerCase().includes(kw) ||
          searchLabel(r.status, r.sRaw).includes(kw)
        );
      })
    : rows;
  if (!f.length) {
    body.innerHTML = `<tr><td colspan="4"><div class="empty-ph">
      <i class="bi bi-person-x"></i>
      <p>${
        rows.length
          ? "Tidak ada hasil pencarian."
          : "Proses text absensi untuk melihat data detail."
      }</p>
    </div></td></tr>`;
    return;
  }
  body.innerHTML = f
    .map(
      (r, i) => `
    <tr class="row-fade" style="animation-delay:${i * 28}ms">
      <td><span class="row-num">${i + 1}</span></td>
      <td><div class="emp-name">${r.name}</div></td>
      <td><div class="emp-dept">${
        r.dept || '<span style="color:var(--txt3)">—</span>'
      }</div></td>
      <td><div class="td-badge">
        ${badge(r.status, r.sRaw)}
        ${
          !counted(r.status)
            ? '<span class="not-counted-tag">※ tidak dihitung</span>'
            : ""
        }
      </div></td>
    </tr>`
    )
    .join("");
}

/* ── ACTIONS ── */
function prosesData() {
  const btn = document.getElementById("btnProses");
  btn.innerHTML = '<i class="bi bi-arrow-repeat spin me-2"></i>MEMPROSES…';
  btn.style.opacity = ".8";
  setTimeout(() => {
    allRows = parse(document.getElementById("inputText").value);
    renderSummary(allRows);
    renderTable(allRows);
    updateKirimButton();
    btn.innerHTML =
      '<i class="bi bi-lightning-charge-fill me-2"></i>PROSES DATA';
    btn.style.opacity = "";
  }, 250);
}

function bersihkan() {
  document.getElementById("inputText").value = "";
  document.getElementById("searchBox").value = "";
  allRows = [];
  ["statAbsen", "statBagian", "statTotal"].forEach((id) =>
    animCount(document.getElementById(id), 0, 300)
  );
  renderSummary([]);
  renderTable([]);
  updateKirimButton();
}

function filterTable() {
  renderTable(allRows, document.getElementById("searchBox").value);
}

function exportExcel() {
  if (!allRows.length) {
    showToast("Tidak ada data untuk diekspor.", "warning");
    return;
  }
  const data = allRows.map((r, i) => ({
    No: i + 1,
    Nama: r.name,
    Bagian: r.dept,
    Keterangan: r.sRaw,
    Dihitung: counted(r.status) ? "Ya" : "Tidak",
  }));
  try {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Absensi");
    const d = new Date();
    XLSX.writeFile(
      wb,
      `absensi_${d.getFullYear()}${String(d.getMonth() + 1).padStart(
        2,
        "0"
      )}${String(d.getDate()).padStart(2, "0")}.xlsx`
    );
  } catch (e) {
    const csv = [
      "No,Nama,Bagian,Keterangan,Dihitung",
      ...data.map(
        (r) =>
          `${r.No},"${r.Nama}","${r.Bagian}","${r.Keterangan}","${r.Dihitung}"`
      ),
    ].join("\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    a.download = "absensi.csv";
    a.click();
  }
}

/* ── SPREADSHEET INTEGRATION ── */

function computeSheetUpdates() {
  const counts = {};
  const unmatched = [];

  allRows
    .filter((r) => counted(r.status))
    .forEach((r) => {
      const entry = getDeptEntry(r.dept);
      if (entry) {
        const k = entry.sheetKeyword;
        if (!counts[k])
          counts[k] = {
            count: 0,
            displayName: entry.displayName,
            depts: new Set(),
          };
        counts[k].count++;
        counts[k].depts.add(r.dept || "(kosong)");
      } else {
        if (r.dept && !unmatched.includes(r.dept)) unmatched.push(r.dept);
      }
    });

  const updates = Object.entries(counts).map(([dept, v]) => ({
    dept,
    count: v.count,
    displayName: v.displayName,
    sourceNames: Array.from(v.depts),
  }));

  return { updates, unmatched };
}

function kirimKeSpreadsheet() {
  const url = localStorage.getItem("appsScriptUrl");
  if (!url) {
    showConfigModal();
    showToast("Silakan set URL Apps Script terlebih dahulu.", "info");
    return;
  }
  if (!allRows.length) {
    showToast(
      "Tidak ada data. Proses text absensi terlebih dahulu.",
      "warning"
    );
    return;
  }

  const { updates, unmatched } = computeSheetUpdates();
  if (!updates.length) {
    showToast(
      "Tidak ada karyawan yang dihitung sebagai tidak hadir.",
      "warning"
    );
    return;
  }
  showPreviewModal(updates, unmatched);
}

function updateKirimButton() {
  const btn = document.getElementById("btnKirim");
  const hasData = allRows.length > 0 && allRows.some((r) => counted(r.status));
  btn.disabled = !hasData;
  if (hasData) btn.classList.add("has-data");
  else btn.classList.remove("has-data");
}

/* ── CONFIG MODAL ── */
function showConfigModal() {
  const savedUrl = localStorage.getItem("appsScriptUrl") || "";
  const savedDate = localStorage.getItem("targetDate") || getTodayDateStr();
  document.getElementById("appsScriptUrlInput").value = savedUrl;
  document.getElementById("targetDateInput").value = savedDate;
  document.getElementById("configModal").classList.add("open");
}

function closeConfigModal(e) {
  if (e && e.target !== document.getElementById("configModal")) return;
  document.getElementById("configModal").classList.remove("open");
}

function saveConfig() {
  const url = document.getElementById("appsScriptUrlInput").value.trim();
  const date = document.getElementById("targetDateInput").value.trim();
  if (!url) {
    showToast("URL tidak boleh kosong.", "error");
    return;
  }
  if (!url.includes("script.google.com")) {
    showToast("URL tidak valid. Harus dari script.google.com", "error");
    return;
  }
  if (date && !/^\d{2}\/\d{2}\/\d{4}$/.test(date)) {
    showToast("Format tanggal harus DD/MM/YYYY", "error");
    return;
  }
  localStorage.setItem("appsScriptUrl", url);
  if (date) localStorage.setItem("targetDate", date);
  updateUrlStatus();
  document.getElementById("configModal").classList.remove("open");
  showToast("Konfigurasi disimpan!", "success");
}

function updateUrlStatus() {
  const url = localStorage.getItem("appsScriptUrl");
  const dot = document.getElementById("urlStatusDot");
  const txt = document.getElementById("urlStatusTxt");
  if (!dot || !txt) return;
  if (url) {
    dot.className = "url-dot set";
    txt.textContent = "TERHUBUNG";
    txt.style.color = "var(--green)";
  } else {
    dot.className = "url-dot unset";
    txt.textContent = "BELUM TERHUBUNG";
    txt.style.color = "var(--txt3)";
  }
}

/* ── PREVIEW MODAL ── */
function showPreviewModal(updates, unmatched) {
  const date = localStorage.getItem("targetDate") || getTodayDateStr();

  document.getElementById(
    "previewDateInfo"
  ).innerHTML = `<i class="bi bi-calendar2-check" style="color:var(--cyan)"></i>
     Data akan dikirim ke sheet <strong>MEI</strong> tanggal <strong>${date}</strong> — kolom G (Tidak Hadir)`;

  const tbody = document.getElementById("previewTbody");
  tbody.innerHTML = updates
    .map(
      (u) => `
    <tr>
      <td>
        ${u.displayName}
        <span class="preview-dept-key">Keyword: "${u.dept}"</span>
        <span class="preview-dept-key" style="color:var(--txt3)">Dari: ${u.sourceNames.join(
          ", "
        )}</span>
      </td>
      <td>${u.count}</td>
    </tr>`
    )
    .join("");

  const unmatchedDiv = document.getElementById("unmatchedDepts");
  if (unmatched.length) {
    unmatchedDiv.innerHTML = `
      <div class="unmatched-alert">
        <strong><i class="bi bi-exclamation-triangle me-1"></i>Dept Tidak Dikenali (tidak akan dikirim):</strong>
        ${unmatched
          .map((d) => `<span style="color:#ffcc80">"${d}"</span>`)
          .join(", ")}
      </div>`;
  } else {
    unmatchedDiv.innerHTML = "";
  }

  window._pendingUpdates = updates;
  window._pendingDate = date;
  document.getElementById("previewModal").classList.add("open");
}

function closePreviewModal(e) {
  if (e && e.target !== document.getElementById("previewModal")) return;
  document.getElementById("previewModal").classList.remove("open");
}

/* ── JSONP helper ── */
function fetchJSONP(url) {
  return new Promise((resolve, reject) => {
    const cbName = "absensiCb_" + Date.now();
    let scriptEl = null;
    let settled = false;

    const cleanup = () => {
      delete window[cbName];
      if (scriptEl && scriptEl.parentNode)
        scriptEl.parentNode.removeChild(scriptEl);
    };

    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(
        new Error(
          "Timeout (20 detik) — Apps Script tidak merespons. Coba lagi."
        )
      );
    }, 20000);

    window[cbName] = (data) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      cleanup();
      resolve(data);
    };

    scriptEl = document.createElement("script");
    scriptEl.src = url + "&callback=" + cbName;
    scriptEl.onerror = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      cleanup();
      reject(
        new Error(
          "Gagal memuat script Apps Script. Pastikan URL benar dan sudah di-deploy."
        )
      );
    };
    document.head.appendChild(scriptEl);
  });
}

async function confirmSend() {
  document.getElementById("previewModal").classList.remove("open");

  const updates = window._pendingUpdates;
  const date = window._pendingDate;
  const url = localStorage.getItem("appsScriptUrl");

  if (!updates || !url) {
    showToast("Data atau URL kosong.", "error");
    return;
  }

  const btn = document.getElementById("btnKirim");
  const origHTML = btn.innerHTML;

  btn.innerHTML = '<i class="bi bi-arrow-repeat spin me-2"></i>MENGIRIM…';
  btn.disabled = true;

  window._keywordDisplayMap = {};
  updates.forEach((u) => {
    window._keywordDisplayMap[u.dept] = u.displayName;
  });

  const payload = {
    date,
    updates: updates.map((u) => ({
      dept: u.dept,
      count: u.count,
    })),
  };

  const fullUrl = url + "?data=" + encodeURIComponent(JSON.stringify(payload));

  try {
    showToast("Mengirim data ke spreadsheet…", "info");
    const json = await fetchJSONP(fullUrl);
    console.log(json);
    if (json.success) {
      showToast("Data berhasil dikirim!", "success");
      showResultModal(json.results, date);
    } else {
      showToast("Gagal mengirim data.", "error");
    }
  } catch (err) {
    console.error(err);
    showToast(err.message, "error");
  } finally {
    btn.innerHTML = origHTML;
    updateKirimButton();
  }
}

/* ── RESULT MODAL ── */
function showResultModal(results, date) {
  document.getElementById(
    "resultDateInfo"
  ).innerHTML = `<i class="bi bi-calendar2-check" style="color:var(--green)"></i> Hasil pengiriman data tanggal <strong>${date}</strong>`;

  const displayMap = window._keywordDisplayMap || {};

  document.getElementById("resultItems").innerHTML = results
    .map((r) => {
      const label = displayMap[r.dept] || getDisplayName(r.dept);
      if (r.status === "updated") {
        return `<div class="result-item result-ok">
        <span><i class="bi bi-check-circle-fill me-2"></i>${label}</span>
        <span>
          <span class="result-row-badge">Baris ${r.row}</span>
          &nbsp;<strong>${r.count}</strong> orang
        </span>
      </div>`;
      } else {
        return `<div class="result-item result-err">
        <span><i class="bi bi-x-circle-fill me-2"></i>${label}</span>
        <span style="font-size:10px; opacity:0.7">Tidak ditemukan di sheet</span>
      </div>`;
      }
    })
    .join("");

  document.getElementById("resultModal").classList.add("open");
}

function closeResultModal(e) {
  if (e && e.target !== document.getElementById("resultModal")) return;
  document.getElementById("resultModal").classList.remove("open");
}

/* ── TOAST ── */
function showToast(msg, type = "info") {
  const icons = {
    success: "bi-check-circle-fill",
    error: "bi-x-circle-fill",
    warning: "bi-exclamation-triangle-fill",
    info: "bi-info-circle-fill",
  };
  const c = document.getElementById("toastContainer");
  const t = document.createElement("div");
  t.className = `toast ${type}`;
  t.innerHTML = `<i class="bi ${
    icons[type] || icons.info
  }"></i><span>${msg}</span>`;
  c.appendChild(t);
  setTimeout(() => {
    t.classList.add("toast-out");
    setTimeout(() => t.remove(), 350);
  }, 4000);
}

/* ── INIT ── */
window.addEventListener("DOMContentLoaded", () => {
  updateUrlStatus();
  if (!localStorage.getItem("targetDate")) {
    localStorage.setItem("targetDate", getTodayDateStr());
  }
  prosesData();
});
