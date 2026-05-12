/**
 * ═══════════════════════════════════════════════════════
 *  ABSENSI PARSER — Google Apps Script
 *  Terima data dari web app dan update kolom G sheet MEI
 * ═══════════════════════════════════════════════════════
 *
 *  CARA DEPLOY:
 *  1. Buka spreadsheet → Extensions → Apps Script
 *  2. Hapus kode default, paste seluruh kode ini
 *  3. Deploy → New deployment → Web app
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  4. Copy URL deployment, paste ke konfigurasi web app
 *
 *  CATATAN:
 *  - Setiap kali kode diubah, buat New Deployment baru
 *  - Sheet ID sudah tertanam di bawah, sesuaikan jika perlu
 * ═══════════════════════════════════════════════════════
 */

// ── Konstanta ──
var SPREADSHEET_ID = "1WOcj48CLuC3lWXDp-KW0zLw3E_3xdm-8jYzzxQ4-GD8";
var SHEET_NAME = "MEI";
var COL_G = 7; // Kolom G = index 7 (1-based)

// ────────────────────────────────────────────────────────
//  doGet — Entry point (dipanggil dari web app via fetch)
// ────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const data = JSON.parse(e.parameter.data);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);

    const lastRow = sh.getLastRow();

    // Ambil kolom B sampai D (index: [0]=B tanggal, [1]=C kategori, [2]=D departemen)
    const rangeVals = sh.getRange(1, 2, lastRow, 3).getValues();

    // Cari range tanggal
    const dateRange = findDateRange(rangeVals, data.date);

    if (dateRange.startIdx === -1) {
      return jsonpOrJson(
        {
          success: false,
          error:
            "Tanggal " +
            data.date +
            " tidak ditemukan di kolom B sheet " +
            SHEET_NAME +
            ". Pastikan tanggal sudah ada di spreadsheet dan format DD/MM/YYYY sesuai.",
        },
        e.parameter.callback
      );
    }

    const results = [];

    data.updates.forEach((u) => {
      let found = false;

      // Cari hanya dalam range tanggal tersebut
      for (let i = dateRange.startIdx; i <= dateRange.endIdx; i++) {
        // column D ada di index 2 karena range mulai dari B
        const cell = String(rangeVals[i][2] || "").trim();

        if (cell.includes(u.dept)) {
          sh.getRange(i + 1, COL_G).setValue(u.count);

          results.push({
            dept: u.dept,
            row: i + 1,
            count: u.count,
            status: "updated",
          });

          found = true;
          break;
        }
      }

      if (!found) {
        results.push({
          dept: u.dept,
          status: "not_found",
        });
      }
    });

    return jsonpOrJson(
      {
        success: true,
        results,
      },
      e.parameter.callback
    );
  } catch (err) {
    return jsonpOrJson(
      {
        success: false,
        error: err.toString(),
      },
      e.parameter.callback
    );
  }
}

// ────────────────────────────────────────────────────────
//  Helpers
// ────────────────────────────────────────────────────────

/**
 * Konversi nilai sel kolom B ke string DD/MM/YYYY.
 * Mendukung: Date object, string berformat "DD/MM/YYYY",
 *            atau string campuran seperti "Senin 11/05/2026".
 */
function extractDateStr(val) {
  if (!val && val !== 0) return "";

  // Jika Date object
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }

  // Convert ke string
  var str = String(val).trim();

  // Cari pola tanggal di dalam string
  var match = str.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);

  if (match) {
    return (
      pad(parseInt(match[1])) + "/" + pad(parseInt(match[2])) + "/" + match[3]
    );
  }

  return "";
}

/** Zero-pad angka jadi 2 digit */
function pad(n) {
  return n < 10 ? "0" + n : String(n);
}

/**
 * Temukan rentang baris (0-based index) untuk tanggal target.
 * Mengembalikan { startIdx, endIdx } atau { startIdx: -1 }
 *
 * Logika: kolom B hanya berisi nilai pada baris PERTAMA dari merged cell.
 * Baris berikutnya dalam merged range itu kosong.
 * Kita cari baris pertama yang cocok (startIdx), lalu terus
 * sampai ketemu baris non-kosong berikutnya (= tanggal lain / total row).
 */
function findDateRange(rangeVals, targetDate) {
  var startIdx = -1;
  var endIdx = rangeVals.length - 1;

  for (var i = 0; i < rangeVals.length; i++) {
    // rangeVals = B:C:D
    // B = tanggal
    // C = kategori
    // D = departemen

    var colBVal = rangeVals[i][0];
    var colDVal = String(rangeVals[i][2] || "").trim();

    // DEBUG
    Logger.log("ROW: " + (i + 1));
    Logger.log("RAW COL B: " + colBVal);

    var cellDate = extractDateStr(colBVal);

    Logger.log("PARSED DATE: " + cellDate);
    Logger.log("TARGET DATE: " + targetDate);

    // Jika cell B kosong (bagian merged row)
    if (!colBVal && colBVal !== 0) {
      // Kalau sudah ketemu tanggal sebelumnya
      // lalu ketemu row total -> stop
      if (
        startIdx !== -1 &&
        colDVal.toLowerCase().indexOf("total keseluruhan") !== -1
      ) {
        endIdx = i - 1;
        break;
      }

      continue;
    }

    // Jika tanggal cocok
    if (cellDate === targetDate) {
      if (startIdx === -1) {
        startIdx = i;

        Logger.log("DATE FOUND AT ROW: " + (i + 1));
      }
    }

    // Jika sudah menemukan tanggal lalu ketemu tanggal lain
    else if (startIdx !== -1 && cellDate !== "") {
      endIdx = i - 1;

      Logger.log("END RANGE AT ROW: " + i);
      break;
    }
  }

  // Rapikan total row di bawah
  if (startIdx !== -1) {
    for (var j = endIdx; j >= startIdx; j--) {
      var dVal = String(rangeVals[j][2] || "")
        .trim()
        .toLowerCase();

      if (dVal.indexOf("total") !== -1) {
        endIdx = j - 1;
      } else {
        break;
      }
    }
  }

  Logger.log("FINAL RANGE: " + startIdx + " -> " + endIdx);

  return {
    startIdx: startIdx,
    endIdx: endIdx,
  };
}

/** Buat ContentService response — support JSONP callback */
function jsonResponse(obj) {
  var json = JSON.stringify(obj);
  // Tidak bisa akses 'e' di sini, jadi jsonResponse hanya untuk non-JSONP.
  // Untuk JSONP, lihat fungsi jsonpOrJson di bawah.
  return ContentService.createTextOutput(json).setMimeType(
    ContentService.MimeType.JSON
  );
}

/** Buat response JSON atau JSONP tergantung ada tidaknya parameter callback */
function jsonpOrJson(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    // JSONP: wrap dalam function call — menghindari CORS
    return ContentService.createTextOutput(
      callback + "(" + json + ");"
    ).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(
    ContentService.MimeType.JSON
  );
}

// ────────────────────────────────────────────────────────
//  TEST — Jalankan fungsi ini di editor untuk debug
// ────────────────────────────────────────────────────────
function testManual() {
  var fakeE = {
    parameter: {
      data: JSON.stringify({
        date: "11/05/2026",
        updates: [
          { dept: "管理部", count: 1 },
          { dept: "CMY SCF 工務", count: 1 },
          { dept: "品管", count: 1 },
        ],
      }),
    },
  };
  var result = doGet(fakeE);
  Logger.log(result.getContent());
}
