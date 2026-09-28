/**
 * Application logic for Flood Situation Report Infographic
 * Compatible with GitHub Pages and Google Apps Script (GAS)
 */

const SHEET_ID = "1-oqfYnCyY2djg9WbASAfH21nX8r39R6j-z0UUs-vd4g";
const SHEET_NAME = "data";

const THAI_MONTHS = [
  "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
  "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
];

// 38 column keys matching schema
const COLUMN_KEYS = [
  "date", "time", "water_level", "villages", "households", "people", "routes_cut",
  "trend", "hosp_affected", "staff_total", "staff_affected", "staff_absent",
  "beds_total", "beds_current", "beds_available", "er_ready", "ambulance_ready",
  "ems_status", "referral_status", "dest_hospital", "bedridden", "oxygen",
  "dialysis", "chronic_med", "urgent_evac", "medicine", "food_water",
  "electricity", "tap_water", "internet", "generator", "fuel",
  "support_1", "support_2", "support_3", "hosp_status", "reporter_name", "reporter_pos"
];

let rawSheetRows = [];
let parsedRecords = [];
let currentRecord = null;
let templateImage = null;
let preloadedAssets = {};

// -------------------------------------------------------------
// Utilities: Thai Date & Time formatting
// -------------------------------------------------------------
function formatThaiDate(dateStr) {
  if (!dateStr) return "";
  const parts = String(dateStr).trim().split(/[/.-]/);
  if (parts.length === 3) {
    let day, month, year;
    if (parts[0].length === 4) {
      year = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10);
      day = parseInt(parts[2], 10);
    } else {
      day = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10);
      year = parseInt(parts[2], 10);
    }
    if (year < 2400) year += 543;
    if (month >= 1 && month <= 12) {
      return `${day} ${THAI_MONTHS[month]} ${year}`;
    }
  }
  return dateStr;
}

function formatThaiTime(timeStr) {
  if (!timeStr) return "";
  const m = String(timeStr).trim().match(/^(\d{1,2})[:.](\d{2})/);
  if (m) {
    const hh = parseInt(m[1], 10);
    const mm = m[2];
    return `${hh < 10 ? '0' + hh : hh}.${mm}`;
  }
  return timeStr;
}

function showToast(message) {
  const toast = document.getElementById("toast");
  if (!toast) return;
  toast.innerText = message;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 3000);
}

// -------------------------------------------------------------
// Preload Assets & Template
// -------------------------------------------------------------
async function initAssets() {
  // Load base template 000.png
  templateImage = new Image();
  templateImage.crossOrigin = "anonymous";
  const templatePromise = new Promise((resolve) => {
    templateImage.onload = () => resolve();
    templateImage.onerror = () => {
      console.warn("Could not load 000.png locally, trying fallback...");
      resolve();
    };
  });
  templateImage.src = "000.png";

  // Preload small badge assets
  const assetPromises = [];
  if (typeof ASSETS !== "undefined") {
    for (const [key, src] of Object.entries(ASSETS)) {
      const img = new Image();
      const p = new Promise((resolve) => {
        img.onload = () => {
          preloadedAssets[key] = img;
          resolve();
        };
        img.onerror = () => resolve();
      });
      img.src = src;
      assetPromises.push(p);
    }
  }

  // Ensure Prompt font is loaded
  let fontPromise = Promise.resolve();
  if (document.fonts) {
    fontPromise = Promise.all([
      document.fonts.load("bold 23px Prompt"),
      document.fonts.load("bold 38px Prompt"),
      document.fonts.load("bold 28px Prompt"),
      document.fonts.load("500 16px Prompt"),
      document.fonts.load("14px Prompt")
    ]);
  }

  await Promise.all([templatePromise, ...assetPromises, fontPromise]);
}

// -------------------------------------------------------------
// Data Fetching: JSONP or Fetch fallback
// -------------------------------------------------------------
function loadSheetData() {
  showToast("กำลังดึงข้อมูลจาก Google Sheets...");

  // If inside Google Apps Script Web App environment
  if (typeof google !== "undefined" && google.script && google.script.run) {
    google.script.run
      .withSuccessHandler((res) => {
        if (res && res.data) {
          processRawRows(res.data);
          showToast("โหลดข้อมูลสำเร็จ (GAS)");
        }
      })
      .withFailureHandler((err) => {
        console.error("GAS error:", err);
        fallbackLocalFetch();
      })
      .getSheetData();
    return;
  }

  // Standard Web / GitHub Pages: JSONP bypasses CORS
  const callbackName = "handleSheetData_" + Math.floor(Math.random() * 100000);
  window[callbackName] = function (response) {
    try {
      delete window[callbackName];
      if (script.parentNode) script.parentNode.removeChild(script);

      if (response && response.table) {
        const rows = [];
        // Header
        const headers = response.table.cols.map((c) => (c ? c.label : ""));
        rows.push(headers);

        // Data rows
        response.table.rows.forEach((r) => {
          const row = r.c.map((cell) => {
            if (!cell || cell.v === null || cell.v === undefined) return "";
            // If date cell
            if (cell.f) return cell.f;
            return String(cell.v);
          });
          if (row.some((val) => val.trim())) {
            rows.push(row);
          }
        });

        processRawRows(rows);
        showToast("ดึงข้อมูลจาก Google Sheets สำเร็จ");
      } else {
        fallbackLocalFetch();
      }
    } catch (e) {
      console.error(e);
      fallbackLocalFetch();
    }
  };

  const script = document.createElement("script");
  script.src = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=responseHandler:${callbackName}&sheet=${SHEET_NAME}`;
  script.onerror = () => {
    delete window[callbackName];
    fallbackLocalFetch();
  };
  document.body.appendChild(script);
}

function fallbackLocalFetch() {
  fetch("sheet_data.json")
    .then((res) => res.json())
    .then((rows) => {
      processRawRows(rows);
      showToast("ใช้ข้อมูลสำรองล่าสุด (sheet_data.json)");
    })
    .catch((err) => {
      console.error("Failed to load sheet data:", err);
      showToast("ไม่สามารถดึงข้อมูลได้ โปรดตรวจสอบการเชื่อมต่อ");
    });
}

function processRawRows(rows) {
  if (!rows || rows.length < 2) return;
  rawSheetRows = rows;
  parsedRecords = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[0] || !r[0].trim()) continue;

    const rec = {};
    for (let c = 0; c < COLUMN_KEYS.length; c++) {
      const key = COLUMN_KEYS[c];
      rec[key] = r[c] !== undefined ? String(r[c]).trim() : "";
    }
    rec._rowIndex = i;
    parsedRecords.push(rec);
  }

  populateDateSelector();
}

function populateDateSelector() {
  const select = document.getElementById("dateSelect");
  if (!select) return;
  select.innerHTML = "";

  parsedRecords.forEach((rec, idx) => {
    const opt = document.createElement("option");
    opt.value = idx;
    const thaiDate = formatThaiDate(rec.date);
    opt.textContent = `${thaiDate || rec.date} (เวลา ${formatThaiTime(rec.time) || "09.00"} น.)`;
    select.appendChild(opt);
  });

  if (parsedRecords.length > 0) {
    select.value = "0";
    selectRecord(0);
  }
}

function selectRecord(index) {
  if (parsedRecords[index]) {
    // Clone record so user edits don't overwrite raw immediately
    currentRecord = { ...parsedRecords[index] };
    updateFormFields(currentRecord);
    renderReport();
  }
}

// -------------------------------------------------------------
// Form Synchronization
// -------------------------------------------------------------
function updateFormFields(rec) {
  COLUMN_KEYS.forEach((key) => {
    const el = document.getElementById("field_" + key);
    if (el) {
      el.value = rec[key] || "";
    }
  });
}

function bindFormEvents() {
  COLUMN_KEYS.forEach((key) => {
    const el = document.getElementById("field_" + key);
    if (el) {
      el.addEventListener("input", (e) => {
        if (!currentRecord) currentRecord = {};
        currentRecord[key] = e.target.value;
        debounceRender();
      });
    }
  });

  const dateSelect = document.getElementById("dateSelect");
  if (dateSelect) {
    dateSelect.addEventListener("change", (e) => {
      selectRecord(parseInt(e.target.value, 10));
    });
  }

  const btnRefresh = document.getElementById("btnRefresh");
  if (btnRefresh) {
    btnRefresh.addEventListener("click", () => loadSheetData());
  }

  const btnPrint = document.getElementById("btnPrint");
  if (btnPrint) {
    btnPrint.addEventListener("click", () => window.print());
  }

  const btnDownload = document.getElementById("btnDownload");
  if (btnDownload) {
    btnDownload.addEventListener("click", () => downloadInfographic());
  }
}

let renderTimer = null;
function debounceRender() {
  if (renderTimer) clearTimeout(renderTimer);
  renderTimer = setTimeout(() => {
    renderReport();
  }, 100);
}

// -------------------------------------------------------------
// Canvas Infographic Rendering
// -------------------------------------------------------------
function renderReport() {
  const canvas = document.getElementById("reportCanvas");
  if (!canvas || !templateImage || !currentRecord) return;

  const ctx = canvas.getContext("2d");
  const rec = currentRecord;

  // Clear & Draw template background
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(templateImage, 0, 0, canvas.width, canvas.height);

  const NAVY = "#002d62";
  const RED = "#e51c24";
  const DARK_TEAL = "#004b6e";

  function drawCentered(xCenter, y, text, font, fill) {
    if (!text) return;
    ctx.font = font;
    ctx.fillStyle = fill;
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    ctx.fillText(String(text), xCenter, y);
  }

  function drawRightAligned(xRight, y, text, font, fill) {
    if (!text) return;
    ctx.font = font;
    ctx.fillStyle = fill;
    ctx.textAlign = "right";
    ctx.textBaseline = "top";
    ctx.fillText(String(text), xRight, y);
  }

  function drawLeftAligned(xLeft, y, text, font, fill) {
    if (!text) return;
    ctx.font = font;
    ctx.fillStyle = fill;
    ctx.textAlign = "left";
    ctx.textBaseline = "top";
    ctx.fillText(String(text), xLeft, y);
  }

  // -------------------------------------------------------------
  // 1. Header: Date & Time
  // -------------------------------------------------------------
  const dateFormatted = formatThaiDate(rec.date);
  const timeFormatted = formatThaiTime(rec.time);
  if (dateFormatted) {
    drawLeftAligned(338, 210, dateFormatted, "bold 23px Prompt", NAVY);
  }
  if (timeFormatted) {
    drawLeftAligned(658, 210, timeFormatted, "bold 23px Prompt", NAVY);
  }

  // -------------------------------------------------------------
  // 2. Box 1: สถานการณ์พื้นที่ (Area Situation)
  // -------------------------------------------------------------
  drawCentered(374, 338, rec.water_level, "bold 38px Prompt", RED);
  drawCentered(374, 390, rec.villages, "bold 38px Prompt", RED);
  drawCentered(374, 444, rec.households, "bold 38px Prompt", RED);
  drawCentered(374, 499, rec.people, "bold 38px Prompt", RED);
  drawCentered(374, 556, rec.routes_cut, "bold 38px Prompt", RED);

  // Trend Pill & Checkboxes
  if (rec.trend === "เพิ่มขึ้น" && preloadedAssets["pill_trend"]) {
    ctx.drawImage(preloadedAssets["pill_trend"], 248, 618);
  } else if (rec.trend) {
    drawDynamicPill(ctx, 248, 618, rec.trend, rec.trend.includes("เพิ่ม") ? "red" : "green");
  }
  if (preloadedAssets["checkboxes"]) {
    ctx.drawImage(preloadedAssets["checkboxes"], 245, 660);
  }

  // -------------------------------------------------------------
  // 3. Box 2: สถานการณ์โรงพยาบาล (Hospital Situation)
  // -------------------------------------------------------------
  drawRightAligned(920, 398, rec.staff_total, "bold 28px Prompt", DARK_TEAL);
  drawRightAligned(920, 438, rec.staff_affected, "bold 28px Prompt", DARK_TEAL);
  drawRightAligned(920, 480, rec.staff_absent, "bold 28px Prompt", DARK_TEAL);

  // Beds
  [
    { val: rec.beds_total, y: 526 },
    { val: rec.beds_current, y: 567 },
    { val: rec.beds_available, y: 608 }
  ].forEach((item) => {
    if (item.val) {
      if (/^\d+$/.test(item.val)) {
        drawRightAligned(920, item.y, item.val, "bold 28px Prompt", DARK_TEAL);
      } else {
        drawCentered(898, item.y + 4, item.val, "bold 20px Prompt", DARK_TEAL);
      }
    }
  });

  // ICU / ER Ready Badge
  const erVal = rec.er_ready.trim();
  if (["ไม่ใช่", "ไม่พร้อม", "งดให้บริการ"].includes(erVal)) {
    if (preloadedAssets["badge_icu_red"]) {
      ctx.drawImage(preloadedAssets["badge_icu_red"], 854, 649);
    } else {
      drawDynamicPill(ctx, 854, 649, "ไม่ใช่", "red");
    }
  } else {
    if (preloadedAssets["badge_icu_green"]) {
      ctx.drawImage(preloadedAssets["badge_icu_green"], 854, 649);
    } else {
      drawDynamicPill(ctx, 854, 649, "พร้อมใช้", "green");
    }
  }

  // -------------------------------------------------------------
  // 4. Box 3: ระบบการแพทย์ฉุกเฉิน (EMS)
  // -------------------------------------------------------------
  if (rec.ambulance_ready) {
    if (/^\d+$/.test(rec.ambulance_ready)) {
      drawCentered(422, 788, rec.ambulance_ready, "bold 28px Prompt", DARK_TEAL);
    } else {
      drawCentered(422, 792, rec.ambulance_ready, "bold 20px Prompt", DARK_TEAL);
    }
  }

  if (["ปกติ", "พร้อม"].includes(rec.ems_status) && preloadedAssets["ems"]) {
    ctx.drawImage(preloadedAssets["ems"], 360, 835);
  } else if (rec.ems_status) {
    drawDynamicPill(ctx, 360, 835, rec.ems_status, "green");
  }

  if (["ปกติ", "พร้อม"].includes(rec.referral_status) && preloadedAssets["ref"]) {
    ctx.drawImage(preloadedAssets["ref"], 360, 895);
  } else if (rec.referral_status) {
    drawDynamicPill(ctx, 360, 895, rec.referral_status, "green");
  }

  // Coordinated hospital
  const destHosp = rec.dest_hospital ? rec.dest_hospital.trim() : "รอระบุ";
  if ((!rec.dest_hospital || destHosp === "รอระบุ") && preloadedAssets["dest_hosp"]) {
    ctx.drawImage(preloadedAssets["dest_hosp"], 380, 955);
  } else {
    drawCentered(422, 960, destHosp, "bold 20px Prompt", NAVY);
  }

  // -------------------------------------------------------------
  // 5. Box 4: กลุ่มเปราะบางในพื้นที่ (Vulnerable Groups)
  // -------------------------------------------------------------
  drawRightAligned(898, 783, rec.bedridden, "bold 32px Prompt", RED);
  drawRightAligned(898, 840, rec.oxygen, "bold 32px Prompt", RED);
  drawRightAligned(898, 897, rec.dialysis, "bold 32px Prompt", RED);
  drawRightAligned(898, 954, rec.chronic_med, "bold 32px Prompt", RED);
  drawRightAligned(898, 1010, rec.urgent_evac, "bold 32px Prompt", RED);

  // -------------------------------------------------------------
  // 6. Box 5: ยา / เวชภัณฑ์ / สาธารณูปโภค (Medicine & Utilities)
  // -------------------------------------------------------------
  // Medicine
  if (rec.medicine === "เพียงพอ" && preloadedAssets["pill_peangphor"]) {
    ctx.drawImage(preloadedAssets["pill_peangphor"], 215, 1083);
  } else if (rec.medicine) {
    drawDynamicPill(ctx, 215, 1083, rec.medicine, "green");
  }
  if (preloadedAssets["yellow_note"]) {
    ctx.drawImage(preloadedAssets["yellow_note"], 240, 1118);
  }

  // Food / Water
  if (rec.food_water === "เพียงพอ" && preloadedAssets["pill_peangphor"]) {
    ctx.drawImage(preloadedAssets["pill_peangphor"], 215, 1148);
  } else if (rec.food_water) {
    drawDynamicPill(ctx, 215, 1148, rec.food_water, "green");
  }

  // Electricity
  if (rec.electricity === "ปกติ" && preloadedAssets["pill_pokati"]) {
    ctx.drawImage(preloadedAssets["pill_pokati"], 215, 1190);
  } else if (rec.electricity) {
    drawDynamicPill(ctx, 215, 1190, rec.electricity, "green");
  }

  // Tap water
  if (rec.tap_water === "ปกติ" && preloadedAssets["pill_pokati"]) {
    ctx.drawImage(preloadedAssets["pill_pokati"], 215, 1233);
  } else if (rec.tap_water) {
    drawDynamicPill(ctx, 215, 1233, rec.tap_water, "green");
  }
  if (preloadedAssets["water_note"]) {
    ctx.drawImage(preloadedAssets["water_note"], 310, 1224);
  }

  // Internet
  if (rec.internet === "ปกติ" && preloadedAssets["pill_pokati"]) {
    ctx.drawImage(preloadedAssets["pill_pokati"], 215, 1283);
  } else if (rec.internet) {
    drawDynamicPill(ctx, 215, 1283, rec.internet, "green");
  }

  // Generator
  if (["ปกติ", "พร้อม"].includes(rec.generator) && preloadedAssets["pill_prom"]) {
    ctx.drawImage(preloadedAssets["pill_prom"], 218, 1324);
  } else if (rec.generator) {
    drawDynamicPill(ctx, 218, 1324, rec.generator, "green");
  }

  // Fuel
  if (rec.fuel === "200" && preloadedAssets["pill_fuel"]) {
    ctx.drawImage(preloadedAssets["pill_fuel"], 218, 1363);
  } else if (rec.fuel) {
    const fuelText = /^\d+$/.test(rec.fuel) ? `${rec.fuel} ลิตร` : rec.fuel;
    drawDynamicPill(ctx, 218, 1363, fuelText, "blue");
  }

  // -------------------------------------------------------------
  // 7. Box 6: สิ่งที่ต้องการสนับสนุนจากจังหวัด (Provincial Support)
  // -------------------------------------------------------------
  if (preloadedAssets["b6_1"]) ctx.drawImage(preloadedAssets["b6_1"], 580, 1146);
  if (preloadedAssets["b6_2"]) ctx.drawImage(preloadedAssets["b6_2"], 580, 1220);
  if (preloadedAssets["b6_3"]) ctx.drawImage(preloadedAssets["b6_3"], 580, 1298);

  // -------------------------------------------------------------
  // 8. Footer: Hospital Status & Reporter
  // -------------------------------------------------------------
  if (preloadedAssets["pill_status"]) {
    ctx.drawImage(preloadedAssets["pill_status"], 225, 1418);
  }

  const reporterName = rec.reporter_name || "รพ.ไทรโยค";
  drawLeftAligned(864, 1410, reporterName, "500 15px Prompt", NAVY);

  if (rec.reporter_pos) {
    drawFittedText(ctx, 864, 1435, 105, rec.reporter_pos, "Prompt", 13, 8, NAVY);
  }

  if (timeFormatted) {
    drawLeftAligned(854, 1461, `${timeFormatted} น.`, "500 15px Prompt", NAVY);
  }
}

// -------------------------------------------------------------
// Canvas Drawing Helpers
// -------------------------------------------------------------
function drawDynamicPill(ctx, x, y, text, type = "green") {
  ctx.save();
  ctx.font = "bold 20px Prompt";
  const tw = ctx.measureText(text).width;
  const h = 34;
  const w = tw + 50;

  let bg = "#009944";
  if (type === "red") bg = "#e51c24";
  if (type === "blue") bg = "#0066b2";

  // Rounded rectangle
  ctx.fillStyle = bg;
  ctx.beginPath();
  ctx.roundRect(x, y, w, h, h / 2);
  ctx.fill();

  // Circle icon
  ctx.fillStyle = "#ffffff";
  const cx = x + 17;
  const cy = y + h / 2;
  ctx.beginPath();
  ctx.arc(cx, cy, 11, 0, Math.PI * 2);
  ctx.fill();

  // Checkmark or X
  ctx.strokeStyle = bg;
  ctx.lineWidth = 2.5;
  ctx.lineCap = "round";
  ctx.beginPath();
  if (type === "green") {
    ctx.moveTo(cx - 5, cy);
    ctx.lineTo(cx - 1, cy + 4);
    ctx.lineTo(cx + 5, cy - 4);
  } else if (type === "red") {
    ctx.moveTo(cx - 4, cy - 4);
    ctx.lineTo(cx + 4, cy + 4);
    ctx.moveTo(cx - 4, cy + 4);
    ctx.lineTo(cx + 4, cy - 4);
  }
  ctx.stroke();

  // Text
  ctx.fillStyle = "#ffffff";
  ctx.textBaseline = "middle";
  ctx.textAlign = "left";
  ctx.fillText(text, x + 35, cy);
  ctx.restore();
}

function drawFittedText(ctx, xLeft, y, maxW, text, fontFamily, baseSize = 13, minSize = 8, fill = "#002d62") {
  ctx.save();
  ctx.fillStyle = fill;
  ctx.textBaseline = "top";
  ctx.textAlign = "left";

  for (let sz = baseSize; sz >= minSize; sz--) {
    ctx.font = `${sz}px ${fontFamily}`;
    const tw = ctx.measureText(text).width;
    if (tw <= maxW || sz === minSize) {
      ctx.fillText(text, xLeft, y);
      ctx.restore();
      return;
    }
  }
  ctx.restore();
}

function downloadInfographic() {
  const canvas = document.getElementById("reportCanvas");
  if (!canvas || !currentRecord) return;

  const dateStr = currentRecord.date ? currentRecord.date.replace(/[/.-]/g, "_") : "report";
  const link = document.createElement("a");
  link.download = `water_report_${dateStr}.png`;
  link.href = canvas.toDataURL("image/png");
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast("ดาวน์โหลดรูปภาพเรียบร้อยแล้ว");
}

// -------------------------------------------------------------
// App Initialization
// -------------------------------------------------------------
window.addEventListener("DOMContentLoaded", async () => {
  bindFormEvents();
  await initAssets();
  loadSheetData();
});
