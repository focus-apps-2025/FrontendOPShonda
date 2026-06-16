import JSZip from 'jszip';
import { apiClient } from '../api/client';

// ============================================================
// HELPERS (unchanged)
// ============================================================
function escHtml(str: string): string {
  return String(str)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;")
    .replace(/\n/g, "<br>");
}
function isImageUrl(str: string): boolean {
  return /\.(jpg|jpeg|png|gif|webp|bmp)(\?.*)?$/i.test(str) || str.startsWith("data:image/");
}
function getAnswerString(response: any, questionId: string): string {
  if (!response?.answers) return "";
  const answer = response.answers[questionId];
  if (answer === null || answer === undefined || answer === "") return "";
  if (typeof answer === "object" && answer !== null) {
    if (answer.status) return String(answer.status);
    if (answer.chassisNumber) return String(answer.chassisNumber);
    if (answer.remark) return String(answer.remark);
    if (answer.zonesData) { const z = Object.keys(answer.zonesData); if (z.length) return `Zones: ${z.join(", ")}`; }
    if (Array.isArray(answer)) return answer.join(", ");
    try { return JSON.stringify(answer); } catch { return String(answer); }
  }
  return String(answer);
}
function getCombinedAnswer(responses: any[], question: any): string {
  if (!question || !responses.length) return "";
  const u = new Set<string>();
  responses.forEach(r => { const a = getAnswerString(r, question.id); if (a) u.add(a); });
  return Array.from(u).join(", ");
}
function getResponseTimestamp(r: any): string | undefined {
  return r.timestamp || r.createdAt || r.submittedAt;
}
function base64ToDataUrl(raw: string, mime = "image/png"): string {
  if (!raw) return "";
  if (raw.startsWith("data:")) return raw;
  return `data:${mime};base64,${raw}`;
}
async function fetchAssetBase64(path: string): Promise<string> {
  try {
    const res = await fetch(path);
    if (!res.ok) return "";
    const blob = await res.blob();
    return await new Promise<string>((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve((reader.result as string).split(",")[1] ?? "");
      reader.readAsDataURL(blob);
    });
  } catch { return ""; }
}
async function fetchImageFromUrl(url: string): Promise<string | null> {
  try {
    console.log("Fetching image from URL:", url);

    // Try direct fetch first (works if CORS is configured)
    try {
      const response = await fetch(url, {
        mode: 'cors',
        credentials: 'omit'
      });
      if (response.ok) {
        const blob = await response.blob();
        if (blob.type.startsWith('image/')) {
          return await new Promise((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result as string);
            reader.onerror = () => resolve(null);
            reader.readAsDataURL(blob);
          });
        }
      }
    } catch (e) {
      console.log("Direct fetch failed, trying proxy...", e);
    }

    // Try with proxy (cors-anywhere)
    const proxyUrl = `https://cors-anywhere.herokuapp.com/${url}`;
    const response = await fetch(proxyUrl, {
      headers: {
        'Origin': window.location.origin,
      }
    });

    if (!response.ok) {
      console.error(`Failed to fetch image: ${response.status}`);
      return null;
    }

    const blob = await response.blob();
    if (!blob.type.startsWith('image/')) {
      console.error("Response is not an image:", blob.type);
      return null;
    }

    return await new Promise((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
  } catch (error) {
    console.error("Failed to fetch image:", error);
    return null;
  }
}
async function renderAnswerHTML(answer: any): Promise<string> {
  if (answer === null || answer === undefined || answer === "") {
    return '<span style="color:#999;">-</span>';
  }

  if (typeof answer === "string") {
    if (isImageUrl(answer)) {
      // FETCH the image and convert to data URL
      const dataUrl = await fetchImageFromUrl(answer);
      if (dataUrl) {
        return `<img src="${dataUrl}" style="max-width:100%;max-height:120px;object-fit:contain;border-radius:3px;" alt="evidence" />`;
      }
      // Fallback to original URL if fetch fails
      return `<img src="${answer}" style="max-width:100%;max-height:120px;object-fit:contain;border-radius:3px;" alt="evidence" />`;
    }
    return escHtml(answer).replace(/\n/g, "<br>");
  }

  // Handle array answers
  if (Array.isArray(answer)) {
    const imageItems = answer.filter(item =>
      typeof item === "string" && isImageUrl(item)
    );

    if (imageItems.length > 0) {
      const processedImages = await Promise.all(imageItems.map(async (img) => {
        const dataUrl = await fetchImageFromUrl(img);
        const src = dataUrl || img;
        return `<img src="${src}" style="max-width:100%;max-height:85px;object-fit:contain;margin:2px;" alt=""/>`;
      }));
      return processedImages.join("");
    }

    // Handle array of objects
    const processedItems = await Promise.all(answer.map(async item => {
      if (typeof item === "object" && item !== null) {
        const imageFields = ['url', 'imageUrl', 'evidenceUrl', 'photo', 'image', 'src'];
        for (const field of imageFields) {
          if (item[field] && typeof item[field] === "string" && isImageUrl(item[field])) {
            const dataUrl = await fetchImageFromUrl(item[field]);
            const src = dataUrl || item[field];
            return `<img src="${src}" style="max-width:100%;max-height:85px;object-fit:contain;margin:2px;" alt=""/>`;
          }
        }
        return escHtml(JSON.stringify(item));
      }
      return escHtml(String(item));
    }));

    if (processedItems.length) return processedItems.join(", ");
    return "";
  }

  // Handle object answers
  if (typeof answer === "object") {
    const p: string[] = [];

    const imageFields = ['url', 'imageUrl', 'evidenceUrl', 'photo', 'image', 'src', 'dataUrl'];
    for (const field of imageFields) {
      if (answer[field] && typeof answer[field] === "string" && isImageUrl(answer[field])) {
        const dataUrl = await fetchImageFromUrl(answer[field]);
        const src = dataUrl || answer[field];
        p.push(`<img src="${src}" style="max-width:100%;max-height:85px;object-fit:contain;margin:2px;" alt=""/>`);
      }
    }

    if (answer.base64 || answer.data) {
      const imgData = answer.base64 || answer.data;
      if (typeof imgData === "string" && (imgData.startsWith('data:image') || imgData.length > 100)) {
        const imgSrc = imgData.startsWith('data:') ? imgData : `data:image/png;base64,${imgData}`;
        p.push(`<img src="${imgSrc}" style="max-width:100%;max-height:85px;object-fit:contain;margin:2px;" alt=""/>`);
      }
    }

    if (answer.status) {
      const c = answer.status.toLowerCase() === "accepted" ? "#16a34a"
        : answer.status.toLowerCase() === "rejected" ? "#dc2626" : "#d97706";
      p.push(`<b style="color:${c}">${escHtml(String(answer.status))}</b>`);
    }
    if (answer.chassisNumber) p.push(`<div><b>${escHtml(String(answer.chassisNumber))}</b></div>`);
    if (answer.remark) p.push(`<div>${escHtml(String(answer.remark))}</div>`);

    if (p.length) return p.join("");
    return escHtml(JSON.stringify(answer));
  }

  return escHtml(String(answer));
}


function groupResponsesByFormat(responses: any[], fQ: any): Map<string, any[]> {
  const g = new Map<string, any[]>();
  responses.forEach(r => {
    const key = fQ ? getAnswerString(r, fQ.id) : "NO_FORMAT";
    if (!g.has(key)) g.set(key, []);
    g.get(key)!.push(r);
  });
  return g;
}
// ============================================================
// MAIN HTML GENERATOR
// ============================================================
// ═══════════════════════════════════════════════════════════════
// TRANSLATION HELPER
// Uses MyMemory free translation API (no key needed, ~5000 chars/day)
// Falls back to LibreTranslate if MyMemory fails
// ═══════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════
// TRANSLATION HELPER
// Uses MyMemory free translation API (no key needed, ~5000 chars/day)
// Falls back to LibreTranslate if MyMemory fails
// ═══════════════════════════════════════════════════════════════

async function translateToHindi(text: string): Promise<string> {
  if (!text || text.trim() === "" || text === "&nbsp;") return "";

  // Strip HTML tags for translation, we'll re-wrap after
  const plainText = text.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  if (!plainText) return "";

  try {
    // MyMemory API — free, no key required
    const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(plainText)}&langpair=en|hi&de=appuprasanna460@gmail.com`;
    const res = await fetch(url);
    const data = await res.json();
    if (data?.responseStatus === 200 && data?.responseData?.translatedText) {
      return escHtml(data.responseData.translatedText);
    }
  } catch (_) {
    // silently fall through
  }

  // Fallback: return empty so row just shows English
  return "";
}

// Batch-translate an array of strings, preserving order
async function batchTranslate(texts: string[]): Promise<string[]> {
  return Promise.all(texts.map(t => translateToHindi(t)));
}

// ═══════════════════════════════════════════════════════════════
// COLUMNS THAT GET HINDI TRANSLATION (text-heavy only)
// idx matches colDefs index positions
// ═══════════════════════════════════════════════════════════════
// We use colDef array position (0-based) to decide which columns get Hindi translation:
// 0  = What/Activity       ✅
// 1  = Method (How)        ✅
// 2  = Frequency/When      ✅
// 3  = Standard            ✅
// 4  = Responsibility      ✅
// 5  = Equipment/Measuring ✅ (text answers expected in future)
// 6  = Possible Abnorm.    ✅
// 7  = Reaction Plan       ✅
// 8  = Part Name & QTY     skip (part codes)
// 9  = PPEs required       ✅ (text answers expected in future)
// 10 = Record / Document   ✅ (text answers expected in future)
// 11 = Remarks             ✅
const TRANSLATE_COL_POSITIONS = new Set([0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 11]);

// ═══════════════════════════════════════════════════════════════
// MAIN FUNCTION
// ═══════════════════════════════════════════════════════════════

// Helper to get question title by ID or index from questions array

function getQuestionLabel(question: any, fallback: string): string {
  if (!question) return fallback;
  // Question objects use 'text' property for the label
  return question.text || question.title || question.label || fallback;
}

async function generateOPSHTML(
  form: any, responses: any[],
  sectionMapping: {
    headerSectionId: string;
    generalInstructionsSectionId: string;
    pastProblemsSectionId: string;
    processStepsSectionId: string;
    associateSignSectionId: string;
    illustrationsSectionId: string;
  },
  formTitle: string
): Promise<string> {

  const findQ = (id: string) => {
    const s = form.sections?.find((x: any) => x.id === id);
    if (!s) return [];
    return (s.questions || []).filter((q: any) => !q.parentId && !q.showWhen?.questionId);
  };
  const hQ = findQ(sectionMapping.headerSectionId);
  const iQ = findQ(sectionMapping.generalInstructionsSectionId);
  const prQ = findQ(sectionMapping.processStepsSectionId);
  const aQ = findQ(sectionMapping.associateSignSectionId);
  const ilQ = findQ(sectionMapping.illustrationsSectionId);
  const pQ = findQ(sectionMapping.pastProblemsSectionId);
  console.log("=== DEBUG pQ ===");
  console.log("Looking for section ID:", sectionMapping.pastProblemsSectionId);
  console.log("pQ length:", pQ.length);
  pQ.forEach((q, idx) => {
    console.log(`pQ[${idx}]:`, q.id, q.text);
  });
  const valid = responses.filter(r => getResponseTimestamp(r));
  const first = valid[0] || null;

  // Find the image question once
  const imageQuestion = ilQ.find((q: any) => q.id === 'q_illustrations_images');

  const hdr = (i: number) => getCombinedAnswer(valid, hQ[i]);
  const ins = (i: number) => { const q = iQ[i]; return (first && q) ? getAnswerString(first, q.id) : ""; };
  const past = (i: number) => { const q = pQ[i]; return (first && q) ? getAnswerString(first, q.id) : ""; };
  const asc = (i: number) => { const q = aQ[i]; return (first && q) ? getAnswerString(first, q.id) : ""; };


  // Dynamic labels from form questions - using .text property
  // SECTION TITLES (from section objects)
  const illustrationsSection = form.sections?.find((s: any) => s.id === sectionMapping.illustrationsSectionId);
  const illustrationsHeader = illustrationsSection?.title || "Illustrations & Process Details";

  const processStepsSection = form.sections?.find((s: any) => s.id === sectionMapping.processStepsSectionId);
  const processStepsHeader = processStepsSection?.title || "Process Steps";

  // QUESTION TITLES (from question objects using .text property)
  const deptSectionTitle = getQuestionLabel(hQ[0], "Dept. / Section");
  const lineZoneTitle = getQuestionLabel(hQ[1], "Line / Zone");
  const modelTitle = getQuestionLabel(hQ[2], "Model");
  const processStationTitle = getQuestionLabel(hQ[3], "Process / Station");

  const formatNoTitle = getQuestionLabel(iQ[0], "Format No.");
  const controlNoTitle = getQuestionLabel(iQ[1], "Control No.");

  // Abnormality section question titles
  const abnormalityTitle = getQuestionLabel(pQ[0], "Abnormality handling route");
  const pastProblemTitle = getQuestionLabel(pQ[1], "Past Problem Details");

  // Associate section question titles
  const associateNameTitle = getQuestionLabel(aQ[0], "Associate Name & Emp. Code");
  const signDateTitle = getQuestionLabel(aQ[1], "Sign & Date");

  // Process step column headers
  const snHeader = getQuestionLabel(prQ[0], "SN");
  const importanceHeader = getQuestionLabel(prQ[1], "Item Importance");

  // Debug: Log question titles to see what's available
  console.log("=== QUESTION TITLES ===");
  console.log("hQ[0]:", hQ[0]);
  console.log("hQ[0] title:", hQ[0]?.title);
  console.log("prQ[0]:", prQ[0]);
  console.log("prQ[0] title:", prQ[0]?.title);
  console.log("iQ[0]:", iQ[0]);
  console.log("iQ[0] title:", iQ[0]?.title);

  // Load assets
  const [logoR, stopR, noSymR, noMobR, ppeGR, ppeGlR, fiveSR, qrR] = await Promise.all([
    fetchAssetBase64("/assets/Companylogo.png"),
    fetchAssetBase64("/assets/Safetyposter.png"),
    fetchAssetBase64("/assets/Dontrun.png"),
    fetchAssetBase64("/assets/dontusemobile.png"),
    fetchAssetBase64("/assets/PPEGuide.png"),
    fetchAssetBase64("/assets/PPEGUIDE2.png"),
    fetchAssetBase64("/assets/5S_Guidelines.png"),
    fetchAssetBase64("/assets/Qrcode.png"),
  ]);
  const logoSrc = base64ToDataUrl(logoR);
  const stopSrc = base64ToDataUrl(stopR);
  const noSymSrc = base64ToDataUrl(noSymR);
  const noMobSrc = base64ToDataUrl(noMobR);
  const ppeGSrc = base64ToDataUrl(ppeGR);
  const ppeGlSrc = base64ToDataUrl(ppeGlR);
  const fiveSSrc = base64ToDataUrl(fiveSR);
  const qrSrc = base64ToDataUrl(qrR);

  // Values
  const dept = hdr(0) || "-";
  const lineZone = hdr(1) || "-";
  const model = hdr(2) || "-";
  const station = hdr(3) || "-";
  const formatNo = ins(0) || "-";
  const controlNo = ins(1) || "-";
  const fifo = ins(4) || "1. Bin/trolley must be changed only after complete usage of all material in it.\n2. Empty bin/trolley should be replaced with new one.\n3. Don't top up partially filled bin.\n4. Follow FIFO on line during Process.\n5. Do not use next bin / Trolley material until running not consumed.";
  const nonLub = ins(5) || "Do not use any lubrication if not specified in OPS / Process Sheet.";
  const noMobTxt = ins(6) || "Do not use mobile on the shopfloor";
  const noRunTxt = ins(7) || "Do not run on the shopfloor";
  const envTxt = ins(8) || "1. Do waste segregation.\n2. Switch off idle lights & machines.\n3. Ensure 3R Principal in daily activities.\n4. If there was any leakage, communicate to Sub Leader.";
  const safeTxt = ins(9) || "1. Follow POS sheet in case of any Chemical.\n2. Follow MSDS/SDS in case of any emergency regarding chemical.\n3. Follow your PPE's.";

  // Shift timing inline table
  const shiftTimingHtml = `
    <table style="width:100%;border-collapse:collapse;font-size:5.5pt;">
      <tr style="background:#d9d9d9;">
        <th style="border:1px solid #aaa;padding:1px 3px;text-align:left;">Activity</th>
        <th style="border:1px solid #aaa;padding:1px 3px;text-align:center;">A</th>
        <th style="border:1px solid #aaa;padding:1px 3px;text-align:center;">B</th>
      </tr>
      <tr>
        <td style="border:1px solid #aaa;padding:1px 3px;">Shift Start</td>
        <td style="border:1px solid #aaa;padding:1px 3px;text-align:center;">06:00 AM</td>
        <td style="border:1px solid #aaa;padding:1px 3px;text-align:center;">02:50 PM</td>
      </tr>
      <tr>
        <td style="border:1px solid #aaa;padding:1px 3px;">Shift End</td>
        <td style="border:1px solid #aaa;padding:1px 3px;text-align:center;">02:50 PM</td>
        <td style="border:1px solid #aaa;padding:1px 3px;text-align:center;">11:40 PM</td>
      </tr>
      <tr>
        <td style="border:1px solid #aaa;padding:1px 3px;">Shift Start 3S &amp; Meeting</td>
        <td style="border:1px solid #aaa;padding:1px 3px;text-align:center;">06:00 AM to 06:10 AM</td>
        <td style="border:1px solid #aaa;padding:1px 3px;text-align:center;">02:50 PM to 03:00 PM</td>
       </tr>
      <tr>
        <td style="border:1px solid #aaa;padding:1px 3px;">Shift End 3S</td>
        <td style="border:1px solid #aaa;padding:1px 3px;text-align:center;">02:40 PM to 02:50 PM</td>
        <td style="border:1px solid #aaa;padding:1px 3px;text-align:center;">11:30 PM to 11:40 PM</td>
       </tr>
      <tr>
        <td colspan="3" style="border:1px solid #aaa;padding:1px 3px;font-size:5pt;font-style:italic;">
          Tea Break / Lunch / Dinner timings will be as per company timings.
        </td>
      </tr>
     </table>`;

  // Process instructions
  const defPI = [
    "2. Do Not Use Fallen Electrical/Functional Parts.",
    "3. Ensure Model / Variant Change.",
    "4. Report in case of part / hardware fallen inside vehicle.",
    "5. TQ Wrench Arrow Mark should be in correct direction.",
    "6. Put Fallen Hardware in Red Bin for Zone In-Charge judgement.",
    "7. Take approval from SH / HOD before changing process sequence.",
    "8. Zone In-Charge is overall responsible to ensure work is as per OPS.",
    "9. Contaminant parts should be covered properly."
  ];
  let procInsHtml = `<div style="margin-bottom:2px;">1. Do Exercise at Shift Start.</div>`;
  for (let i = 10; i <= 17; i++) {
    const v = ins(i);
    procInsHtml += `<div style="margin-bottom:2px;">${escHtml(v || defPI[i - 10] || "")}</div>`;
  }

  const assocName = asc(0) || "";
  const signDate = asc(1) || "";

  let pastProbHtml = "";
  if (pQ[0] && first) { pastProbHtml = await renderAnswerHTML(first.answers?.[pQ[0].id]); }

  // Process step columns definition
  const colDefs = [
    { idx: 2, label: "स्टेप<br>(What / Activity)", w: "8%" },
    { idx: 3, label: "Method<br>(How)", w: "6%" },
    { idx: 4, label: "Frequency<br>/ When", w: "4.5%" },
    { idx: 5, label: "Standard<br>(Spec./Judgment Criteria)", w: "9%" },
    { idx: 6, label: "Responsibility", w: "5%" },
    { idx: 7, label: "Equipment /<br>Measuring Eq.", w: "5.5%" },
    { idx: 8, label: "Possible<br>Abnormalities", w: "5.5%" },
    { idx: 9, label: "Reaction<br>Plan", w: "4.5%" },
    { idx: 10, label: "Part Name<br>&amp; QTY", w: "5%" },
    { idx: 11, label: "PPEs<br>required", w: "4%" },
    { idx: 12, label: "Record /<br>Document", w: "4.5%" },
    { idx: 13, label: "Remarks", w: "4%" },
  ];

  // ═══════════════════════════════════════════════════
  // BUILD PROCESS ROWS WITH TRANSLATION
  // Each row = top (English) + bottom (Hindi) split
  // ═══════════════════════════════════════════════════
  const actualMaxRows = Math.max(5, valid.length);

  // Pre-fetch all cell HTML for all rows in parallel
  const allRowCells: string[][] = [];
  for (let i = 0; i < actualMaxRows; i++) {
    const resp = valid[i];
    if (!resp) { allRowCells.push([]); continue; }
    const cells: string[] = [];
    for (const col of colDefs) {
      const q = prQ[col.idx];
      cells.push(await renderAnswerHTML(q ? resp.answers?.[q.id] : undefined));
    }
    allRowCells.push(cells);
  }

  // Now batch-translate text-heavy columns for all rows
  // Collect all texts that need translation
  interface TranslateJob { rowIdx: number; colPos: number; text: string; }
  const jobs: TranslateJob[] = [];
  for (let i = 0; i < actualMaxRows; i++) {
    const cells = allRowCells[i];
    if (!cells.length) continue;
    for (let colPos = 0; colPos < colDefs.length; colPos++) {
      if (TRANSLATE_COL_POSITIONS.has(colPos)) {
        jobs.push({ rowIdx: i, colPos, text: cells[colPos] });
      }
    }
  }

  // Translate all jobs in parallel (API calls go in parallel)
  const translatedTexts = await Promise.all(
    jobs.map(j => translateToHindi(j.text))
  );

  // Map translations back: translations[rowIdx][colPos]
  const translations: Map<string, string> = new Map();
  jobs.forEach((j, idx) => {
    translations.set(`${j.rowIdx}:${j.colPos}`, translatedTexts[idx]);
  });

  // Build row HTML
  const B1 = "border:1px solid #999;";
  let procRowsHtml = "";

  for (let i = 0; i < actualMaxRows; i++) {
    const resp = valid[i];
    const isLast = i === actualMaxRows - 1;
    const bb = isLast ? "2px solid #000" : "1px solid #999";
    const star = i === 0 ? "★" : i === 1 ? "★★" : i === 2 ? "★★★" : "";
    const cells = allRowCells[i];

    // Image for illustration column
    let illusHtml = "";
    if (resp && imageQuestion) {
      const imageAnswer = resp.answers?.['q_illustrations_images'];
      if (imageAnswer && typeof imageAnswer === "string" && imageAnswer.startsWith('http')) {
        illusHtml = `<img src="${imageAnswer}" style="max-width:85%;max-height:150px;object-fit:contain;" alt="Illustration" />`;
      } else if (imageAnswer) {
        illusHtml = `<span style="color:#666;">${escHtml(String(imageAnswer))}</span>`;
      }
    }

    if (!resp || !cells.length) {
      // Empty row — no translation needed
      procRowsHtml += `
        <tr style="height:56px;">
          <td style="${B1}border-bottom:${bb};background:#ffff00;">&nbsp;</td>
          <td style="${B1}border-bottom:${bb};text-align:center;vertical-align:middle;font-weight:700;font-size:8pt;">${i + 1}</td>
          <td style="${B1}border-bottom:${bb};text-align:center;vertical-align:middle;font-size:10pt;">☆</td>
          ${colDefs.map(() => `<td style="${B1}border-bottom:${bb};padding:2px;">&nbsp;</td>`).join("")}
        </tr>`;
      continue;
    }

    // Build each cell with English top / Hindi bottom split
    const cellsHtml = cells.map((englishContent, colPos) => {
      const w = colDefs[colPos].w;
      const hindiContent = translations.get(`${i}:${colPos}`) || "";
      const shouldTranslate = TRANSLATE_COL_POSITIONS.has(colPos);

      if (shouldTranslate && hindiContent) {
        // Split cell: English top, divider, Hindi bottom
        return `<td style="${B1}border-bottom:${bb};width:${w};vertical-align:top;padding:0;font-size:7pt;">
          <div style="padding:3px;border-bottom:1px dashed #bbb;min-height:30px;">
            ${englishContent}
          </div>
          <div style="padding:3px;min-height:22px;font-size:6pt;color:#1a1a8c;font-family:'Noto Sans Devanagari',Arial,sans-serif;">
            ${hindiContent}
          </div>
        </td>`;
      } else {
        // Single content cell (no translation or translation empty)
        return `<td style="${B1}border-bottom:${bb};width:${w};vertical-align:top;padding:3px;font-size:7pt;">
          ${englishContent}
        </td>`;
      }
    }).join("");

    procRowsHtml += `
      <tr style="min-height:56px;">
        <td style="${B1}border-bottom:${bb};background:#ffff00;vertical-align:top;padding:2px;font-size:6.5pt;">
          <div style="width:100%;overflow:hidden;">
            ${illusHtml}
          </div>
        </td>
        <td style="${B1}border-bottom:${bb};text-align:center;vertical-align:middle;font-weight:700;font-size:8pt;">${i + 1}</td>
        <td style="${B1}border-bottom:${bb};text-align:center;vertical-align:middle;font-size:10pt;">${star || "☆"}</td>
        ${cellsHtml}
      </tr>`;
  }

  // Associate sign columns
  const ACOLS = 22;

  // CSS shorthand
  const B2 = "border:2px solid #000;";
  const PAD = "padding:2px 3px;";
  const LBL = `${B1}${PAD}background:#e8e8e8;font-weight:700;vertical-align:middle;font-size:6.5pt;`;
  const VAL = `${B1}${PAD}background:#fff;font-size:7pt;vertical-align:middle;`;
  const HDR = `${B1}${PAD}background:#d9d9d9;font-weight:700;font-size:7pt;text-align:center;vertical-align:middle;`;

  return `<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8">
<title>${escHtml(formTitle || "OPS Report")}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Devanagari:wght@400;700&display=swap');
  @page { size: A3 landscape; margin: 2mm; margin-left: 4mm; margin-right: 4mm; }
  @media print { .no-print { display:none!important; } body { margin:0; } }
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:Arial,sans-serif; font-size:7pt; background:#fff; color:#000; }
  .page { width:100%; max-width:414mm; margin:0; }
  table { border-collapse:collapse; width:100%; table-layout:fixed; }
  td,th { vertical-align:middle; word-break:break-word; overflow:hidden; }
  .hindi { font-family:'Noto Sans Devanagari',Arial,sans-serif; color:#1a1a8c; font-size:6pt; }
  .en-text { font-size:7pt; }
  .cell-divider { border-bottom:1px dashed #bbb; }
  .print-btn { position:fixed; bottom:14px; right:14px; padding:8px 18px;
    background:#1d4ed8; color:#fff; border:none; border-radius:6px;
    cursor:pointer; font-size:12px; font-weight:700; z-index:9999;
    box-shadow:0 2px 8px rgba(0,0,0,.25); }
  .print-btn:hover { background:#1e40af; }
</style>
</head><body>
<button class="print-btn no-print" onclick="window.print()">📄 Save as PDF</button>
<div class="page">

<!-- RETENTION BAR -->
<table style="${B2}margin-bottom:0;">
  <tr>
    <td style="padding:1px 6px;text-align:right;font-weight:700;font-size:6.5pt;">
      Retention Period : 20 years after Model is discontinued
    </td>
  </tr>
</table>

<!-- TOP HEADER TABLE -->
<table style="${B2}border-top:none;">
  <colgroup>
    <col style="width:22mm"> <col style="width:13mm"> <col style="width:20mm"> <col style="width:13mm"> <col style="width:28mm">
    <col style="width:30mm"> <col style="width:7mm"> <col style="width:48mm"> <col style="width:38mm"> <col style="width:22mm">
    <col style="width:22mm"> <col style="width:22mm"> <col style="width:7mm"> <col style="width:14mm"> <col style="width:28mm"> <col style="width:26mm">
  </colgroup>

  <tr style="height:14px;">
    <td rowspan="8" style="${B2}text-align:center;vertical-align:middle;padding:2px;">
      ${logoSrc && logoSrc !== "data:," && logoSrc !== "data:image/png;base64,"
      ? `<img src="${logoSrc}" style="width:18mm; height:28mm; max-height:20mm; object-fit:contain;" alt="Logo" />`
      : `<span style="color:#fff;font-weight:700;font-size:18pt;">C</span>`}
    </td>
    <td style="${LBL}">${escHtml(deptSectionTitle)} :</td>
    <td style="${VAL}font-weight:700;">${escHtml(dept)}</td>
    <td style="${LBL}">${escHtml(lineZoneTitle)} :</td>
    <td style="${VAL}font-weight:700;">${escHtml(lineZone)}</td>
    <td colspan="4" style="${B2}text-align:center;vertical-align:middle;padding:3px;">
      <div style="font-size:18pt;font-weight:700;letter-spacing:1px;">Operation Standard</div>
    </td>
    <td rowspan="7" style="${B2}background:#fff;padding:0;vertical-align:top;"></td>
    <td rowspan="7" style="${B2}background:#fff;padding:0;vertical-align:top;"></td>
    <td rowspan="7" style="${B2}background:#fff;padding:0;vertical-align:top;"></td>
    <td rowspan="7" style="${B2}background:#fff;padding:0;vertical-align:top;">
      <table style="width:100%;height:100%;border-collapse:collapse;">${Array(12).fill('<tr><td style="border-bottom:1px solid #ccc;height:14px">&nbsp;</td></tr>').join('')}</table>
    </td>
    <td rowspan="7" style="${B2}background:#fff;padding:0;vertical-align:top;">
      <table style="width:100%;height:100%;border-collapse:collapse;">${Array(12).fill('<tr><td style="border-bottom:1px solid #ccc;height:14px">&nbsp;</td></tr>').join('')}</table>
    </td>
    <td rowspan="7" style="${B2}background:#fff;padding:0;vertical-align:top;">
      <table style="width:100%;height:100%;border-collapse:collapse;">${Array(12).fill('<tr><td style="border-bottom:1px solid #ccc;height:14px">&nbsp;</td></tr>').join('')}</table>
    </td>
    <td rowspan="8" style="${B2}vertical-align:top;padding:4px;font-size:6.5pt; height:180px;">
  <div style="font-weight:700;color:#c00;margin-bottom:2px;margin-top:5px;">${escHtml(formatNoTitle)} :</div>
  <div style="font-weight:700;font-size:7.5pt;margin-bottom:6px;">${escHtml(formatNo)}</div>
  <div style="border-top:1px solid #999; margin:4px 0 6px 0;"></div>
  <div style="font-weight:700;color:#c00;margin-bottom:2px;margin-top:5px;">${escHtml(controlNoTitle)} :</div>
  <div style="font-weight:700;font-size:7.5pt;margin-bottom:6px;">${escHtml(controlNo)}</div>
  <div style="border-top:1px solid #999; margin:4px 0 6px 0;"></div>
  <div style="font-weight:700;margin-bottom:2px;">QR Code :</div>
  ${qrSrc ? `<img src="${qrSrc}" style="width:28mm;height:18mm;object-fit:contain;" alt="QR"/>` : `<div style="width:22mm;height:16mm;background:#000;"></div>`}
</td>
  </tr>

  <tr style="height:13px;">
      <td style="${LBL} padding:8px 3px;">${escHtml(modelTitle)} :</td>
  <td style="${VAL}font-weight:700; padding:8px 3px;">${escHtml(model)}</td>
  <td style="${LBL} padding:8px 3px;">${escHtml(processStationTitle)} :</td>
  <td style="${VAL}font-weight:700; padding:8px 3px;">${escHtml(station)}</td>
    <td colspan="4" style="${HDR}font-size:7pt;">
      Your Work When Trouble Stopped The Production Line
    </td>
  </tr>

  <tr style="height:13px;">
    <td rowspan="6" colspan="2" style="${B2}padding:2px 3px;vertical-align:top;font-size:6pt;">
      <div style="font-weight:700;margin-bottom:2px;">REJECTION HANDLING :-</div>
      Clearly Identify Rejected / NG parts. Keep them properly with proper identification at defined Location.
    </td>
    <td rowspan="6" style="${B2}padding:2px;text-align:center;vertical-align:middle;font-weight:700;font-size:6.5pt;">
      Measuring<br>Instruments<br>or Gauges
    </td>
    <td rowspan="6" style="${B2}padding:0;vertical-align:top;font-size:6pt;">
      <div style="padding:2px 3px;border-bottom:1px solid #bbb;">Always use Calibrated Measuring Instruments / Gauges (Ensure Calibration status before using the instrument).</div>
      <div style="padding:2px 3px;border-bottom:1px solid #bbb;">Ensure Zero setting before use.</div>
      <div style="padding:2px 3px;border-bottom:1px solid #bbb;">Do Not Use Unidentified Measuring Tool / Gauges.</div>
      <div style="padding:2px 3px;">In case of any abnormality, inform Line leader and Quality Engineer to take action for suspected NG parts.</div>
    </td>
    <td rowspan="6" style="${B2}text-align:center;vertical-align:middle;padding:2px;">
      ${stopSrc
      ? `<img src="${stopSrc}" style="max-width:28mm;max-height:28mm;object-fit:contain;" alt="Stop Call Wait"/>`
      : `<div style="font-size:8pt;font-weight:700;color:red;border:2px solid red;padding:6px;text-align:center;">STOP<br>CALL<br>WAIT</div>`}
    </td>
    <td style="${HDR}">S.<br>No.</td>
    <td style="${HDR}">Trouble</td>
    <td style="${HDR}">Your task</td>
  </tr>

  <tr style="height:13px;">
    <td style="${B1}${PAD}text-align:center;font-size:7pt;">1</td>
    <td style="${B1}${PAD}font-size:6.5pt;">Equipment Trouble / Machine Break Down</td>
    <td rowspan="5" style="${B1}${PAD}font-size:6.5pt;vertical-align:middle;text-align:center;">
      Stop The Line<br>
      Inform the Zone Leader<br>
      Write on card if mentioned in OPS
    </td>
  </tr>
  <tr style="height:13px;">
    <td style="${B1}${PAD}text-align:center;font-size:7pt;">2</td>
    <td style="${B1}${PAD}font-size:6.5pt;">A Trouble You Are Responsible For</td>
  </tr>
  <tr style="height:13px;">
    <td style="${B1}${PAD}text-align:center;font-size:7pt;">3</td>
    <td style="${B1}${PAD}font-size:6.5pt;">Empty Marshal Carrier</td>
  </tr>
  <tr style="height:13px;">
    <td style="${B1}${PAD}text-align:center;font-size:7pt;">4</td>
    <td style="${B1}${PAD}font-size:6.5pt;">Stock Out / Material Shortage</td>
  </tr>
  <tr>
    <td style="${B1}${PAD}text-align:center;font-size:7pt;">5</td>
    <td style="${B1}${PAD}font-size:6.5pt;">A Trouble From Different Section</td>
    <td style="${HDR}font-size:6.5pt;font-weight:700; height:30px">Prepared</td>
    <td style="${HDR}font-size:6.5pt;font-weight:700; height:22px">Checked</td>
    <td style="${HDR}font-size:6.5pt;font-weight:700; height:22px">Approved</td>
    <td style="${HDR}font-size:6.5pt; ">No.</td>
    <td style="${HDR}font-size:6.5pt; ">DD /MM/ YY</td>
    <td style="${HDR}font-size:6.5pt;">Issuance / Revision details</td>
  </tr>
</table>

<!-- GENERAL INSTRUCTIONS -->
<table style="${B2}border-top:none;">
  <tr>
    <td style="padding:2px 6px;font-weight:700;font-size:8pt;text-align:center;background:#d9d9d9;">
      General Instructions
    </td>
  </tr>
</table>

<table style="${B2}border-top:none;table-layout:fixed;">
  <colgroup>
    <col style="width:11%">
    <col style="width:5%">
    <col style="width:5%">
    <col style="width:6%">
    <col style="width:6%">
    <col style="width:14%">
    <col style="width:7%">
    <col style="width:7%">
    <col style="width:16%">
    <col style="width:14%">
  </colgroup>
  <tr>
    <td style="${HDR}">FIFO System</td>
    <td colspan="2" style="${HDR}font-size:7pt;font-weight:700;">Non Lubrication Rule:</td>
    <td style="${HDR}font-size:6pt;">Always wear PPEs /<br>Proper uniform</td>
    <td style="${HDR}font-size:6pt;">Wear PPEs as per your<br>station's requirements</td>
    <td style="${HDR}">Shift Timings</td>
    <td colspan="2" style="${HDR}">EMS &amp; Safety Guidelines</td>
    <td style="${HDR}">5S Guidelines</td>
    <td style="${HDR}">Process Instructions</td>
  </tr>
  <tr>
    <td rowspan="99" style="${B1}padding:4px;vertical-align:top;font-size:6.5pt;">
      <div style="font-weight:700;margin-bottom:3px;">FIFO System</div>
      ${escHtml(fifo).replace(/\n/g, "<br>")}
    </td>
    <td colspan="2" rowspan="99" style="${B1}padding:0;vertical-align:top;font-size:6.5pt;">
      <table style="width:100%;height:100%;border-collapse:collapse;">
        <tr>
          <td colspan="2" style="border-bottom:1px solid #999;padding:4px;text-align:center;font-size:6.5pt;">
            ${escHtml(nonLub)}
          </td>
        </tr>
        <tr>
          <td style="border-bottom:1px solid #999;border-right:1px solid #999;padding:3px;text-align:center;font-weight:700;font-size:6pt;background:#d9d9d9;">
            Do not use mobile on the shopfloor
          </td>
          <td style="border-bottom:1px solid #999;padding:3px;text-align:center;font-weight:700;font-size:6pt;background:#d9d9d9;">
            Do not run on the shopfloor
          </td>
        </tr>
        <tr>
          <td style="border-right:1px solid #999;padding:4px;text-align:center;vertical-align:middle;">
            ${noMobSrc ? `<img src="${noMobSrc}" style="max-width:100%;max-height:75px;object-fit:contain;" alt="No Mobile"/>` : ""}
          </td>
          <td style="padding:4px;text-align:center;vertical-align:middle;">
            ${noSymSrc ? `<img src="${noSymSrc}" style="max-width:100%;max-height:75px;object-fit:contain;" alt="No Run"/>` : ""}
          </td>
        </tr>
      </table>
    </td>
    <td rowspan="99" style="${B1}padding:3px;vertical-align:top;text-align:center;font-size:6pt;">
      ${ppeGSrc ? `<img src="${ppeGSrc}" style="width:100%;max-height:130px;object-fit:contain;display:block;margin-top:14px;" alt="PPE Uniform"/>` : ""}
    </td>
    <td rowspan="99" style="${B1}padding:3px;vertical-align:top;text-align:center;font-size:6pt;">
      ${ppeGlSrc ? `<img src="${ppeGlSrc}" style="width:100%;max-height:130px;object-fit:contain;display:block;margin-top:14px;" alt="PPE Items"/>` : ""}
    </td>
    <td rowspan="99" style="${B1}padding:3px;vertical-align:top;margin-top:10px;">
      ${shiftTimingHtml}
    </td>
    <td style="${HDR}font-size:6pt;">Environmental Issues</td>
    <td style="${HDR}font-size:6pt;">Safety Issues</td>
    <td rowspan="99" style="${B1}padding:0;vertical-align:top;text-align:center;">
      ${fiveSSrc ? `<img src="${fiveSSrc}" style="width:100%;height:70%;max-height:180px;object-fit:fill;display:block;" alt="5S Guidelines"/>` : ""}
    </td>
    <td rowspan="99" style="${B1}padding:4px;vertical-align:top;font-size:6.5pt;">
      ${procInsHtml}
    </td>
  </tr>
  <tr>
    <td style="${B1}padding:2px;vertical-align:top;font-size:6.5pt;">
      ${escHtml(envTxt).replace(/\n/g, "<br>")}
    </td>
    <td style="${B1}padding:2px;vertical-align:top;font-size:6.5pt;">
      ${escHtml(safeTxt).replace(/\n/g, "<br>")}
    </td>
  </tr>
</table>


<!-- PROCESS STEPS — with English/Hindi split rows -->
<table style="${B2}border-top:none;table-layout:fixed;">
  <colgroup>
    <col style="width:9%">
    <col style="width:2.5%">
    <col style="width:4%">
    <col style="width:8%">
    <col style="width:6%">
    <col style="width:5.5%">
    <col style="width:9%">
    <col style="width:5%">
    <col style="width:6.5%">
    <col style="width:5.5%">
    <col style="width:4.5%">
    <col style="width:5%">
    <col style="width:4%">
    <col style="width:4.5%">
    <col style="width:4%">
  </colgroup>
  <thead>
    <tr>
      <th style="${HDR}background:#ffff00;">${escHtml(illustrationsHeader)}</th>
<th style="${HDR}">${escHtml(snHeader)}</th>
<th style="${HDR}">${escHtml(importanceHeader)}</th>
      ${colDefs.map(c => `<th style="${HDR}">${c.label}</th>`).join("")}
    </tr>
  </thead>
  <tbody>
    ${procRowsHtml}
  </tbody>
</table>

<!-- ABNORMALITY + PAST PROBLEMS -->
<table style="${B2}border-top:none;">
  <colgroup>
    <col style="width:14%">
    <col style="width:86%">
  </colgroup>
  <tr>
    <td rowspan="2" style="${B1}padding:4px;vertical-align:top;font-size:6.5pt;">
      <div style="font-weight:700;margin-bottom:2px;">${escHtml(abnormalityTitle)} :</div>
      In case of any abnormality inform the Zone In-Charge<br>
      Flow of Communication :-<br>
      Operator &#9658; Team Member &#9658; Section Mgr &#9658; As required
    </td>
   <td style="${B1}padding:2px;font-weight:700;text-align:center;background:#e8e8e8;font-size:7pt;">
    ${escHtml(pastProblemTitle)}
  </td>
  </tr>
  <tr>
    <td style="${B1}padding:4px;vertical-align:top;height:60px;font-size:7pt;">
      ${pastProbHtml}
    </td>
  </tr>
</table>

<!-- ASSOCIATE NAME & SIGN -->
<table style="${B2}border-top:none;table-layout:fixed;">
  <colgroup>
    <col style="width:5%">
    ${Array.from({ length: ACOLS }, () => `<col style="width:${(95 / ACOLS).toFixed(2)}%">`).join("")}
  </colgroup>
  <tr>
    <td style="${B1}padding:3px;font-weight:700;text-align:center;background:#e8e8e8;font-size:6.5pt;vertical-align:middle;">
      Associate Name<br>&amp; Emp. Code
    </td>
    ${Array.from({ length: ACOLS }, () =>
        `<td style="${B1}padding:2px;text-align:center;font-size:7pt;height:22px;"></td>`
      ).join("")}
  </tr>
  <tr>
    <td style="${B1}padding:3px;font-weight:700;text-align:center;background:#e8e8e8;font-size:6.5pt;vertical-align:middle;">
      Sign &amp; Date
    </td>
    ${Array.from({ length: ACOLS }, () =>
        `<td style="${B1}padding:2px;text-align:center;font-size:7pt;height:26px;"></td>`
      ).join("")}
  </tr>
</table>

<!-- PAGE NUMBER -->
<table style="${B2}border-top:none;">
  <colgroup>
    <col style="width:82%">
    <col style="width:18%">
  </colgroup>
  <tr>
    <td style="${B1}padding:2px;">&nbsp;</td>
    <td style="${B1}padding:3px;font-weight:700;font-size:8.5pt;text-align:center;">Page Number : XX / XX</td>
  </tr>
</table>

</div>
</body></html>`;
}
// ============================================================
// MAIN EXPORT (unchanged)
// ============================================================
export async function exportResponsesToOPSPDF(
  form: any, responses: any[],
  sectionMapping: {
    headerSectionId: 'sec_basic_info';
    generalInstructionsSectionId: 'sec_doc_control';
    pastProblemsSectionId: 'sec_abnormalities_handling';
    processStepsSectionId: 'sec_process_steps';
    associateSignSectionId: 'sec_associate_sign';
    illustrationsSectionId: 'sec_illustrations',
  },
  formTitle?: string,
  onProgress?: (message: string) => void
): Promise<void> {
  const findSQ = (id: string) => {
    const s = form.sections?.find((x: any) => x.id === id);
    if (!s) return [];
    return (s.questions || []).filter((q: any) => !q.parentId && !q.showWhen?.questionId);
  };
  const iQ = findSQ(sectionMapping.generalInstructionsSectionId);
  const fQ = iQ[0], cQ = iQ[1];

  const valid = responses.filter(r => getResponseTimestamp(r));
  if (!valid.length) throw new Error("No valid responses to export");

  let allSame = true, firstFmt = "", firstCtl = "";
  for (const r of valid) {
    const fmt = fQ ? getAnswerString(r, fQ.id) : "";
    const ctl = cQ ? getAnswerString(r, cQ.id) : "";
    if (!firstFmt && !firstCtl) { firstFmt = fmt; firstCtl = ctl; }
    else if (fmt !== firstFmt || ctl !== firstCtl) { allSame = false; break; }
  }

  const title = formTitle || form?.title || "OPS_Report";
  const dateStr = new Date().toLocaleDateString("en-CA").replace(/\//g, "-");

  if (allSame) {
    onProgress?.("Generating PDF...");
    const html = await generateOPSHTML(form, valid, sectionMapping, title);
    const blob = await apiClient.generateOPSPDF({
      htmlContent: html,
      filename: `${title}_${firstFmt}_${firstCtl}_${dateStr}.pdf`,
      compressed: html.length > 500000
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `${title}_${firstFmt}_${firstCtl}_${dateStr}.pdf`; a.click();
    URL.revokeObjectURL(url);
    onProgress?.("PDF downloaded successfully!");
  } else {
    onProgress?.("Multiple Format/Control combinations — generating ZIP...");
    const groups = groupResponsesByFormat(valid, fQ);
    const zip = new JSZip(); let idx = 0;
    for (const [key, gr] of groups.entries()) {
      idx++; onProgress?.(`Processing ${idx}/${groups.size}...`);
      const [fmt, ctl] = key.split("|");
      const html = await generateOPSHTML(form, gr, sectionMapping, title);
      const blob = await apiClient.generateOPSPDF({
        htmlContent: html,
        filename: `${fmt || "OPS"}${ctl ? "_" + ctl : ""}.pdf`,
        compressed: html.length > 500000
      });
      zip.file(
        `${(fmt || "OPS").replace(/[^a-zA-Z0-9_-]/g, "_")}${ctl ? "_" + ctl.replace(/[^a-zA-Z0-9_-]/g, "_") : ""}.pdf`,
        blob
      );
    }
    const zb = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(zb);
    const a = document.createElement("a");
    a.href = url; a.download = `${title}_${dateStr}.zip`; a.click();
    URL.revokeObjectURL(url);
    onProgress?.(`ZIP with ${groups.size} PDFs downloaded!`);
  }
}