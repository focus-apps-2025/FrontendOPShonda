/**
 * opsExcelGenerator.ts
 *
 * Generates OPS (Operation Standard) Excel files from form data,
 * including embedded images from S3 URLs.
 *
 * Uses xlsx-js-style for styling and image embedding.
 */

import * as XLSX from 'xlsx';
import XLSX_STYLE from 'xlsx-js-style';
import type { OpsTemplateConfig } from '../types/forms';

const { utils: styleUtils, writeFile } = XLSX_STYLE;
const { utils, read } = XLSX;

interface OpsFormImages {
  logoBImageBase64?: string;
  stopCallWaitImageBase64?: string;
  noSymbolImageBase64?: string;
  ppeGuideImageBase64?: string;
  fiveSImageBase64?: string;
  qrCodeImageBase64?: string;
}

interface ProcessStepData {
  sn: string;
  itemImportance: string;
  stepWhat: string;
  methodHow: string;
  frequencyWhen: string;
  standardCriteria: string;
  responsibility: string;
  equipmentMeasuring: string;
  possibleAbnormalities: string;
  reactionPlan: string;
  partName: string;
  partQty: string;
  ppeRequired: string;
  recordDocument: string;
  remarks: string;
}

function col(letter: string): number {
  return utils.decode_col(letter);
}

function cellStyle(
  bold = false,
  fontSize = 11,
  hAlign: string = 'left',
  vAlign: string = 'center',
  wrapText = false,
  top: string | null = null,
  bottom: string | null = null,
  left: string | null = null,
  right: string | null = null,
  fillColor: string | null = null
) {
  const border: Record<string, any> = {};
  if (top) border.top = { style: top };
  if (bottom) border.bottom = { style: bottom };
  if (left) border.left = { style: left };
  if (right) border.right = { style: right };
  return {
    font: { name: 'Arial', bold, sz: fontSize },
    alignment: { horizontal: hAlign as any, vertical: vAlign, wrapText },
    border,
    ...(fillColor ? { fill: { patternType: 'solid', fgColor: { rgb: fillColor } } } : {}),
  };
}

function mergeAndSet(
  ws: Record<string, any>,
  range: string,
  value: any,
  style: object
) {
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push(utils.decode_range(range));
  const topLeft = range.split(':')[0];
  if (!ws[topLeft]) ws[topLeft] = {};
  ws[topLeft].v = value;
  ws[topLeft].t = typeof value === 'number' ? 'n' : 's';
  ws[topLeft].s = style;
}

function setCell(ws: Record<string, any>, cellAddr: string, value: any, style: object) {
  if (!ws[cellAddr]) ws[cellAddr] = {};
  ws[cellAddr].v = value;
  ws[cellAddr].t = typeof value === 'number' ? 'n' : 's';
  ws[cellAddr].s = style;
}

function addImage(
  ws: Record<string, any>,
  base64: string,
  type: 'jpeg' | 'png' | 'gif',
  colStart: number,
  rowStart: number,
  colEnd: number,
  rowEnd: number
) {
  if (!ws['!images']) ws['!images'] = [];
  ws['!images'].push({
    '!pos': { r: rowStart, c: colStart, R: rowEnd, C: colEnd },
    '!datatype': 'base64',
    '!type': type,
    '!data': base64,
  });
}

function toRawBase64(dataUrl: string): string {
  if (dataUrl && dataUrl.startsWith('data:')) {
    return dataUrl.split(',')[1] || '';
  }
  return dataUrl || '';
}

function fillBorderRange(
  ws: Record<string, any>,
  startCol: number,
  endCol: number,
  startRow: number,
  endRow: number,
  outerTop: string | null,
  outerBottom: string | null,
  outerLeft: string | null,
  outerRight: string | null,
  innerH: string | null = null,
  innerV: string | null = null
) {
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const t = r === startRow ? outerTop : innerH;
      const b = r === endRow ? outerBottom : innerH;
      const l = c === startCol ? outerLeft : innerV;
      const ri = c === endCol ? outerRight : innerV;
      const addr = utils.encode_cell({ c, r });
      if (!ws[addr]) ws[addr] = { v: '', t: 's' };
      const border: Record<string, any> = {};
      if (t) border.top = { style: t };
      if (b) border.bottom = { style: b };
      if (l) border.left = { style: l };
      if (ri) border.right = { style: ri };
      ws[addr].s = { border };
    }
  }
}

/**
 * Generate the OPS Format Excel file from form data
 */
export function generateOPSExcel(
  formTitle: string,
  basicInfo: Record<string, string>,
  processSteps: ProcessStepData[],
  abnormalityHandling: string,
  abnormalityDetails: string,
  images: OpsFormImages = {},
  templateConfig?: OpsTemplateConfig
): ArrayBuffer {
  const ws: Record<string, any> = {};
  ws['!ref'] = 'B1:CM77';
  ws['!merges'] = [];

  // Row heights
  ws['!rows'] = [];
  const rowHeights: Record<number, number> = {
    0: 27, 1: 39, 2: 39, 3: 39, 4: 39, 5: 39, 6: 39,
    7: 45, 8: 45, 9: 45, 10: 45,
    11: 33.75, 12: 37.5, 13: 37.5, 14: 37.5, 15: 47.25,
    16: 37.5, 17: 37.5, 18: 37.5, 19: 37.5, 20: 37.5,
    21: 15, 22: 54.75, 23: 54.75,
    64: 15,
    65: 25, 66: 46.5, 67: 47.25, 68: 29.25, 69: 14.25,
    70: 29.25, 71: 29.25, 72: 29.25,
    73: 29.25, 74: 29.25, 75: 29.25,
    76: 39,
  };
  for (let i = 24; i <= 63; i++) rowHeights[i] = 62.15;
  for (const [r, h] of Object.entries(rowHeights)) {
    ws['!rows'][Number(r)] = { hpt: h };
  }

  // Column widths
  const colWidthMap: Record<string, number> = {
    A: 15.27, B: 4.54, D: 5.73, E: 22.73, F: 4.54, G: 8.0,
    I: 7.73, J: 8.73, M: 4.54, O: 9.27, P: 7.54, Q: 6.45,
    R: 13.18, S: 11.82, T: 13.82, U: 10.27, V: 11.54, W: 12.27,
    X: 4.54, AD: 8.18, AE: 7.18, AF: 7.45, AG: 4.54, AH: 8.0,
    AI: 6.82, AJ: 9.27, AK: 14.82, AS: 4.27, AX: 9.45, AY: 23.0,
    BH: 30.82, BI: 8.45, BJ: 22.54, BK: 23.54, BL: 7.27,
    BM: 6.73, BN: 8.45, BP: 13.45, BQ: 11.0, BR: 15.18,
    BS: 17.0, BT: 10.82, BU: 5.18, BW: 11.82, BX: 5.18,
    BZ: 10.18, CA: 5.82, CB: 3.82, CD: 6.73, CE: 11.0,
    CG: 6.45, CI: 5.73, CJ: 9.54, CK: 9.27, CL: 12.0,
    CM: 5.45,
  };
  ws['!cols'] = [];
  for (const [letter, width] of Object.entries(colWidthMap)) {
    ws['!cols'][(col as any)(letter)] = { wch: width };
  }

  // ── ROW 1
  mergeAndSet(ws, 'B1:CM1', formTitle || 'Operation Standard',
    cellStyle(true, 18, 'right', 'center', true, 'medium', null, 'medium', 'medium'));

  // ── ROWS 2-11: Top Header
  fillBorderRange(ws, col('B'), col('E'), 1, 10, 'medium', 'medium', 'medium', 'medium');
  mergeAndSet(ws, 'F2:I4', templateConfig?.basicInfoLabels?.deptSection || 'Dept. / Section :',
    cellStyle(true, 20, 'left', 'center', true, 'medium', 'thin', null, null));
  mergeAndSet(ws, 'J2:L4', basicInfo.deptSection || 'AF / Flexi Line 1.0',
    cellStyle(true, 20, 'center', 'center', false, 'medium', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'M2:R4', templateConfig?.basicInfoLabels?.lineZone || 'Line / Zone :',
    cellStyle(true, 20, 'left', 'center', false, 'medium', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'S2:W4', basicInfo.lineZone || 'MAIN LINE 2',
    cellStyle(true, 20, 'center', 'center', false, 'medium', 'thin', 'thin', 'medium'));
  mergeAndSet(ws, 'X2:BA4', 'Operation Standard ',
    cellStyle(true, 72, 'center', 'center', false, 'medium', 'medium', 'medium', null));
  mergeAndSet(ws, 'F5:I7', templateConfig?.basicInfoLabels?.model || 'Model :',
    cellStyle(true, 20, 'left', 'center', true, 'thin', 'thin', 'medium', 'thin'));
  mergeAndSet(ws, 'J5:L7', basicInfo.model || 'MLWA',
    cellStyle(true, 20, 'center', 'center', true, 'thin', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'M5:R7', templateConfig?.basicInfoLabels?.processStation || 'Process / Station :',
    cellStyle(true, 20, 'left', 'center', true, 'thin', 'medium', 'thin', 'thin'));
  setCell(ws, 'S5', basicInfo.processStation || '',
    cellStyle(true, 20, 'left', 'center', true, 'thin', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'AH5:BA5', 'Your Work When Trouble Stopped The Production Line',
    cellStyle(true, 20, 'center', 'center', false, 'medium', 'thin', 'thin', null));

  const defaultTroubles = [
    [7, 1, 'Equipment Trouble / Machine Break Down',
      'Stop The Line\\nInform the Zone Leader\\nWrite on card if mentioned in OPS'],
    [8, 2, 'A Trouble You Are Responsible For', ''],
    [9, 3, 'Empty Marshal Carrier ', ''],
    [10, 4, 'Stock Out / Material Shortage ', ''],
    [11, 5, 'A Trouble From Different Section', ''],
  ];

  const troublesToUse = templateConfig?.troubleTasks?.map(t => [t.sno + 6, t.sno, t.trouble, t.task]) || defaultTroubles;

  for (const [r, sno, trouble, task] of troublesToUse) {
    mergeAndSet(ws, `AH${r}:AI${r}`, sno,
      cellStyle(true, 20, 'center', 'center', false, 'thin', 'thin', 'thin', 'thin'));
    mergeAndSet(ws, `AJ${r}:AR${r}`, trouble,
      cellStyle(false, 20, 'left', 'center', true, 'thin', 'thin', 'thin', 'thin'));
    if (task !== undefined) {
      mergeAndSet(ws, `AS${r}:BA${r}`, task,
        cellStyle(false, 20, 'center', 'center', true, 'thin', r === 7 ? null : 'thin', 'thin', null));
    }
  }

  mergeAndSet(ws, 'F8:J11',
    templateConfig?.rejectionHandling || 'REJECTION HANDLING :-\\n\\nClearly Identify Rejected / NG parts. Keep them properly with proper identification at defined Location.',
    cellStyle(false, 18, 'left', 'center', true, 'medium', 'medium', 'medium', 'medium'));
  mergeAndSet(ws, 'K8:O11', 'Measuring Instruments or Gauges ',
    cellStyle(true, 18, 'center', 'center', true, 'medium', 'medium', null, 'thin'));

  const measuringInstruments = templateConfig?.measuringInstruments || [
    'Always use Calibrated Measuring Instruments / Gauges (Ensure Calibration status before using the same).',
    'Ensure Zero setting before use.',
    'Do Not Use Unidentified Measuring Tool / Gauges.',
    'In case of any abnormality, inform Line leader and Quality Engineer to take action for suspected NG material range.'
  ];

  mergeAndSet(ws, 'P8:W8', measuringInstruments[0] || '',
    cellStyle(false, 18, 'left', 'center', true, 'medium', 'thin', 'thin', 'medium'));
  mergeAndSet(ws, 'P9:W9', measuringInstruments[1] || '',
    cellStyle(false, 18, 'left', 'center', true, 'thin', 'thin', 'thin', 'medium'));
  mergeAndSet(ws, 'P10:W10', measuringInstruments[2] || '',
    cellStyle(false, 18, 'left', 'center', true, 'thin', 'thin', 'thin', 'medium'));
  mergeAndSet(ws, 'P11:W11', measuringInstruments[3] || '',
    cellStyle(false, 18, 'left', 'center', true, 'thin', 'medium', 'thin', 'medium'));

  fillBorderRange(ws, col('BB'), col('BE'), 1, 7, 'medium', 'medium', 'medium', 'medium');
  fillBorderRange(ws, col('BF'), col('BI'), 1, 7, 'medium', 'medium', 'medium', 'medium');
  fillBorderRange(ws, col('BJ'), col('BL'), 1, 7, 'medium', 'medium', 'medium', 'medium');
  mergeAndSet(ws, 'BB9:BE11', 'Prepared',
    cellStyle(true, 28, 'center', 'center', false, null, 'medium', 'medium', 'medium'));
  mergeAndSet(ws, 'BF9:BI11', 'Checked',
    cellStyle(true, 28, 'center', 'center', false, null, 'medium', 'medium', 'medium'));
  mergeAndSet(ws, 'BJ9:BL11', 'Approved',
    cellStyle(true, 28, 'center', 'center', false, null, 'medium', null, 'medium'));

  for (let r = 2; r <= 11; r++) {
    const topB = r === 2 ? 'medium' : 'thin';
    const botB = r === 11 ? null : 'thin';
    mergeAndSet(ws, `BM${r}:BN${r}`, r === 11 ? 'No.' : '',
      cellStyle(r === 11, 20, 'center', 'center', false, topB, botB, 'medium', 'thin'));
    mergeAndSet(ws, `BO${r}:BS${r}`, r === 11 ? 'DD /MM/ YY' : '',
      cellStyle(r === 11, 20, 'center', 'center', false, topB, r === 11 ? 'medium' : botB, 'thin', 'thin'));
    mergeAndSet(ws, `BT${r}:CD${r}`, r === 11 ? 'Issuance / Revision details' : '',
      cellStyle(r === 11, 20, 'center', 'center', false, topB, botB, 'thin', 'thin'));
  }

  mergeAndSet(ws, 'CE2:CM4', (templateConfig?.basicInfoLabels?.formatNo || 'Format No. : ') + (basicInfo.formatNo || '07010-QMHO-F0-171'),
    cellStyle(true, 20, 'left', 'center', false, 'medium', 'thin', 'thin', 'medium'));
  mergeAndSet(ws, 'CE5:CM8', (templateConfig?.basicInfoLabels?.controlNo || 'Control No. : ') + (basicInfo.controlNo || '07010-AF1F-Z1-9161'),
    cellStyle(true, 20, 'left', 'center', false, 'thin', 'thin', 'thin', 'medium'));
  mergeAndSet(ws, 'CE9:CM11', 'QR Code :',
    cellStyle(true, 20, 'left', 'center', false, 'thin', null, 'thin', 'medium'));

  // ── ROW 12: Section headers
  mergeAndSet(ws, 'B12:BC12', 'General Instructions',
    cellStyle(true, 20, 'center', 'center', true, 'medium', 'thin', 'medium', 'thin'));
  mergeAndSet(ws, 'BD12:BK12', 'EMS & Safety Guidelines',
    cellStyle(true, 20, 'center', 'center', false, 'medium', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'BL12:BW12', '5S Guidelines',
    cellStyle(true, 20, 'center', 'center', true, 'medium', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'BX12:CM12', 'Process Instructions ',
    cellStyle(true, 20, 'center', 'center', true, 'medium', 'thin', 'medium', 'thin'));

  mergeAndSet(ws, 'B13:L13', 'FIFO System',
    cellStyle(true, 20, 'center', 'center', true, 'thin', 'thin', 'medium', 'thin'));
  mergeAndSet(ws, 'M13:U13', 'Non Lubrication Rule: ',
    cellStyle(true, 20, 'center', 'center', true, 'thin', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'V13:AD13', 'Always wear PPEs / Proper uniform',
    cellStyle(true, 20, 'center', 'center', false, 'thin', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'AE13:AL13', "Wear PPEs as per your station's requirements",
    cellStyle(true, 20, 'center', 'center', false, 'thin', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'AM13:BC13', 'Shift Timings',
    cellStyle(true, 20, 'center', 'center', false, 'thin', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'BD13:BG13', 'Environmental Issues',
    cellStyle(true, 20, 'center', 'center', false, 'thin', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'BH13:BK13', 'Safety Issues',
    cellStyle(true, 20, 'center', 'center', false, 'thin', 'thin', 'thin', 'thin'));
  mergeAndSet(ws, 'BX13:CM13', '1. Do Exercise at Shift Start.',
    cellStyle(false, 20, 'left', 'center', true, null, 'thin', 'medium', 'thin'));

  const genInstr = [
    ['B14:L21', "1. Bin/trolley must be changed only after complete usage of all material in it. \n2. Empty bin/trolley should be replaced with new one \n3. Don't top up partially filled bin\n4.Follow FIFO on line during Process .\n5.Do not use next bin / Trolley material until running not consumed."],
    ['M14:U15', 'Do not use any lubrication if not specified in OPS / Process Sheet.'],
    ['M16:R16', 'Do not use mobile on the shopfloor'],
    ['S16:U16', 'Do not run on the shopfloor'],
  ];
  for (const [range, text] of genInstr) {
    mergeAndSet(ws, range, text, cellStyle(false, 20, 'left', 'center', true, 'medium', 'thin', 'thin', 'medium'));
  }
  for (let r = 17; r <= 21; r++) {
    mergeAndSet(ws, `M${r}:R${r}`, '', cellStyle(false, 20, 'left', 'center', false, 'thin', r === 21 ? 'medium' : 'thin', 'thin', 'thin'));
    mergeAndSet(ws, `S${r}:U${r}`, '', cellStyle(false, 20, 'left', 'center', false, 'thin', r === 21 ? 'medium' : 'thin', 'thin', 'thin'));
  }

  fillBorderRange(ws, col('V'), col('AD'), 13, 20, 'thin', 'medium', 'thin', 'thin');
  fillBorderRange(ws, col('AE'), col('AL'), 13, 20, 'thin', 'medium', 'thin', 'thin');
  fillBorderRange(ws, col('AM'), col('BC'), 13, 20, 'thin', 'medium', 'thin', 'thin');

  const emsInstr = [
    "1. Do waste segregation.\n2. Switch off idle lights & machines\n3. Ensure 3R Principal in daily activities\n4. If there was any leakage, communicate to Sub Leader",
    "1. Follow POS sheet in case of any Chemical\n2. Follow MSDS/SDS in case of any emergency regarding chemical\n3. Follow your PPE's",
  ];
  mergeAndSet(ws, "BD14:BG21", emsInstr[0], cellStyle(false, 20, "left", "center", true, "thin", "thin", "thin", "thin"));
  mergeAndSet(ws, "BH14:BK21", emsInstr[1], cellStyle(false, 20, "left", "center", true, "thin", "thin", "thin", "thin"));

  fillBorderRange(ws, col("BL"), col("BW"), 12, 20, "thin", "medium", "thin", "thin");

  const processInstr = templateConfig?.processInstructions || [
    '2. Do Not Use Fallen Electrical/Functional Parts.',
    '3. Ensure Model / Variant Change.',
    '4. Report in case of part / hardware fallen inside vehicle.',
    '5. TQ Wrench Arrow Mark should be in correct direction.',
    '6. Put Fallen Hardware in Red Bin for Zone In-Charge judgement.',
    '7. Take approval from SH / HOD before changing process sequence.',
    '8. Zone In-Charge is overall responsible to ensure work is as per OPS.',
    '9. Contaminant parts should be covered properly.',
  ];
  processInstr.forEach((text, i) => {
    mergeAndSet(ws, `BX${14 + i}:CM${14 + i}`, text,
      cellStyle(false, 20, 'left', 'center', true, 'thin', i === 7 ? 'medium' : 'thin', 'medium', 'thin'));
  });

  mergeAndSet(ws, 'B22:CM22', '', cellStyle(false, 11, 'left', 'center', false, null, null, 'medium', 'medium'));

  mergeAndSet(ws, 'B23:U64', 'Illustrations & Process Details',
    cellStyle(true, 26, 'center', 'top', false, 'medium', 'medium', 'medium', 'thin'));

  const defaultHeaders: Record<string, string> = {
    sn: 'SN', itemImportance: 'Item Importance', stepWhat: 'Shtep \\n(What / Activity)',
    methodHow: 'Method \\n(How)', frequencyWhen: 'Frequency / When', standardCriteria: 'Standard \\n(Spec. / Judgment Criteria)',
    responsibility: 'Responsibility', equipmentMeasuring: 'Equipment /\\nMeasuring Eq.',
    possibleAbnormalities: 'Possible \\nAbnormalities', reactionPlan: 'Reaction \\nPlan',
    partName: 'Part Name \\n& QTY', ppeRequired: 'PPEs\\nrequired',
    recordDocument: 'Record /\\nDocument', remarks: 'Remarks'
  };
  const th = templateConfig?.tableHeaders || defaultHeaders;

  const colHeaders = [
    ['V23:V24', th.sn || 'SN'], ['W23:Z24', th.itemImportance || 'Item Importance'], ['AA23:AJ24', th.stepWhat || 'Shtep \\n(What / Activity)'],
    ['AK23:AR24', th.methodHow || 'Method \\n(How)'], ['AS23:AX24', th.frequencyWhen || 'Frequency / When'], ['AY23:BG24', th.standardCriteria || 'Standard \\n(Spec. / Judgment Criteria)'],
    ['BH23:BH24', th.responsibility || 'Responsibility'], ['BI23:BJ24', th.equipmentMeasuring || 'Equipment /\\nMeasuring Eq.'],
    ['BK23:BO24', th.possibleAbnormalities || 'Possible \\nAbnormalities'], ['BP23:BS24', th.reactionPlan || 'Reaction \\nPlan'],
    ['BT23:BY24', th.partName || 'Part Name \\n& QTY'], ['BZ23:CD24', th.ppeRequired || 'PPEs\\nrequired'],
    ['CE23:CI24', th.recordDocument || 'Record /\\nDocument'], ['CJ23:CM24', th.remarks || 'Remarks'],
  ];
  for (const [range, label] of colHeaders) {
    mergeAndSet(ws, range, label,
      cellStyle(true, 22, 'center', 'center', true, 'medium', 'thin', 'thin', range.startsWith('CJ') ? 'medium' : 'thin'));
  }

  processSteps.forEach((step, idx) => {
    const er = 25 + idx * 8;
    const isLast = idx === processSteps.length - 1;
    const ob = isLast ? 'medium' : 'thin';

    mergeAndSet(ws, `V${er}:V${er}`, step.sn,
      cellStyle(true, 22, 'center', 'center', false, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `W${er}:Z${er}`, step.itemImportance,
      cellStyle(true, 72, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `AA${er}:AJ${er}`, step.stepWhat,
      cellStyle(false, 22, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `AK${er}:AR${er}`, step.methodHow,
      cellStyle(false, 22, 'left', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `AS${er}:AX${er}`, step.frequencyWhen,
      cellStyle(false, 22, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `AY${er}:BG${er}`, step.standardCriteria,
      cellStyle(false, 22, 'left', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `BH${er}:BH${er}`, step.responsibility,
      cellStyle(false, 22, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `BI${er}:BJ${er}`, step.equipmentMeasuring,
      cellStyle(false, 22, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `BK${er}:BO${er}`, step.possibleAbnormalities,
      cellStyle(false, 22, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `BP${er}:BS${er}`, step.reactionPlan,
      cellStyle(false, 22, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `BT${er}:BW${er}`, step.partName,
      cellStyle(false, 22, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `BX${er}:BY${er}`, step.partQty,
      cellStyle(false, 22, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `BZ${er}:CD${er}`, step.ppeRequired,
      cellStyle(false, 22, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `CE${er}:CI${er}`, step.recordDocument,
      cellStyle(false, 22, 'center', 'center', true, 'thin', ob, 'thin', 'thin'));
    mergeAndSet(ws, `CJ${er}:CM${er}`, step.remarks,
      cellStyle(false, 22, 'left', 'center', true, 'thin', ob, 'thin', 'medium'));
  });

  mergeAndSet(ws, 'B65:CM65', '', cellStyle(false, 11, 'left', 'center', false, null, null, 'medium', 'medium'));

  mergeAndSet(ws, 'B66:T69',
    templateConfig?.abnormalityHandlingRoute || 'Abnormality handling route : \\nIn case of any abnormality inform the Zone In-Charge\\nFlow of Communication :-\\nOperator ► Team Member ► Section Mgr ► As required',
    cellStyle(true, 26, 'left', 'center', true, 'medium', 'thin', 'medium', 'thin'));
  mergeAndSet(ws, 'U66:CM69', abnormalityDetails || templateConfig?.abnormalityDetailsLabel || 'Past Problem Details',
    cellStyle(true, 26, 'center', 'top', false, 'medium', 'thin', 'thin', 'medium'));

  for (let r = 67; r <= 69; r++) {
    mergeAndSet(ws, `B${r}:T${r}`, '', cellStyle(false, 11, 'left', 'center', false, 'thin', r === 69 ? 'medium' : 'thin', 'medium', 'thin'));
    mergeAndSet(ws, `U${r}:CM${r}`, '', cellStyle(false, 11, 'left', 'center', false, 'thin', r === 69 ? 'medium' : 'thin', 'thin', 'medium'));
  }

  mergeAndSet(ws, 'B70:CM70', '', cellStyle(false, 11, 'left', 'center', false, null, null, 'medium', 'medium'));

  mergeAndSet(ws, 'B71:G73', templateConfig?.associateSignArea?.title1 || 'Associate Name \\n& Emp. Code',
    cellStyle(true, 26, 'left', 'center', true, 'medium', 'thin', 'medium', 'thin'));
  mergeAndSet(ws, 'B74:G76', templateConfig?.associateSignArea?.title2 || 'Sign & Date',
    cellStyle(true, 26, 'left', 'center', true, 'thin', 'thin', 'medium', 'thin'));

  const a71 = ['H71:K73', 'L71:Q73', 'R71:U73', 'V71:Z73', 'AA71:AF73', 'AG71:AK73', 'AL71:AO73', 'AP71:AS73', 'AT71:AZ73', 'BA71:BD73', 'BE71:BE73', 'BF71:BG73', 'BH71:BH73', 'BI71:BJ73', 'BK71:BL73', 'BM71:BQ73', 'BR71:BT73', 'BU71:BY73', 'BZ71:CD73', 'CE71:CG73', 'CH71:CJ73', 'CK71:CM73'];
  const a74 = ['H74:K76', 'L74:Q76', 'R74:U76', 'V74:Z76', 'AA74:AF76', 'AG74:AK76', 'AL74:AO76', 'AP74:AS76', 'AT74:AZ76', 'BA74:BD76', 'BE74:BE76', 'BF74:BG76', 'BH74:BH76', 'BI74:BJ76', 'BK74:BL76', 'BM74:BQ76', 'BR74:BT76', 'BU74:BY76', 'BZ74:CD76', 'CE74:CG76', 'CH74:CJ76', 'CK74:CM76'];
  for (const range of a71) {
    mergeAndSet(ws, range, '', cellStyle(false, 11, 'left', 'center', false, 'medium', 'thin', 'thin', range.startsWith('CK') ? 'medium' : 'thin'));
  }
  for (const range of a74) {
    mergeAndSet(ws, range, '', cellStyle(false, 11, 'left', 'center', false, 'thin', 'thin', 'thin', range.startsWith('CK') ? 'medium' : 'thin'));
  }

  mergeAndSet(ws, 'B77:CD77', '', cellStyle(false, 11, 'left', 'center', false, 'medium', 'medium', 'medium', null));
  mergeAndSet(ws, 'CE77:CM77', 'Page Number : XX / XX',
    cellStyle(true, 26, 'center', 'center', false, 'medium', 'medium', null, 'medium'));

  // Image embedding
  if (images.logoBImageBase64) addImage(ws, toRawBase64(images.logoBImageBase64), 'jpeg', col('B'), 1, col('E'), 10);
  if (images.stopCallWaitImageBase64) addImage(ws, toRawBase64(images.stopCallWaitImageBase64), 'jpeg', col('B'), 65, col('T'), 68);
  if (images.noSymbolImageBase64) addImage(ws, toRawBase64(images.noSymbolImageBase64), 'jpeg', col('V'), 13, col('X'), 15);
  if (images.ppeGuideImageBase64) addImage(ws, toRawBase64(images.ppeGuideImageBase64), 'jpeg', col('V'), 13, col('AD'), 20);
  if (images.fiveSImageBase64) addImage(ws, toRawBase64(images.fiveSImageBase64), 'jpeg', col('BL'), 12, col('BW'), 20);
  if (images.qrCodeImageBase64) addImage(ws, toRawBase64(images.qrCodeImageBase64), 'jpeg', col('CE'), 8, col('CM'), 10);

  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, ws, 'Format');
  return writeFile(workbook, 'OPS_Format.xlsx');
}

export { OpsFormImages, ProcessStepData };