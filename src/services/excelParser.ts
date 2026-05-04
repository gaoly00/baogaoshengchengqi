/**
 * Excel 数据提取服务
 *
 * 表格截取规则：
 * 1. {{table:名称}} 所在行是标记行，不纳入表格数据
 * 2. 竖向：从标记下一行开始向下，遇空行或下一个表格标记停止
 * 3. 横向：从标记下一行首列开始向右，遇空列或标记停止
 * 4. 每个 Table 独立提取，输出纯二维数组
 */

import * as XLSX from 'xlsx';

export interface FieldMapping {
  fieldName: string;
  sheetName: string;
  cellAddress: string;
}

export interface ExtractedData {
  [key: string]: string;
}

/** 纯二维数组：data[row][col]，data[0][0] = 表格标记正下方第一个单元格 */
export interface TableData {
  tableName: string;
  data: string[][];
}

export interface ExtractionResult {
  fields: ExtractedData;
  tables: TableData[];
  debug?: string;
}

// ---- 正则 ----

const RE_TABLE_MARKER = /^\{\{table:(.+?)\}\}$/;
const RE_ANY_MARKER   = /^\{\{.+?\}\}$/;

// ---- 辅助 ----

function isBlank(val: string | undefined): boolean {
  return val === undefined || val === null || String(val).trim() === '';
}

function isTableMarker(val: string | undefined): boolean {
  if (!val) return false;
  return RE_TABLE_MARKER.test(String(val));
}

function isAnyMarker(val: string | undefined): boolean {
  if (!val) return false;
  return RE_ANY_MARKER.test(String(val));
}

/** 读取单元格值（1-indexed），支持合并单元格 */
function readCell(sheet: XLSX.WorkSheet, row: number, col: number): string | undefined {
  const r = row - 1;
  const c = col - 1;
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = sheet[addr];
  if (cell) {
    if (cell.v !== undefined && cell.v !== null) return String(cell.v);
    if (cell.w !== undefined && cell.w !== null) return String(cell.w);
  }
  // 如果当前单元格为空，检查是否属于合并区域，返回合并区域左上角的值
  const merges = sheet['!merges'];
  if (merges) {
    for (const range of merges) {
      if (r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c) {
        const originAddr = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c });
        const originCell = sheet[originAddr];
        if (originCell) {
          if (originCell.v !== undefined && originCell.v !== null) return String(originCell.v);
          if (originCell.w !== undefined && originCell.w !== null) return String(originCell.w);
        }
        return undefined;
      }
    }
  }
  return undefined;
}

/** 判断单元格是否属于合并区域但不是左上角原点（应当跳过） */
function isNonOriginMerge(sheet: XLSX.WorkSheet, row: number, col: number): boolean {
  const r = row - 1;
  const c = col - 1;
  const merges = sheet['!merges'];
  if (!merges) return false;
  for (const range of merges) {
    if (r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c) {
      return r !== range.s.r || c !== range.s.c;
    }
  }
  return false;
}

/** 获取 sheet 实际行列范围（含合并区域） */
function getSheetBounds(sheet: XLSX.WorkSheet): { maxRow: number; maxCol: number } | null {
  let maxR = -1;
  let maxC = -1;
  for (const key of Object.keys(sheet)) {
    if (key.startsWith('!')) continue;
    try {
      const addr = XLSX.utils.decode_cell(key);
      if (addr.r > maxR) maxR = addr.r;
      if (addr.c > maxC) maxC = addr.c;
    } catch { /* skip */ }
  }
  // 合并区域可能超出已有单元格的范围
  const merges = sheet['!merges'];
  if (merges) {
    for (const range of merges) {
      if (range.e.r > maxR) maxR = range.e.r;
      if (range.e.c > maxC) maxC = range.e.c;
    }
  }
  if (maxR < 0 || maxC < 0) return null;
  return { maxRow: maxR + 1, maxCol: maxC + 1 };
}

// ---- 独立函数：按流程拆分为 find → getRange → readArray ----

interface CellPos { row: number; col: number }

interface TableRange { startRow: number; startCol: number; endRow: number; endCol: number }

/**
 * Step 1: 扫描 sheet，找到所有 {{table:名称}} 标记
 */
function findTableMarkers(
  sheet: XLSX.WorkSheet,
  maxRow: number,
  maxCol: number
): { marker: CellPos; tableName: string }[] {
  const list: { marker: CellPos; tableName: string }[] = [];
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const val = readCell(sheet, r, c);
      if (!val) continue;
      const m = val.match(RE_TABLE_MARKER);
      if (m) list.push({ marker: { row: r, col: c }, tableName: m[1] });
    }
  }
  return list;
}

/**
 * Step 2: 对单个 table 标记，计算它在 sheet 中的矩形范围
 *
 * 规则：
 *   startRow = markerRow + 1, startCol = markerCol
 *   横向：扫描所有数据行，取最大非空列（处理合并单元格导致的空格）
 *   纵向从 startRow 向下扫描，遇空行 / table 标记停止
 */
function getTableRange(
  sheet: XLSX.WorkSheet,
  markerRow: number,
  markerCol: number,
  maxRow: number,
  maxCol: number
): TableRange | null {
  const startRow = markerRow + 1;
  const startCol = markerCol;

  if (startRow > maxRow) return null;

  // -- 先确定纵向边界 --
  // 初步横向上限：用 sheet 的 maxCol
  let endRow = startRow;
  for (let rr = startRow; rr <= maxRow; rr++) {
    // 检查该行是否有 table 标记
    let foundTableMarker = false;
    for (let cc = startCol; cc <= maxCol; cc++) {
      if (isTableMarker(readCell(sheet, rr, cc))) {
        foundTableMarker = true;
        break;
      }
    }
    if (foundTableMarker) { endRow = rr - 1; break; }

    // 检查该行从 startCol 开始是否全空
    let allBlank = true;
    for (let cc = startCol; cc <= maxCol; cc++) {
      if (!isBlank(readCell(sheet, rr, cc))) {
        allBlank = false;
        break;
      }
    }
    if (allBlank) { endRow = rr - 1; break; }

    endRow = rr;
  }
  if (endRow < startRow) return null;

  // -- 横向边界：逐列扫描，遇全空列或标记列停止 --
  // 对每一列，检查所有数据行：有标记 → 停；全空 → 停；否则纳入并继续
  let endCol = startCol;
  for (let cc = startCol; cc <= maxCol; cc++) {
    let columnHasMarker = false;
    let columnHasContent = false;
    for (let rr = startRow; rr <= endRow; rr++) {
      const v = readCell(sheet, rr, cc);
      if (isAnyMarker(v)) { columnHasMarker = true; break; }
      if (!isBlank(v)) { columnHasContent = true; }
    }
    if (columnHasMarker) break;
    if (!columnHasContent) break;
    endCol = cc;
  }
  if (endCol < startCol) return null;

  return { startRow, startCol, endRow, endCol };
}

/**
 * Step 3: 将 sheet 中 range 指定的区域读取为纯二维数组
 *   tableData[row][col] 0-indexed, 一一对应 Excel 单元格
 */
function readRangeToArray(sheet: XLSX.WorkSheet, range: TableRange): string[][] {
  const data: string[][] = [];
  for (let r = range.startRow; r <= range.endRow; r++) {
    const row: string[] = [];
    for (let c = range.startCol; c <= range.endCol; c++) {
      if (isNonOriginMerge(sheet, r, c)) continue;
      row.push(readCell(sheet, r, c) || '');
    }
    data.push(row);
  }
  return data;
}

// ---- 核心：合并流程 ----

export function extractAllFromExcel(fileBuffer: ArrayBuffer): ExtractionResult {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const fields: ExtractedData = {};
  const tables: TableData[] = [];
  const debugLines: string[] = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const bounds = getSheetBounds(sheet);
    if (!bounds) continue;

    const { maxRow, maxCol } = bounds;
    debugLines.push(`Sheet "${sheetName}": ${maxRow} rows × ${maxCol} cols`);

    // 记录被表格占用的行，字段提取时跳过
    const consumedRows = new Set<number>();
    // 记录已处理的表格标记 "r,c"，避免重复处理
    const processedKeys = new Set<string>();

    const markers = findTableMarkers(sheet, maxRow, maxCol);

    for (const { marker, tableName } of markers) {
      const key = `${marker.row},${marker.col}`;
      if (processedKeys.has(key)) continue;
      processedKeys.add(key);

      // 标记行本身不纳入数据
      consumedRows.add(marker.row);

      const range = getTableRange(sheet, marker.row, marker.col, maxRow, maxCol);

      if (!range) {
        debugLines.push(`  Table "${tableName}" @(${marker.row},${marker.col}): EMPTY (no data rows)`);
        continue;
      }

      // 标记数据行已被占用
      for (let rr = range.startRow; rr <= range.endRow; rr++) {
        consumedRows.add(rr);
      }

      const tableData = readRangeToArray(sheet, range);

      // ---- 调试日志 ----
      console.log('tableName:', tableName);
      console.log('range:', range);
      console.log('tableData[0][0]:', tableData?.[0]?.[0]);
      console.log('rowCount:', tableData.length);
      console.log('colCount:', tableData[0]?.length);

      debugLines.push(
        `  Table "${tableName}" @(${marker.row},${marker.col}): ` +
        `range=${range.startRow}-${range.endRow} × ${range.startCol}-${range.endCol} ` +
        `data=${tableData.length}r×${(tableData[0] || []).length}c ` +
        `[0][0]="${tableData[0]?.[0] ?? ''}"`
      );

      tables.push({ tableName, data: tableData });
    }

    // ---- 字段提取：从未被表格占用的行中提取 {{字段名}} ----
    for (let r = 1; r <= maxRow; r++) {
      if (consumedRows.has(r)) continue;
      const val = readCell(sheet, r, 1);
      if (!val) continue;
      const m = val.match(RE_ANY_MARKER);
      if (!m) continue;
      if (isTableMarker(val)) continue;
      const fieldName = m[1];
      const value = readCell(sheet, r, 2) || '';
      fields[fieldName] = value;
    }
  }

  const result: ExtractionResult = { fields, tables };
  if (debugLines.length > 0) {
    result.debug = debugLines.join('\n');
  }
  return result;
}

// ---- 向后兼容 ----

export function extractTableRegions(fileBuffer: ArrayBuffer): TableData[] {
  return extractAllFromExcel(fileBuffer).tables;
}

export function extractDataFromExcel(
  fileBuffer: ArrayBuffer,
  mappings: FieldMapping[]
): ExtractedData {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const result: ExtractedData = {};

  for (const mapping of mappings) {
    const sheet = workbook.Sheets[mapping.sheetName];
    if (!sheet) continue;
    const cell = sheet[mapping.cellAddress];
    if (!cell) continue;
    result[mapping.fieldName] = cell.v !== undefined ? String(cell.v) : '';
  }

  return result;
}

export function autoExtractFromExcel(fileBuffer: ArrayBuffer): ExtractedData {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const result: ExtractedData = {};

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    let row = 1;
    while (true) {
      const fieldCell = sheet[`A${row}`];
      if (!fieldCell) break;
      const fieldName = fieldCell.v;
      if (typeof fieldName === 'string') {
        const match = fieldName.match(/^\{\{(.+)\}\}$/);
        if (match) {
          const valueCell = sheet[`B${row}`];
          result[match[1]] = valueCell?.v !== undefined ? String(valueCell.v) : '';
        }
      }
      row++;
    }
  }

  return result;
}

export function getSheetNames(fileBuffer: ArrayBuffer): string[] {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  return workbook.SheetNames;
}

export function getSheetRange(
  fileBuffer: ArrayBuffer,
  sheetName: string
): { minRow: number; maxRow: number; minCol: number; maxCol: number } | null {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const sheet = workbook.Sheets[sheetName];
  if (!sheet || !sheet['!ref']) return null;
  const range = XLSX.utils.decode_range(sheet['!ref']);
  return {
    minRow: range.s.r + 1,
    maxRow: range.e.r + 1,
    minCol: range.s.c + 1,
    maxCol: range.e.c + 1,
  };
}

export function getCellValue(
  fileBuffer: ArrayBuffer,
  sheetName: string,
  cellAddress: string
): string | null {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return null;
  const cell = sheet[cellAddress];
  return cell?.v !== undefined ? String(cell.v) : null;
}
