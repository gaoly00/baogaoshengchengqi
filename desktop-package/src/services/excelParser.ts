/**
 * Excel 数据提取服务
 * 方案B：{{字段名}} 格式映射表
 */

import * as XLSX from 'xlsx';

export interface FieldMapping {
  fieldName: string;      // 字段名，如 "项目名称"
  sheetName: string;      // Sheet名称
  cellAddress: string;    // 单元格地址，如 "B2"
}

export interface ExtractedData {
  [key: string]: string;  // 字段名 -> 值
}

/**
 * 从 Excel 文件中提取字段映射表定义的数据
 * @param fileBuffer Excel 文件的 ArrayBuffer
 * @param mappings 字段映射数组
 * @returns 提取的数据对象
 */
export function extractDataFromExcel(
  fileBuffer: ArrayBuffer,
  mappings: FieldMapping[]
): ExtractedData {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const result: ExtractedData = {};

  for (const mapping of mappings) {
    const sheet = workbook.Sheets[mapping.sheetName];
    if (!sheet) {
      console.warn(`Sheet "${mapping.sheetName}" not found, skipping field "${mapping.fieldName}"`);
      continue;
    }

    const cell = sheet[mapping.cellAddress];
    if (!cell) {
      console.warn(`Cell "${mapping.cellAddress}" not found in sheet "${mapping.sheetName}"`);
      continue;
    }

    // 获取单元格的值
    const value = cell.v !== undefined ? String(cell.v) : '';
    result[mapping.fieldName] = value;
  }

  return result;
}

/**
 * 自动扫描 Excel 文件，提取所有 {{字段名}} 格式的字段
 * A列放字段名（格式：{{字段名}}），B列放对应的值
 * @param fileBuffer Excel 文件的 ArrayBuffer
 * @returns 提取的数据对象
 */
export function autoExtractFromExcel(fileBuffer: ArrayBuffer): ExtractedData {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const result: ExtractedData = {};

  // 遍历所有 Sheet
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    
    // 扫描 A 列，查找 {{字段名}} 格式的内容
    let row = 1;
    while (true) {
      const cellRef = `A${row}`;
      const valueCellRef = `B${row}`;
      
      const fieldCell = sheet[cellRef];
      if (!fieldCell) break;
      
      const fieldName = fieldCell.v;
      if (typeof fieldName !== 'string') {
        row++;
        continue;
      }
      
      // 检查是否是 {{字段名}} 格式
      const match = fieldName.match(/^\{\{(.+)\}\}$/);
      if (match) {
        const valueCell = sheet[valueCellRef];
        const value = valueCell?.v !== undefined ? String(valueCell.v) : '';
        result[match[1]] = value;
      }
      
      row++;
    }
  }

  return result;
}

/**
 * 获取 Excel 文件的所有 Sheet 名称
 */
export function getSheetNames(fileBuffer: ArrayBuffer): string[] {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  return workbook.SheetNames;
}

/**
 * 读取指定 Sheet 的数据范围
 */
export function getSheetRange(fileBuffer: ArrayBuffer, sheetName: string): { minRow: number; maxRow: number; minCol: number; maxCol: number } | null {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const sheet = workbook.Sheets[sheetName];
  if (!sheet || !sheet['!ref']) return null;
  
  const range = XLSX.utils.decode_range(sheet['!ref']);
  return {
    minRow: range.s.r + 1,
    maxRow: range.e.r + 1,
    minCol: range.s.c + 1,
    maxCol: range.e.c + 1
  };
}

/**
 * 读取指定单元格的值
 */
export function getCellValue(fileBuffer: ArrayBuffer, sheetName: string, cellAddress: string): string | null {
  const workbook = XLSX.read(fileBuffer, { type: 'array' });
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return null;
  
  const cell = sheet[cellAddress];
  return cell?.v !== undefined ? String(cell.v) : null;
}
