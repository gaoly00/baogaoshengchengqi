/**
 * Word 模板处理服务
 * 支持读取 .docx 文件并替换占位符
 * 处理跨 w:r 分割的占位符（不同文字格式导致的分割）
 */

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType
} from 'docx';
import JSZip from 'jszip';
import { ExtractedData, TableData } from './excelParser';

const PLACEHOLDER_REGEX = /\{\{(.+?)\}\}/g;

// <w:t ...>content</w:t>
const WT_TAG_REGEX = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;

interface TElement {
  /** 在原始 XML 中的起始位置（整个 <w:t ...> 的起始） */
  start: number;
  /** 内容起始位置（> 之后） */
  contentStart: number;
  /** 内容结束位置（</w:t> 之前） */
  contentEnd: number;
  /** 标签结束位置（</w:t> 之后） */
  end: number;
  /** 文本内容 */
  content: string;
}

/**
 * 在 XML 字符串中定位所有 <w:t> 元素
 */
function findTextElements(xml: string): TElement[] {
  const result: TElement[] = [];
  const regex = new RegExp(WT_TAG_REGEX.source, 'g');
  let match: RegExpExecArray | null;

  while ((match = regex.exec(xml)) !== null) {
    const fullMatch = match[0];
    const content = match[1];
    const start = match.index;
    const tagOpenEnd = start + fullMatch.indexOf('>') + 1;
    const end = start + fullMatch.length;
    const contentEnd = end - '</w:t>'.length;

    result.push({
      start,
      contentStart: tagOpenEnd,
      contentEnd,
      end,
      content
    });
  }

  return result;
}

/**
 * XML 字符转义
 */
function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

const TABLE_PLACEHOLDER_REGEX = /\{\{table:(.+?)\}\}/g;

export interface TableOptions {
  fontFamily: string;  // '宋体' | '黑体' | '微软雅黑'
  fontSize: number;    // pt, e.g. 10
}

function buildCellXml(text: string, isHeader: boolean, opts: TableOptions): string {
  const escText = escapeXml(text);
  const width = 2400;
  const bold = isHeader ? '<w:b/><w:bCs/>' : '';
  const fontSizeHalfPt = opts.fontSize * 2;

  return `
<w:tc>
  <w:tcPr>
    <w:tcW w:w="${width}" w:type="dxa"/>
    <w:tcBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tcBorders>
  </w:tcPr>
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:r>
      <w:rPr>
        ${bold}
        <w:rFonts w:eastAsia="${opts.fontFamily}"/>
        <w:sz w:val="${fontSizeHalfPt}"/>
        <w:szCs w:val="${fontSizeHalfPt}"/>
      </w:rPr>
      <w:t xml:space="preserve">${escText}</w:t>
    </w:r>
  </w:p>
</w:tc>`.trim();
}

function buildTableXml(table: TableData, opts: TableOptions): string {
  const data = table.data;
  if (data.length === 0) return '';

  const colCount = data[0].length;
  const colWidth = Math.floor(9000 / colCount);

  const gridCols = Array(colCount).fill(null).map(() =>
    `<w:gridCol w:w="${colWidth}"/>`
  ).join('');

  // Title paragraph above the table
  const titleFontSizeHalfPt = (opts.fontSize + 2) * 2;
  const titleXml = `
<w:p>
  <w:pPr>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:b/>
      <w:rFonts w:eastAsia="${opts.fontFamily}"/>
      <w:sz w:val="${titleFontSizeHalfPt}"/>
      <w:szCs w:val="${titleFontSizeHalfPt}"/>
    </w:rPr>
    <w:t>${escapeXml(table.tableName)}</w:t>
  </w:r>
</w:p>`.trim();

  // Render all rows from data[r][c]; first row bold as header
  const rowsXml = data.map((rowData, ri) => {
    const isFirstRow = ri === 0;
    const cells = rowData.map(cell =>
      buildCellXml(cell, isFirstRow, opts)
    ).join('');
    return `<w:tr>${cells}</w:tr>`;
  }).join('');

  const tableXml = `
<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tblGrid>${gridCols}</w:tblGrid>
  ${rowsXml}
</w:tbl>`.trim();

  return titleXml + tableXml;
}

/**
 * 在文档 XML 中替换 {{table:表名}} 为实际表格
 */
function replaceTablePlaceholders(documentXml: string, tableData: TableData[], opts: TableOptions): string {
  if (!tableData || tableData.length === 0) return documentXml;

  let xml = documentXml;
  const tableMap = new Map(tableData.map(t => [t.tableName, t]));

  // 用正则找所有 <w:p> 段落
  const pRegex = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g;
  const paragraphs: { full: string; start: number; end: number; plainText: string }[] = [];

  let match: RegExpExecArray | null;
  while ((match = pRegex.exec(xml)) !== null) {
    const full = match[0];
    const plainText = full.replace(/<[^>]+>/g, '');
    paragraphs.push({
      full,
      start: match.index,
      end: match.index + full.length,
      plainText,
    });
  }

  // 从后往前替换，避免位置偏移
  for (const para of paragraphs.reverse()) {
    const m = TABLE_PLACEHOLDER_REGEX.exec(para.plainText);
    TABLE_PLACEHOLDER_REGEX.lastIndex = 0;
    if (!m) continue;

    const tableName = m[1];
    const table = tableMap.get(tableName);
    if (!table) continue;

    const generated = buildTableXml(table, opts);
    xml = xml.substring(0, para.start) + generated + xml.substring(para.end);
  }

  return xml;
}

/**
 * 从 Word 文档中提取 {{table:表名}} 引用的表名列表
 */
export async function extractTablePlaceholders(docxBuffer: ArrayBuffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(docxBuffer);
  const documentXml = await zip.file('word/document.xml')?.async('string');
  if (!documentXml) return [];

  const plainText = documentXml.replace(/<[^>]+>/g, '');
  const names = new Set<string>();
  let m: RegExpExecArray | null;
  const regex = new RegExp(TABLE_PLACEHOLDER_REGEX.source, 'g');
  while ((m = regex.exec(plainText)) !== null) {
    names.add(m[1]);
  }
  return Array.from(names);
}

/**
 * 替换 Word 文档中的占位符
 * 支持跨多个 w:r 元素的占位符（因格式不同而被分割）
 */
const DEFAULT_TABLE_OPTIONS: TableOptions = {
  fontFamily: '宋体',
  fontSize: 10,
};

export async function replacePlaceholders(
  docxBuffer: ArrayBuffer,
  data: ExtractedData,
  tableData?: TableData[],
  tableOpts?: TableOptions
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(docxBuffer);
  let documentXml = await zip.file('word/document.xml')?.async('string');

  if (!documentXml) {
    throw new Error('无法读取 Word 文档内容');
  }

  // 先替换表格占位符
  if (tableData && tableData.length > 0) {
    const opts = tableOpts || DEFAULT_TABLE_OPTIONS;
    documentXml = replaceTablePlaceholders(documentXml, tableData, opts);
  }

  for (const [fieldName, value] of Object.entries(data)) {
    const placeholder = `{{${fieldName}}}`;
    const escValue = escapeXml(value || '');

    // 循环替换，处理同一字段在文档中多次出现的情况
    let safety = 0;
    while (safety++ < 1000) {
      // 快速路径：占位符在 XML 中是连续字符串
      if (documentXml.includes(placeholder)) {
        documentXml = documentXml.split(placeholder).join(escValue);
        continue;
      }

      // 慢速路径：占位符可能被 XML 标签分割
      const tElements = findTextElements(documentXml);
      const fullText = tElements.map(te => te.content).join('');

      const pos = fullText.indexOf(placeholder);
      if (pos === -1) break;

      // 定位占位符跨越了哪些 w:t 元素
      let startIdx = -1;
      let endIdx = -1;
      let charOffset = 0;

      for (let i = 0; i < tElements.length; i++) {
        const nextOffset = charOffset + tElements[i].content.length;

        if (startIdx === -1 && charOffset <= pos && nextOffset > pos) {
          startIdx = i;
        }

        if (startIdx !== -1 && nextOffset >= pos + placeholder.length) {
          endIdx = i;
          break;
        }

        charOffset = nextOffset;
      }

      if (startIdx === -1 || endIdx === -1) break;

      // 收集要修改的区间
      const changes: { start: number; end: number; replacement: string }[] = [];

      for (let i = startIdx; i <= endIdx; i++) {
        const text = tElements[i].content;
        const elStart = tElements.slice(0, i).reduce((s, te) => s + te.content.length, 0);
        const elEnd = elStart + text.length;

        const localStart = Math.max(0, pos - elStart);
        const localEnd = Math.min(text.length, pos + placeholder.length - elStart);

        let newContent: string;
        if (i === startIdx && i === endIdx) {
          newContent = text.substring(0, localStart) + escValue + text.substring(localEnd);
        } else if (i === startIdx) {
          newContent = text.substring(0, localStart) + escValue;
        } else if (i === endIdx) {
          newContent = text.substring(localEnd);
        } else {
          newContent = '';
        }

        changes.push({
          start: tElements[i].contentStart,
          end: tElements[i].contentEnd,
          replacement: newContent
        });
      }

      // 从后往前应用修改，避免位置偏移
      for (const change of changes.reverse()) {
        documentXml = documentXml.substring(0, change.start) + change.replacement + documentXml.substring(change.end);
      }
    }
  }

  zip.file('word/document.xml', documentXml);
  return await zip.generateAsync({ type: 'arraybuffer' });
}

/**
 * 从 Word 文档中提取所有占位符
 * 支持跨多个 w:r 元素的占位符
 */
export async function extractPlaceholders(docxBuffer: ArrayBuffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(docxBuffer);
  const documentXml = await zip.file('word/document.xml')?.async('string');

  if (!documentXml) {
    throw new Error('无法读取 Word 文档内容');
  }

  // 去掉所有 XML 标签得到纯文本，然后匹配占位符
  const plainText = documentXml.replace(/<[^>]+>/g, '');

  const placeholders = new Set<string>();
  let match: RegExpExecArray | null;
  const regex = new RegExp(PLACEHOLDER_REGEX.source, 'g');

  while ((match = regex.exec(plainText)) !== null) {
    placeholders.add(match[1]);
  }

  return Array.from(placeholders);
}

/**
 * 验证数据是否包含所有必需的占位符
 */
export function validatePlaceholders(
  placeholders: string[],
  data: ExtractedData
): { missing: string[]; extra: string[] } {
  const required = new Set(placeholders);
  const provided = new Set(Object.keys(data));

  const missing = Array.from(required).filter(p => !provided.has(p));
  const extra = Array.from(provided).filter(p => !required.has(p));

  return { missing, extra };
}

/**
 * 创建包含占位符的示例 Word 文档
 */
export function createSampleTemplate(placeholders: string[]): Document {
  const children: Paragraph[] = [];

  children.push(
    new Paragraph({
      text: '房地产估价报告',
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
    })
  );

  children.push(new Paragraph({ text: '' }));

  children.push(
    new Paragraph({
      text: '一、基本信息',
      heading: HeadingLevel.HEADING_1,
    })
  );

  for (const placeholder of placeholders) {
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `${placeholder}：`,
            bold: true,
          }),
          new TextRun({
            text: `{{${placeholder}}}`,
            highlight: 'yellow',
          }),
        ],
      })
    );
  }

  children.push(new Paragraph({ text: '' }));

  children.push(
    new Paragraph({
      text: '二、估价说明',
      heading: HeadingLevel.HEADING_1,
    })
  );

  children.push(
    new Paragraph({
      text: '本报告基于现场勘查和市场调研编制，估价师对报告内容的真实性和准确性负责。',
    })
  );

  return new Document({
    creator: '估价报告生成器',
    title: '房地产估价报告模板',
    description: '使用 {{字段名}} 格式标记需要替换的内容',
    sections: [
      {
        properties: {},
        children: children,
      },
    ],
  });
}

/**
 * 下载 docx 文件（直接下载到默认目录）
 */
export function downloadDocx(buffer: ArrayBuffer, filename: string): void {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });
  downloadBlob(blob, filename);
}

/**
 * "另存为" — 弹出系统保存对话框让用户选择路径
 * 返回保存后的文件名，如果用户取消则返回 null
 */
export async function saveAsDocx(buffer: ArrayBuffer, suggestedName: string): Promise<string | null> {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });

  try {
    const handle = await (window as any).showSaveFilePicker({
      suggestedName,
      types: [{
        description: 'Word 文档',
        accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] }
      }]
    });
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return handle.name;
  } catch (err: any) {
    if (err.name === 'AbortError') return null; // 用户取消
    throw err;
  }
}

/**
 * 下载 Blob 文件
 */
export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
