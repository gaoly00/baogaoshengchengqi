/**
 * Word 模板处理服务
 * 支持读取 .docx 文件并替换占位符
 */

import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  HeadingLevel,
  AlignmentType,
  convertInchesToTwip,
  Table,
  TableRow,
  TableCell,
  BorderStyle
} from 'docx';
import JSZip from 'jszip';
import { ExtractedData } from './excelParser';

/**
 * 占位符格式说明
 * - 单行占位符: {{字段名}}
 * - 保留占位符（不替换）: {{{字段名}}}
 */

// 占位符正则
const PLACEHOLDER_REGEX = /\{\{(.+?)\}\}/g;
const ESCAPE_REGEX = /^\{(.+)\}$/;

/**
 * 替换 Word 文档中的占位符
 * @param docxBuffer 原始 docx 文件的 ArrayBuffer
 * @param data 要替换的数据
 * @returns 替换后的 docx 文件的 ArrayBuffer
 */
export async function replacePlaceholders(
  docxBuffer: ArrayBuffer,
  data: ExtractedData
): Promise<ArrayBuffer> {
  // 解析 docx 文件
  const zip = await JSZip.loadAsync(docxBuffer);
  const documentXml = await zip.file('word/document.xml')?.async('string');
  
  if (!documentXml) {
    throw new Error('无法读取 Word 文档内容');
  }

  // 替换占位符
  let modifiedXml = documentXml;
  
  for (const [fieldName, value] of Object.entries(data)) {
    const placeholder = `{{${fieldName}}}`;
    modifiedXml = modifiedXml.split(placeholder).join(value || '');
  }

  // 更新文档 XML
  zip.file('word/document.xml', modifiedXml);

  // 生成新的 docx 文件
  const newDocxBuffer = await zip.generateAsync({ type: 'arraybuffer' });
  return newDocxBuffer;
}

/**
 * 从 Word 文档中提取所有占位符
 */
export async function extractPlaceholders(docxBuffer: ArrayBuffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(docxBuffer);
  const documentXml = await zip.file('word/document.xml')?.async('string');
  
  if (!documentXml) {
    throw new Error('无法读取 Word 文档内容');
  }

  const placeholders = new Set<string>();
  let match;
  
  while ((match = PLACEHOLDER_REGEX.exec(documentXml)) !== null) {
    placeholders.add(match[1]);
  }
  
  // 去重
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

  // 标题
  children.push(
    new Paragraph({
      text: '房地产估价报告',
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
    })
  );

  // 基本信息
  children.push(
    new Paragraph({
      text: '',
    })
  );

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

  // 结尾
  children.push(
    new Paragraph({
      text: '',
    })
  );

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
 * 下载 docx 文件
 */
export function downloadDocx(buffer: ArrayBuffer, filename: string): void {
  const blob = new Blob([buffer], { 
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
  });
  downloadBlob(blob, filename);
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
