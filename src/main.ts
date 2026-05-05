/**
 * 估价报告生成器 - 主应用
 */

import './index.css';
import {
  extractDataFromExcel,
  autoExtractFromExcel,
  extractAllFromExcel,
  getSheetNames,
  getCellValue,
  FieldMapping,
  ExtractedData,
  TableData,
  FieldPosition
} from './services/excelParser';
import {
  replacePlaceholders,
  extractPlaceholders,
  validatePlaceholders,
  createSampleTemplate,
  downloadDocx,
  saveAsDocx,
  downloadBlob,
  TableOptions
} from './services/wordProcessor';
import { Packer } from 'docx';

// 应用版本号（每次发布新版本时手动修改这里）
const APP_VERSION = "1.0.0";

// 版本检测地址（GitHub Gist 的 raw 文件链接）
const VERSION_URL = "https://gist.githubusercontent.com/gaoly00/79066c73f63d9c94d91338cf623dc353/raw/version.json";

// 版本检测函数
async function checkForUpdates(): Promise<void> {
  try {
    const response = await fetch(VERSION_URL + "?t=" + Date.now());
    const data = await response.json();
    const latestVersion = data.version;
    const updateUrl = data.url;
    const updateNotes = data.notes || "";

    if (latestVersion > APP_VERSION) {
      const updateDiv = document.getElementById("update-notice");
      if (updateDiv) {
        updateDiv.innerHTML = `
          <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 16px 24px; border-radius: 12px; margin-bottom: 24px; display: flex; justify-content: space-between; align-items: center;">
            <div>
              <div style="font-size: 18px; font-weight: bold; margin-bottom: 4px;">发现新版本 v${latestVersion}</div>
              <div style="font-size: 13px; opacity: 0.85;">当前版本：v${APP_VERSION}</div>
            </div>
            <a href="${updateUrl}" target="_blank" style="background: white; color: #667eea; padding: 10px 20px; border-radius: 8px; text-decoration: none; font-weight: 600; font-size: 14px;">查看更新</a>
          </div>
        `;
      }
    }
  } catch (error) {
    // 检测失败，静默处理
    console.log("版本检测失败", error);
  }
}

// 应用状态
interface AppState {
  excelFile: File | null;
  wordFile: File | null;
  wordBuffer: ArrayBuffer | null;
  excelData: ExtractedData;
  tableData: TableData[];
  fieldPositions: FieldPosition[];
  placeholders: string[];
  mappings: FieldMapping[];
  sheetNames: string[];
  selectedSheet: string;
  status: 'idle' | 'loading' | 'success' | 'error';
  message: string;
  instructionCollapsed: boolean;
  tableFontFamily: string;
  tableFontSize: number;
}

const state: AppState = {
  excelFile: null,
  wordFile: null,
  wordBuffer: null,
  excelData: {},
  tableData: [],
  fieldPositions: [],
  placeholders: [],
  mappings: [],
  sheetNames: [],
  selectedSheet: '',
  status: 'idle',
  message: '',
  instructionCollapsed: true,
  tableFontFamily: '宋体',
  tableFontSize: 10
};

// 文件上传处理
function handleExcelUpload(event: Event): void {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  if (!file) return;

  state.excelFile = file;
  state.status = 'loading';
  render();

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const buffer = e.target?.result as ArrayBuffer;
      state.sheetNames = getSheetNames(buffer);
      state.selectedSheet = state.sheetNames[0] || '';
      const result = extractAllFromExcel(buffer);
      state.excelData = result.fields;
      state.tableData = result.tables;
      state.fieldPositions = result.fieldPositions || [];
      state.mappings = Object.keys(state.excelData).map((field, i) => ({
        fieldName: field,
        sheetName: state.selectedSheet,
        cellAddress: 'B' + (i + 1)
      }));
      state.status = 'success';
      const fieldCount = Object.keys(state.excelData).length;
      const tableCount = state.tableData.length;
      state.message = `成功读取 Excel，共 ${fieldCount} 个字段` + (tableCount > 0 ? `、${tableCount} 个表格区域` : '');
      // DEBUG: 附加提取详情
      if (result.debug) {
        state.message += '\n[调试]\n' + result.debug;
      }
    } catch (err) {
      state.status = 'error';
      state.message = '读取 Excel 文件失败：' + (err as Error).message;
    }
    render();
  };
  reader.onerror = () => {
    state.status = 'error';
    state.message = '读取 Excel 文件失败';
    render();
  };
  reader.readAsArrayBuffer(file);
}

function handleWordUpload(event: Event): void {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  if (!file) return;

  state.wordFile = file;
  state.status = 'loading';
  render();

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const buffer = e.target?.result as ArrayBuffer;
      // 缓存 buffer，避免后续 render() 重建 DOM 后 arrayBuffer() 失败
      state.wordBuffer = buffer;
      state.placeholders = await extractPlaceholders(buffer);
      state.status = 'success';
      state.message = state.placeholders.length > 0
        ? `成功读取 Word 模板，发现 ${state.placeholders.length} 个占位符`
        : 'Word 模板中没有发现占位符，请确保使用 {{字段名}} 格式';
    } catch (err) {
      state.status = 'error';
      state.message = '读取 Word 文件失败：' + (err as Error).message;
    }
    render();
  };
  reader.onerror = () => {
    state.status = 'error';
    state.message = '读取 Word 文件失败';
    render();
  };
  reader.readAsArrayBuffer(file);
}

// 生成报告
async function generateReport(): Promise<void> {
  const hasData = Object.keys(state.excelData).length > 0 || state.tableData.length > 0;
  if (!state.wordBuffer || !hasData) {
    state.status = 'error';
    state.message = !state.wordBuffer
      ? 'Word 文件尚未处理完毕，请稍等'
      : 'Excel 数据尚未提取完毕，请稍等';
    render();
    return;
  }

  const validation = validatePlaceholders(state.placeholders, state.excelData);
  const hasIssues = validation.extra.length > 0 || validation.missing.length > 0;

  if (hasIssues) {
    showPreviewDialog(validation, state.excelData, state.fieldPositions);
    return;
  }

  await doGenerate();
}

/** 实际执行生成 */
async function doGenerate(): Promise<void> {
  state.status = 'loading';
  state.message = '正在生成报告...';
  render();

  try {
    const wordBuffer = state.wordBuffer || await state.wordFile!.arrayBuffer();

    const tableOpts: TableOptions = {
      fontFamily: state.tableFontFamily,
      fontSize: state.tableFontSize,
    };
    const resultBuffer = await replacePlaceholders(wordBuffer, state.excelData, state.tableData, tableOpts);

    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const originalName = state.wordFile!.name.replace('.docx', '');
    const outputName = `${originalName}_${timestamp}.docx`;

    // 尝试弹出 "另存为" 对话框
    let savedName: string | null = null;
    try {
      savedName = await saveAsDocx(resultBuffer, outputName);
    } catch {
      // showSaveFilePicker 不支持时回退到直接下载
      downloadDocx(resultBuffer, outputName);
      savedName = outputName;
    }

    if (savedName === null) {
      // 用户取消了保存对话框
      render();
      return;
    }

    const validation = validatePlaceholders(state.placeholders, state.excelData);

    state.status = 'success';
    state.message = `报告生成成功：${savedName}`;

    showResultDialog(savedName, validation, state.excelData, state.placeholders, state.tableData);

  } catch (err) {
    state.status = 'error';
    state.message = '生成报告失败：' + (err as Error).message;
  }

  render();
}

// 显示生成前问题预览对话框
function showPreviewDialog(
  validation: { missing: string[]; extra: string[] },
  excelData: ExtractedData,
  fieldPositions: FieldPosition[]
): void {
  const posMap = new Map(fieldPositions.map(p => [p.fieldName, p]));

  const extraHtml = validation.extra.length > 0 ? `
    <div class="mb-4">
      <div class="flex items-center gap-2 mb-2">
        <span class="text-lg">📋</span>
        <span class="text-sm font-medium text-amber-800">Excel 中有，但 Word 中未使用（${validation.extra.length}个）</span>
      </div>
      <div class="bg-amber-50 rounded-lg p-3">
        ${validation.extra.map(field => {
          const pos = posMap.get(field);
          const row = pos ? `第${pos.row}行` : '';
          const emptyHint = pos && !pos.hasValue ? `<span class="text-red-500 ml-2">（B${pos.row}为空）</span>` : '';
          return `<div class="text-sm text-amber-700 py-0.5 font-mono">{{${field}}}${row ? `<span class="text-xs text-gray-500 ml-2">→ ${row}${emptyHint}</span>` : ''}</div>`;
        }).join('')}
      </div>
    </div>
  ` : '';

  const missingHtml = validation.missing.length > 0 ? `
    <div class="mb-4">
      <div class="flex items-center gap-2 mb-2">
        <span class="text-lg">📋</span>
        <span class="text-sm font-medium text-blue-800">Word 中有，但 Excel 中未提供（${validation.missing.length}个）</span>
      </div>
      <div class="bg-blue-50 rounded-lg p-3">
        ${validation.missing.map(field => `
          <div class="text-sm text-blue-700 py-0.5 font-mono">{{${field}}}</div>
        `).join('')}
      </div>
    </div>
  ` : '';

  const dialogHtml = `
    <div id="preview-dialog" class="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onclick="closePreviewDialog(event)">
      <div class="bg-white rounded-2xl shadow-2xl max-w-xl w-full mx-4 max-h-[80vh] overflow-hidden" onclick="event.stopPropagation()">
        <!-- Header -->
        <div class="bg-amber-500 text-white px-6 py-4 flex items-center justify-between">
          <div class="flex items-center gap-3">
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L4.082 16.5c-.77.833.192 2.5 1.732 2.5z"></path>
            </svg>
            <span class="font-semibold text-lg">生成前检查发现问题</span>
          </div>
          <button onclick="closePreviewDialog()" class="hover:bg-amber-600 rounded-lg p-1 transition-colors">
            <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path>
            </svg>
          </button>
        </div>

        <!-- Content -->
        <div class="p-6 overflow-y-auto max-h-[55vh]">
          ${extraHtml}
          ${missingHtml}

          <div class="bg-gray-50 rounded-lg p-3 border border-gray-200">
            <div class="flex items-start gap-2">
              <span class="text-base shrink-0">💡</span>
              <p class="text-xs text-gray-600">
                Word 中缺失的字段会保持占位符原始样式，最终报告中显示为 <code class="bg-gray-100 px-1 rounded">&#123;&#123;字段名&#125;&#125;</code>
              </p>
            </div>
          </div>
        </div>

        <!-- Footer -->
        <div class="px-6 py-4 bg-gray-50 border-t border-gray-200 flex justify-end gap-3">
          <button onclick="closePreviewDialog()" class="px-6 py-2 bg-white border border-gray-300 hover:bg-gray-100 text-gray-700 rounded-lg font-medium transition-colors">
            返回修改
          </button>
          <button onclick="ignorePreviewAndGenerate()" class="px-6 py-2 bg-amber-500 hover:bg-amber-600 text-white rounded-lg font-medium transition-colors">
            忽略问题，直接生成
          </button>
        </div>
      </div>
    </div>
  `;

  const container = document.createElement('div');
  container.innerHTML = dialogHtml;
  document.body.appendChild(container);

  (window as unknown as Record<string, unknown>).closePreviewDialog = function(event?: MouseEvent) {
    if (event && (event.target as HTMLElement).closest('#preview-dialog .bg-white')) return;
    const dialog = document.getElementById('preview-dialog');
    if (dialog) dialog.parentElement?.remove();
  };

  (window as unknown as Record<string, unknown>).ignorePreviewAndGenerate = async function() {
    const dialog = document.getElementById('preview-dialog');
    if (dialog) dialog.parentElement?.remove();
    await doGenerate();
  };
}

// 显示结果对话框
function showResultDialog(
  filename: string,
  validation: { missing: string[]; extra: string[] },
  excelData: ExtractedData,
  placeholders: string[],
  tableData?: TableData[]
): void {
  const matched = placeholders.filter(p => excelData[p] !== undefined);

  const dialogHtml = `
    <div id="result-dialog" class="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onclick="closeResultDialog(event)">
      <div class="bg-white rounded-2xl shadow-2xl max-w-lg w-full mx-4 max-h-[80vh] overflow-hidden" onclick="event.stopPropagation()">
        <!-- Header -->
        <div class="bg-green-500 text-white px-6 py-4 flex items-center justify-between">
          <div class="flex items-center gap-3">
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path>
            </svg>
            <span class="font-semibold text-lg">报告生成成功</span>
          </div>
          <button onclick="closeResultDialog()" class="hover:bg-green-600 rounded-lg p-1 transition-colors">
            <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path>
            </svg>
          </button>
        </div>

        <!-- Content -->
        <div class="p-6 overflow-y-auto max-h-[60vh]">
          <div class="mb-4">
            <div class="text-sm text-gray-500 mb-1">已生成文件</div>
            <div class="font-mono text-sm bg-gray-100 rounded-lg px-3 py-2">${filename}</div>
          </div>

          <div class="grid grid-cols-4 gap-3 mb-4">
            <div class="bg-green-50 rounded-lg p-3 text-center">
              <div class="text-2xl font-bold text-green-600">${matched.length}</div>
              <div class="text-xs text-green-700">成功替换</div>
            </div>
            <div class="bg-amber-50 rounded-lg p-3 text-center">
              <div class="text-2xl font-bold text-amber-600">${validation.extra.length}</div>
              <div class="text-xs text-amber-700">Excel 多余</div>
            </div>
            <div class="bg-blue-50 rounded-lg p-3 text-center">
              <div class="text-2xl font-bold text-blue-600">${validation.missing.length}</div>
              <div class="text-xs text-blue-700">Word 缺失</div>
            </div>
            ${tableData && tableData.length > 0 ? `
            <div class="bg-purple-50 rounded-lg p-3 text-center">
              <div class="text-2xl font-bold text-purple-600">${tableData.length}</div>
              <div class="text-xs text-purple-700">表格填充</div>
            </div>
            ` : ''}
          </div>

          ${validation.extra.length > 0 ? `
            <div class="mb-4">
              <div class="flex items-center gap-2 mb-2">
                <span class="w-2 h-2 bg-amber-500 rounded-full"></span>
                <span class="text-sm font-medium text-amber-800">Excel 中有，但 Word 模板中未使用</span>
              </div>
              <div class="bg-amber-50 rounded-lg p-3">
                <div class="text-xs text-amber-700 mb-2">以下字段已从 Excel 读取，但 Word 模板中没有对应的占位符：</div>
                <div class="flex flex-wrap gap-2">
                  ${validation.extra.map(field => `
                    <span class="px-2 py-1 bg-amber-100 text-amber-700 rounded text-xs font-mono">{{${field}}}</span>
                  `).join('')}
                </div>
              </div>
            </div>
          ` : ''}

          ${validation.missing.length > 0 ? `
            <div class="mb-4">
              <div class="flex items-center gap-2 mb-2">
                <span class="w-2 h-2 bg-blue-500 rounded-full"></span>
                <span class="text-sm font-medium text-blue-800">Word 模板中有，但 Excel 中未提供</span>
              </div>
              <div class="bg-blue-50 rounded-lg p-3">
                <div class="text-xs text-blue-700 mb-2">以下占位符在 Word 中存在，但 Excel 中没有提供对应的值：</div>
                <div class="flex flex-wrap gap-2">
                  ${validation.missing.map(field => `
                    <span class="px-2 py-1 bg-blue-100 text-blue-700 rounded text-xs font-mono">{{${field}}}</span>
                  `).join('')}
                </div>
                <div class="text-xs text-blue-600 mt-2">提示：这些占位符在生成的报告中会保持原样</div>
              </div>
            </div>
          ` : ''}

          ${validation.extra.length === 0 && validation.missing.length === 0 ? `
            <div class="bg-green-50 rounded-lg p-4 text-center">
              <svg class="w-8 h-8 text-green-500 mx-auto mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path>
              </svg>
              <p class="text-green-700 font-medium">完美匹配！</p>
              <p class="text-sm text-green-600">Excel 和 Word 的字段完全对应</p>
            </div>
          ` : ''}
        </div>

        <div class="px-6 py-4 bg-gray-50 border-t border-gray-200">
          <button onclick="closeResultDialog()" class="w-full py-2 bg-gray-900 hover:bg-gray-800 text-white rounded-lg font-medium transition-colors">
            知道了
          </button>
        </div>
      </div>
    </div>
  `;

  const container = document.createElement('div');
  container.innerHTML = dialogHtml;
  document.body.appendChild(container);

  (window as unknown as Record<string, unknown>).closeResultDialog = function(event?: MouseEvent) {
    if (!event || !(event.target as HTMLElement).closest('#result-dialog .bg-white')) {
      const dialog = document.getElementById('result-dialog');
      if (dialog) {
        dialog.parentElement?.remove();
      }
    }
  };
}

// 生成示例模板
async function downloadSampleTemplate(): Promise<void> {
  try {
    const sampleFields = ['项目名称', '估价师', '估价时间', '总价', '建筑面积', '地址'];
    const doc = createSampleTemplate(sampleFields);
    const arrayBuffer = await Packer.toArrayBuffer(doc);
    downloadDocx(arrayBuffer, '估价报告模板.docx');
  } catch (err) {
    console.error('生成示例模板失败:', err);
    alert('生成示例模板失败：' + (err as Error).message);
  }
}

// 切换说明折叠状态
function toggleInstructions(): void {
  state.instructionCollapsed = !state.instructionCollapsed;
  render();
}

// 渲染 UI
function render(): void {
  const app = document.getElementById('app');
  if (!app) return;

  const statusClass = {
    idle: 'border-gray-200',
    loading: 'border-blue-500 bg-blue-50',
    success: 'border-green-500 bg-green-50',
    error: 'border-red-500 bg-red-50'
  }[state.status];

  const statusTextClass = {
    idle: 'text-gray-500',
    loading: 'text-blue-600',
    success: 'text-green-600',
    error: 'text-red-600'
  }[state.status];

  app.innerHTML = `
    <div class="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">
      <!-- Header -->
      <header class="bg-white shadow-sm border-b border-gray-200">
        <div class="max-w-6xl mx-auto px-6 py-4">
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-3">
              <div class="w-10 h-10 bg-blue-600 rounded-lg flex items-center justify-center">
                <svg class="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                </svg>
              </div>
              <div>
                <h1 class="text-xl font-bold text-gray-900">估价报告生成器</h1>
                <p class="text-sm text-gray-500">Excel 数据一键填充 Word 模板</p>
              </div>
            </div>
          </div>
        </div>
      </header>

      <!-- Main Content -->
      <main class="max-w-6xl mx-auto px-6 py-8">

        <!-- 使用说明 -->
        <div class="bg-white rounded-xl shadow-sm border border-gray-200 mb-8">
          <button onclick="toggleInstructions()" class="w-full px-6 py-4 flex items-center justify-between hover:bg-gray-50 transition-colors">
            <div class="flex items-center gap-3">
              <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
              </svg>
              <span class="font-semibold text-gray-900">使用说明</span>
            </div>
            <svg class="w-5 h-5 text-gray-400 transition-transform ${state.instructionCollapsed ? '' : 'rotate-180'}" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"></path>
            </svg>
          </button>

          ${state.instructionCollapsed ? `
            <div class="px-6 pb-4">
              <div class="flex flex-wrap gap-4 text-sm text-gray-600">
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 bg-green-100 text-green-600 rounded flex items-center justify-center text-xs font-bold">1</span>
                  <span>准备 Excel 数据文件（字段 + 表格）</span>
                </div>
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 bg-blue-100 text-blue-600 rounded flex items-center justify-center text-xs font-bold">2</span>
                  <span>准备 Word 模板文件</span>
                </div>
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 bg-purple-100 text-purple-600 rounded flex items-center justify-center text-xs font-bold">3</span>
                  <span>上传文件并生成报告</span>
                </div>
              </div>
            </div>
          ` : `
            <div class="px-6 pb-6 border-t border-gray-100">
              <div class="pt-6 mb-6">
                <div class="flex items-center gap-3 mb-3">
                  <span class="flex items-center justify-center w-8 h-8 bg-green-500 text-white rounded-full text-sm font-bold">1</span>
                  <h3 class="font-semibold text-gray-900">准备 Excel 数据文件</h3>
                </div>
                <div class="ml-11 space-y-3">
                  <p class="text-sm text-gray-600">在你的 Excel 文件中，新增一个专门的数据表（Sheet）：</p>
                  <div class="bg-blue-50 border border-blue-200 rounded-lg p-4">
                    <div class="text-sm font-medium text-blue-800 mb-2">Sheet 命名要求：</div>
                    <ul class="text-sm text-blue-700 space-y-1">
                      <li>• 可以随意命名（如"数据"、"report"、"报告"等）</li>
                      <li>• 放在任意位置都可以（第一个、最后一个、或者中间都行）</li>
                      <li>• 推荐放在第一个 Sheet，方便查找</li>
                    </ul>
                  </div>
                  <div class="bg-gray-50 rounded-lg p-4">
                    <div class="text-sm font-medium text-gray-700 mb-2">数据表格式（必须按此格式填写）：</div>
                    <table class="w-full text-sm border-collapse">
                      <thead>
                        <tr class="bg-gray-100">
                          <th class="border border-gray-300 px-4 py-2 text-left">A列（必填）</th>
                          <th class="border border-gray-300 px-4 py-2 text-left">B列（必填）</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr><td class="border border-gray-300 px-4 py-2 font-mono">&#123;&#123;项目名称&#125;&#125;</td><td class="border border-gray-300 px-4 py-2">阳光小区</td></tr>
                        <tr><td class="border border-gray-300 px-4 py-2 font-mono">&#123;&#123;估价师&#125;&#125;</td><td class="border border-gray-300 px-4 py-2">张三</td></tr>
                        <tr><td class="border border-gray-300 px-4 py-2 font-mono">&#123;&#123;总价&#125;&#125;</td><td class="border border-gray-300 px-4 py-2">5,000,000</td></tr>
                      </tbody>
                    </table>
                    <p class="text-xs text-gray-500 mt-2">注意：A列填写字段名（格式：<code class="bg-gray-100 px-1">&#123;&#123;字段名&#125;&#125;</code>），B列填写对应的值</p>
                  </div>
                  <div class="bg-indigo-50 border border-indigo-200 rounded-lg p-4">
                    <div class="text-sm font-medium text-indigo-800 mb-2">进阶：表格区域填充（支持多表横向/纵向排列）</div>
                    <p class="text-sm text-indigo-700 mb-2">用 <code class="bg-indigo-100 px-1">&#123;&#123;table:表名&#125;&#125;</code> 在任意单元格标记表格区域：</p>
                    <ul class="text-sm text-indigo-700 space-y-1 mb-3">
                      <li>• 标记所在列 → 行标签（如因素名称）</li>
                      <li>• 标记同行右侧列 → 列标题（如案例A、案例B）</li>
                      <li>• 标记下方行 → 数据值</li>
                      <li>• 同一 Sheet 可放多个表（横向并排或纵向堆叠）</li>
                    </ul>
                    <p class="text-sm text-indigo-700 mt-2">示例（横向并排三张表）：</p>
                    <table class="w-full text-sm border-collapse mb-2">
                      <thead>
                        <tr class="bg-indigo-100">
                          <th class="border border-indigo-300 px-3 py-1.5 text-left text-indigo-800 w-12">行</th>
                          <th class="border border-indigo-300 px-3 py-1.5 text-left text-indigo-800">A列</th>
                          <th class="border border-indigo-300 px-3 py-1.5 text-left text-indigo-800">B列</th>
                          <th class="border border-indigo-300 px-3 py-1.5 text-left text-indigo-800">C列</th>
                          <th class="border border-indigo-300 px-3 py-1.5 text-left text-indigo-800">D列</th>
                          <th class="border border-indigo-300 px-3 py-1.5 text-left text-indigo-800">E列</th>
                          <th class="border border-indigo-300 px-3 py-1.5 text-left text-indigo-800">F列</th>
                          <th class="border border-indigo-300 px-3 py-1.5 text-left text-indigo-800">G列</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr class="bg-indigo-50">
                          <td class="border border-indigo-200 px-3 py-1 text-xs text-gray-400">1</td>
                          <td class="border border-indigo-200 px-3 py-1 font-mono text-indigo-600">&#123;&#123;table:因素描述表&#125;&#125;</td>
                          <td class="border border-indigo-200 px-3 py-1"></td>
                          <td class="border border-indigo-200 px-3 py-1"></td>
                          <td class="border border-indigo-200 px-3 py-1"></td>
                          <td class="border border-indigo-200 px-3 py-1 font-mono text-indigo-600">&#123;&#123;table:打分表&#125;&#125;</td>
                          <td class="border border-indigo-200 px-3 py-1"></td>
                          <td class="border border-indigo-200 px-3 py-1"></td>
                        </tr>
                        <tr>
                          <td class="border border-indigo-200 px-3 py-1 text-xs text-gray-400">2</td>
                          <td class="border border-indigo-200 px-3 py-1">因素</td>
                          <td class="border border-indigo-200 px-3 py-1">案例A</td>
                          <td class="border border-indigo-200 px-3 py-1">案例B</td>
                          <td class="border border-indigo-200 px-3 py-1">案例C</td>
                          <td class="border border-indigo-200 px-3 py-1">因素</td>
                          <td class="border border-indigo-200 px-3 py-1">案例A</td>
                          <td class="border border-indigo-200 px-3 py-1">案例B</td>
                        </tr>
                        <tr>
                          <td class="border border-indigo-200 px-3 py-1 text-xs text-gray-400">3</td>
                          <td class="border border-indigo-200 px-3 py-1">位置</td>
                          <td class="border border-indigo-200 px-3 py-1">好</td>
                          <td class="border border-indigo-200 px-3 py-1">一般</td>
                          <td class="border border-indigo-200 px-3 py-1">差</td>
                          <td class="border border-indigo-200 px-3 py-1">位置</td>
                          <td class="border border-indigo-200 px-3 py-1">100</td>
                          <td class="border border-indigo-200 px-3 py-1">85</td>
                        </tr>
                        <tr>
                          <td class="border border-indigo-200 px-3 py-1 text-xs text-gray-400">4</td>
                          <td class="border border-indigo-200 px-3 py-1">级别</td>
                          <td class="border border-indigo-200 px-3 py-1">一级</td>
                          <td class="border border-indigo-200 px-3 py-1">三级</td>
                          <td class="border border-indigo-200 px-3 py-1">二级</td>
                          <td class="border border-indigo-200 px-3 py-1">级别</td>
                          <td class="border border-indigo-200 px-3 py-1">95</td>
                          <td class="border border-indigo-200 px-3 py-1">80</td>
                        </tr>
                      </tbody>
                    </table>
                    <p class="text-xs text-indigo-500 mt-1">第1行：表格名称（标记行），第2行：列标题，第3行起：数据。Word 中用 <code class="bg-indigo-100 px-1">&#123;&#123;table:表名&#125;&#125;</code> 单独一段标记插入位置。</p>
                  </div>
                </div>
              </div>
              <div class="mb-6">
                <div class="flex items-center gap-3 mb-3">
                  <span class="flex items-center justify-center w-8 h-8 bg-blue-500 text-white rounded-full text-sm font-bold">2</span>
                  <h3 class="font-semibold text-gray-900">准备 Word 模板文件</h3>
                </div>
                <div class="ml-11 space-y-3">
                  <p class="text-sm text-gray-600">在你的 Word 文档中，用占位符标记需要替换的位置：</p>
                  <div class="bg-gray-50 rounded-lg p-4">
                    <div class="text-sm font-medium text-gray-700 mb-2">Word 模板格式：</div>
                    <pre class="text-sm text-gray-700 bg-white border border-gray-200 rounded p-3 whitespace-pre-line">本报告由 &#123;&#123;估价师&#125;&#125; 出具
项目名称：&#123;&#123;项目名称&#125;&#125;
评估总价：&#123;&#123;总价&#125;&#125;</pre>
                    <p class="text-xs text-gray-500 mt-2">注意：Word 中的字段名必须与 Excel 中 A 列的字段名完全一致</p>
                  </div>
                </div>
              </div>
              <div class="mb-6">
                <div class="flex items-center gap-3 mb-3">
                  <span class="flex items-center justify-center w-8 h-8 bg-purple-500 text-white rounded-full text-sm font-bold">3</span>
                  <h3 class="font-semibold text-gray-900">上传文件并生成报告</h3>
                </div>
                <div class="ml-11">
                  <p class="text-sm text-gray-600">上传 Excel 和 Word 文件后，点击<strong>"生成报告"</strong>按钮即可获得填充好的 Word 文档。</p>
                </div>
              </div>
              <div class="pt-4 border-t border-gray-200">
                <div class="bg-amber-50 border border-amber-200 rounded-lg p-4">
                  <ul class="text-sm text-amber-700 space-y-1">
                    <li class="font-medium mb-1">注意事项</li>
                    <li>• Excel 中每个字段名只能出现一次，不能重复</li>
                    <li>• Word 中缺少的字段会自动跳过，不会影响生成</li>
                    <li>• Excel 中多余的字段（Word 中没有对应占位符）也会自动忽略</li>
                    <li>• 数据处理完全在本地完成，不会上传到任何服务器</li>
                  </ul>
                </div>
              </div>
            </div>
          `}
        </div>

        <!-- 版本更新提示 -->
        <div id="update-notice"></div>

        <!-- 文件上传区 -->
        <div class="grid md:grid-cols-2 gap-6 mb-8">
          <!-- Excel 上传 -->
          <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
            <h3 class="font-semibold text-gray-900 mb-4 flex items-center gap-2">
              <div class="w-8 h-8 bg-green-100 rounded-lg flex items-center justify-center">
                <svg class="w-5 h-5 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                </svg>
              </div>
              Excel 数据文件
            </h3>
            <label class="block">
              <div class="border-2 border-dashed ${state.excelFile ? 'border-green-400 bg-green-50' : 'border-gray-300 hover:border-blue-400'} rounded-xl p-8 text-center cursor-pointer transition-all">
                <input type="file" accept=".xlsx,.xls" onchange="handleExcelUpload(event)" class="hidden" />
                ${state.excelFile
                  ? `<div class="text-green-600">
                       <svg class="w-12 h-12 mx-auto mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                         <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path>
                       </svg>
                       <p class="font-medium">${state.excelFile.name}</p>
                       <p class="text-sm text-gray-500 mt-1">点击重新选择</p>
                     </div>`
                  : `<div class="text-gray-400">
                       <svg class="w-12 h-12 mx-auto mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                         <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path>
                       </svg>
                       <p class="font-medium">点击上传 Excel 文件</p>
                       <p class="text-sm mt-1">支持 .xlsx, .xls 格式</p>
                     </div>`
                }
              </div>
            </label>
          </div>

          <!-- Word 上传 -->
          <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
            <h3 class="font-semibold text-gray-900 mb-4 flex items-center gap-2">
              <div class="w-8 h-8 bg-blue-100 rounded-lg flex items-center justify-center">
                <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                </svg>
              </div>
              Word 模板文件
            </h3>
            <label class="block">
              <div class="border-2 border-dashed ${state.wordFile ? 'border-green-400 bg-green-50' : 'border-gray-300 hover:border-blue-400'} rounded-xl p-8 text-center cursor-pointer transition-all">
                <input type="file" accept=".docx" onchange="handleWordUpload(event)" class="hidden" />
                ${state.wordFile
                  ? `<div class="text-green-600">
                       <svg class="w-12 h-12 mx-auto mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                         <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path>
                       </svg>
                       <p class="font-medium">${state.wordFile.name}</p>
                       <p class="text-sm text-gray-500 mt-1">点击重新选择</p>
                     </div>`
                  : `<div class="text-gray-400">
                       <svg class="w-12 h-12 mx-auto mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                         <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path>
                       </svg>
                       <p class="font-medium">点击上传 Word 模板</p>
                       <p class="text-sm mt-1">支持 .docx 格式</p>
                     </div>`
                }
              </div>
            </label>
          </div>
        </div>

        <!-- 状态消息 -->
        ${state.message ? `
          <div class="mb-8 p-4 rounded-xl border ${statusClass}">
            <p class="text-sm ${statusTextClass} flex items-center gap-2">
              ${state.status === 'loading' ? `
                <svg class="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24">
                  <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
                  <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
              ` : state.status === 'success' ? `
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path>
                </svg>
              ` : state.status === 'error' ? `
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path>
                </svg>
              ` : ''}
              <span class="whitespace-pre-line">${state.message}</span>
            </p>
          </div>
        ` : ''}

        <!-- 数据预览 -->
        ${Object.keys(state.excelData).length > 0 ? `
          <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-6 mb-8">
            <h3 class="font-semibold text-gray-900 mb-4">数据预览（共 ${Object.keys(state.excelData).length} 个字段）</h3>
            <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
              ${Object.entries(state.excelData).map(([key, value]) => `
                <div class="bg-gray-50 rounded-lg p-3 border border-gray-100">
                  <div class="text-xs text-gray-500 mb-1">{{${key}}}</div>
                  <div class="text-sm font-medium text-gray-900 truncate" title="${value}">${value || '<空>'}</div>
                </div>
              `).join('')}
            </div>
          </div>
        ` : ''}

        <!-- 表格区域预览 -->
        ${state.tableData.length > 0 ? `
          <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-6 mb-8">
            <div class="flex items-center justify-between mb-4">
              <h3 class="font-semibold text-gray-900">表格区域预览（共 ${state.tableData.length} 个）</h3>
              <div class="flex items-center gap-3 text-sm">
                <span class="text-gray-500">表格样式：</span>
                <select onchange="state.tableFontFamily=this.value;render()" class="border border-gray-300 rounded-lg px-2 py-1 text-sm bg-white">
                  <option value="宋体" ${state.tableFontFamily === '宋体' ? 'selected' : ''}>宋体</option>
                  <option value="黑体" ${state.tableFontFamily === '黑体' ? 'selected' : ''}>黑体</option>
                  <option value="微软雅黑" ${state.tableFontFamily === '微软雅黑' ? 'selected' : ''}>微软雅黑</option>
                </select>
                <select onchange="state.tableFontSize=Number(this.value);render()" class="border border-gray-300 rounded-lg px-2 py-1 text-sm bg-white">
                  <option value="9" ${state.tableFontSize === 9 ? 'selected' : ''}>小五 (9pt)</option>
                  <option value="10" ${state.tableFontSize === 10 ? 'selected' : ''}>五号 (10pt)</option>
                  <option value="10.5" ${state.tableFontSize === 10.5 ? 'selected' : ''}>五号 (10.5pt)</option>
                  <option value="12" ${state.tableFontSize === 12 ? 'selected' : ''}>小四 (12pt)</option>
                </select>
              </div>
            </div>
            ${state.tableData.map(t => {
              const dataRows = t.data;
              const colCount = dataRows.length > 0 ? dataRows[0].length : 0;
              const headerRow = dataRows.length > 0 ? dataRows[0] : [];
              const bodyRows = dataRows.slice(1);
              return `
              <div class="mb-5 last:mb-0 border border-gray-200 rounded-lg overflow-hidden">
                <div class="bg-blue-50 px-4 py-2 border-b border-gray-200 flex items-center gap-2">
                  <span class="font-medium text-blue-800">{{table:${t.tableName}}}</span>
                  <span class="text-xs text-blue-600">${colCount} 列 × ${dataRows.length} 行</span>
                </div>
                <div class="overflow-x-auto">
                  <table class="w-full text-sm border-collapse">
                    <thead>
                      <tr class="bg-gray-100">
                        ${headerRow.map(h => `
                          <th class="border border-gray-300 px-3 py-1.5 text-left font-medium text-gray-700">${h || '<空>'}</th>
                        `).join('')}
                      </tr>
                    </thead>
                    <tbody>
                      ${bodyRows.slice(0, 5).map(row => `
                        <tr>
                          ${row.map(cell => `
                            <td class="border border-gray-200 px-3 py-1 text-gray-600">${cell || '<空>'}</td>
                          `).join('')}
                        </tr>
                      `).join('')}
                    </tbody>
                  </table>
                </div>
                ${bodyRows.length > 5 ? `
                  <div class="px-4 py-2 bg-gray-50 text-xs text-gray-500 text-center">
                    仅显示前 5 行数据，共 ${bodyRows.length} 行
                  </div>
                ` : ''}
              </div>
            `}).join('')}
          </div>
        ` : ''}

        <!-- 占位符预览 -->
        ${state.placeholders.length > 0 ? `
          <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-6 mb-8">
            <h3 class="font-semibold text-gray-900 mb-4">Word 模板占位符（共 ${state.placeholders.length} 个）</h3>
            <div class="flex flex-wrap gap-2">
              ${state.placeholders.map(p => `
                <span class="px-3 py-1.5 bg-yellow-50 text-yellow-700 rounded-full text-sm font-mono border border-yellow-200">
                  {{${p}}}
                </span>
              `).join('')}
            </div>
          </div>
        ` : ''}

        <!-- 生成按钮 -->
        <div class="flex justify-center">
          <button
            onclick="generateReport()"
            ${!state.excelFile || !state.wordFile || !state.wordBuffer || (Object.keys(state.excelData).length === 0 && state.tableData.length === 0) || state.status === 'loading' ? 'disabled' : ''}
            class="px-8 py-4 bg-blue-600 hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed text-white rounded-xl font-semibold text-lg shadow-lg shadow-blue-600/30 transition-all flex items-center gap-3"
          >
            ${state.status === 'loading' ? `
              <svg class="animate-spin w-5 h-5" fill="none" viewBox="0 0 24 24">
                <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
                <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
              </svg>
              生成中...
            ` : `
              <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"></path>
              </svg>
              生成报告
            `}
          </button>
        </div>

        <div class="mt-12 text-center text-sm text-gray-400">
          <p class="text-xs mb-2">
            <svg class="w-3.5 h-3.5 inline-block mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 15v2m0 0v2m0-2h2m-2 0H10m9.364-6.364a9 9 0 11-12.728 0 9 9 0 0112.728 0zM12 8v4"></path>
            </svg>
            数据处理完全在本地完成，不会上传到任何服务器
          </p>
          <p class="text-xs">
            出现报错，截图发给
            <span class="relative inline-block group cursor-pointer text-blue-500 hover:text-blue-600 font-medium underline decoration-dotted underline-offset-2 select-none">
              gaoly3
              <!-- 二维码弹窗 -->
              <span class="absolute bottom-full left-1/2 -translate-x-1/2 mb-3 hidden group-hover:block animate-fade-in z-50">
                <span class="block bg-white rounded-xl shadow-2xl border border-gray-200 p-3 w-48">
                  <span class="block w-full aspect-square bg-gray-100 rounded-lg overflow-hidden">
                    <img src="./wechat-qr.png" alt="微信二维码" class="w-full h-full object-cover" onerror="this.parentElement.innerHTML='<div class=\\'w-full h-full flex items-center justify-center text-gray-400 text-xs\\'>请将二维码图片<br/>放入 public/wechat-qr.png</div>'" />
                  </span>
                  <span class="block text-xs text-gray-400 mt-2 text-center">微信扫码联系</span>
                </span>
                <!-- 小三角箭头 -->
                <span class="absolute top-full left-1/2 -translate-x-1/2 -mt-[5px] w-3 h-3 bg-white border-r border-b border-gray-200 rotate-45"></span>
              </span>
            </span>
          </p>
          <p class="text-xs mt-2">估价报告生成器 v${APP_VERSION}</p>
        </div>
      </main>
    </div>
  `;
}

// 初始化
export function initApp(): void {
  (window as unknown as Record<string, unknown>).handleExcelUpload = handleExcelUpload;
  (window as unknown as Record<string, unknown>).handleWordUpload = handleWordUpload;
  (window as unknown as Record<string, unknown>).generateReport = generateReport;
  (window as unknown as Record<string, unknown>).downloadSampleTemplate = downloadSampleTemplate;
  (window as unknown as Record<string, unknown>).toggleInstructions = toggleInstructions;

  checkForUpdates();
  render();
}
