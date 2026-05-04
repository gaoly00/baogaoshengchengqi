/**
 * 估价报告生成器 - 主应用
 */

import './index.css';
import { 
  extractDataFromExcel, 
  autoExtractFromExcel,
  getSheetNames,
  getCellValue,
  FieldMapping,
  ExtractedData
} from './services/excelParser';
import { 
  replacePlaceholders, 
  extractPlaceholders,
  validatePlaceholders,
  createSampleTemplate,
  downloadDocx,
  downloadBlob
} from './services/wordProcessor';
import { Packer } from 'docx';

// 应用状态
interface AppState {
  excelFile: File | null;
  wordFile: File | null;
  excelData: ExtractedData;
  placeholders: string[];
  mappings: FieldMapping[];
  sheetNames: string[];
  selectedSheet: string;
  status: 'idle' | 'loading' | 'success' | 'error';
  message: string;
  instructionCollapsed: boolean;
}

const state: AppState = {
  excelFile: null,
  wordFile: null,
  excelData: {},
  placeholders: [],
  mappings: [],
  sheetNames: [],
  selectedSheet: '',
  status: 'idle',
  message: '',
  instructionCollapsed: true
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
      state.excelData = autoExtractFromExcel(buffer);
      state.mappings = Object.keys(state.excelData).map(field => ({
        fieldName: field,
        sheetName: state.selectedSheet,
        cellAddress: 'B' + (Object.keys(state.excelData).indexOf(field) + 1)
      }));
      state.status = 'success';
      state.message = `成功读取 Excel，共 ${Object.keys(state.excelData).length} 个字段`;
    } catch (err) {
      state.status = 'error';
      state.message = '读取 Excel 文件失败：' + (err as Error).message;
    }
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
  reader.readAsArrayBuffer(file);
}

// 生成报告
async function generateReport(): Promise<void> {
  if (!state.wordFile || Object.keys(state.excelData).length === 0) {
    state.status = 'error';
    state.message = '请先上传 Excel 数据和 Word 模板';
    render();
    return;
  }

  state.status = 'loading';
  state.message = '正在生成报告...';
  render();

  try {
    // 读取 Word 文件
    const wordBuffer = await state.wordFile.arrayBuffer();
    
    // 替换占位符
    const resultBuffer = await replacePlaceholders(wordBuffer, state.excelData);
    
    // 生成文件名
    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const originalName = state.wordFile.name.replace('.docx', '');
    const outputName = `${originalName}_${timestamp}.docx`;
    
    // 下载文件
    downloadDocx(resultBuffer, outputName);
    
    // 获取差异信息用于提示
    const validation = validatePlaceholders(state.placeholders, state.excelData);
    
    state.status = 'success';
    state.message = `报告生成成功：${outputName}`;
    
    // 显示结果对话框
    showResultDialog(outputName, validation, state.excelData, state.placeholders);
    
  } catch (err) {
    state.status = 'error';
    state.message = '生成报告失败：' + (err as Error).message;
  }
  
  render();
}

// 显示结果对话框
function showResultDialog(
  filename: string, 
  validation: { missing: string[]; extra: string[] },
  excelData: ExtractedData,
  placeholders: string[]
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
          <!-- 文件名 -->
          <div class="mb-4">
            <div class="text-sm text-gray-500 mb-1">已生成文件</div>
            <div class="font-mono text-sm bg-gray-100 rounded-lg px-3 py-2">${filename}</div>
          </div>
          
          <!-- 统计 -->
          <div class="grid grid-cols-3 gap-3 mb-4">
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
          </div>
          
          <!-- Excel 多余字段 -->
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
          
          <!-- Word 缺失字段 -->
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
        
        <!-- Footer -->
        <div class="px-6 py-4 bg-gray-50 border-t border-gray-200">
          <button onclick="closeResultDialog()" class="w-full py-2 bg-gray-900 hover:bg-gray-800 text-white rounded-lg font-medium transition-colors">
            知道了
          </button>
        </div>
      </div>
    </div>
  `;
  
  // 添加到页面
  const container = document.createElement('div');
  container.innerHTML = dialogHtml;
  document.body.appendChild(container);
  
  // 定义关闭函数
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
        
        <!-- 使用说明（可折叠） -->
        <div class="bg-white rounded-xl shadow-sm border border-gray-200 mb-8">
          <!-- 折叠 Header -->
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
          
          <!-- 简单说明（折叠状态显示） -->
          ${state.instructionCollapsed ? `
            <div class="px-6 pb-4">
              <div class="flex flex-wrap gap-4 text-sm text-gray-600">
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 bg-green-100 text-green-600 rounded flex items-center justify-center text-xs font-bold">1</span>
                  <span>Excel：A列 <code class="bg-gray-100 px-1 rounded">&#123;&#123;字段名&#125;&#125;</code>，B列值</span>
                </div>
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 bg-blue-100 text-blue-600 rounded flex items-center justify-center text-xs font-bold">2</span>
                  <span>Word：用 <code class="bg-gray-100 px-1 rounded">&#123;&#123;字段名&#125;&#125;</code> 占位</span>
                </div>
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 bg-purple-100 text-purple-600 rounded flex items-center justify-center text-xs font-bold">3</span>
                  <span>点击生成报告</span>
                </div>
              </div>
            </div>
          ` : ''}
          
          <!-- 详细说明（展开状态显示） -->
          ${!state.instructionCollapsed ? `
            <div class="px-6 pb-6 border-t border-gray-100">
              <!-- Step 1 -->
              <div class="pt-6 mb-6">
                <div class="flex items-center gap-3 mb-3">
                  <span class="flex items-center justify-center w-8 h-8 bg-green-500 text-white rounded-full text-sm font-bold">1</span>
                  <h3 class="font-semibold text-gray-900">准备 Excel 数据文件</h3>
                </div>
                <div class="ml-11 space-y-3">
                  <p class="text-sm text-gray-600">在你的 Excel 文件中，新增一个专门的数据表（Sheet）：</p>
                  <div class="bg-gray-50 rounded-lg p-4">
                    <div class="text-sm font-medium text-gray-700 mb-2">Sheet 命名要求：</div>
                    <ul class="text-sm text-gray-600 space-y-1">
                      <li>• 可以随意命名（如"数据"、"report"、"报告"等）</li>
                      <li>• 放在任意位置都可以（第一个、最后一个、或者中间都行）</li>
                      <li>• <strong>推荐放在第一个 Sheet</strong>，方便查找</li>
                    </ul>
                  </div>
                  <div class="bg-gray-50 rounded-lg p-4">
                    <div class="text-sm font-medium text-gray-700 mb-2">数据表格式（必须按此格式填写）：</div>
                    <div class="overflow-x-auto">
                      <table class="w-full text-sm border-collapse">
                        <thead>
                          <tr class="bg-gray-100">
                            <th class="border border-gray-300 px-4 py-2 text-left font-semibold text-gray-700">A列（必填）</th>
                            <th class="border border-gray-300 px-4 py-2 text-left font-semibold text-gray-700">B列（必填）</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td class="border border-gray-300 px-4 py-2 text-purple-600 font-mono">&#123;&#123;项目名称&#125;&#125;</td>
                            <td class="border border-gray-300 px-4 py-2 text-gray-900">阳光小区</td>
                          </tr>
                          <tr class="bg-white">
                            <td class="border border-gray-300 px-4 py-2 text-purple-600 font-mono">&#123;&#123;估价师&#125;&#125;</td>
                            <td class="border border-gray-300 px-4 py-2 text-gray-900">张三</td>
                          </tr>
                          <tr class="bg-white">
                            <td class="border border-gray-300 px-4 py-2 text-purple-600 font-mono">&#123;&#123;总价&#125;&#125;</td>
                            <td class="border border-gray-300 px-4 py-2 text-gray-900">5,000,000</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                    <p class="text-xs text-gray-500 mt-2">注意：A列填写字段名（格式：&#123;&#123;字段名&#125;&#125;），B列填写对应的值</p>
                  </div>
                </div>
              </div>

              <!-- Step 2 -->
              <div class="mb-6">
                <div class="flex items-center gap-3 mb-3">
                  <span class="flex items-center justify-center w-8 h-8 bg-blue-500 text-white rounded-full text-sm font-bold">2</span>
                  <h3 class="font-semibold text-gray-900">准备 Word 模板文件</h3>
                </div>
                <div class="ml-11 space-y-3">
                  <p class="text-sm text-gray-600">在你的 Word 文档中，用占位符标记需要替换的位置：</p>
                  <div class="bg-gray-50 rounded-lg p-4">
                    <div class="text-sm font-medium text-gray-700 mb-2">Word 模板格式：</div>
                    <div class="bg-white rounded border border-gray-200 p-4 font-mono text-sm space-y-2">
                      <div class="text-gray-700">本报告由 <span class="bg-yellow-100 text-yellow-700 px-1">&#123;&#123;估价师&#125;&#125;</span> 出具</div>
                      <div class="text-gray-700">项目名称：<span class="bg-yellow-100 text-yellow-700 px-1">&#123;&#123;项目名称&#125;&#125;</span></div>
                      <div class="text-gray-700">评估总价：<span class="bg-yellow-100 text-yellow-700 px-1">&#123;&#123;总价&#125;&#125;</span></div>
                    </div>
                    <p class="text-xs text-gray-500 mt-2">注意：Word 中的字段名必须与 Excel 中 A 列的字段名完全一致</p>
                  </div>
                </div>
              </div>

              <!-- Step 3 -->
              <div class="mb-6">
                <div class="flex items-center gap-3 mb-3">
                  <span class="flex items-center justify-center w-8 h-8 bg-purple-500 text-white rounded-full text-sm font-bold">3</span>
                  <h3 class="font-semibold text-gray-900">上传文件并生成报告</h3>
                </div>
                <div class="ml-11">
                  <p class="text-sm text-gray-600">上传 Excel 和 Word 文件后，点击"生成报告"按钮即可获得填充好的 Word 文档。</p>
                </div>
              </div>

              <!-- 注意事项 -->
              <div class="pt-4 border-t border-gray-200">
                <div class="bg-amber-50 border border-amber-200 rounded-lg p-4">
                  <div class="flex items-start gap-3">
                    <svg class="w-5 h-5 text-amber-600 mt-0.5 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path>
                    </svg>
                    <div>
                      <div class="font-medium text-amber-800 text-sm mb-1">注意事项</div>
                      <ul class="text-sm text-amber-700 space-y-1">
                        <li>• Excel 中每个字段名只能出现一次，不能重复</li>
                        <li>• Word 中缺少的字段会自动跳过，不会影响生成</li>
                        <li>• Excel 中多余的字段（Word 中没有对应占位符）也会自动忽略</li>
                        <li>• 数据处理完全在本地完成，不会上传到任何服务器</li>
                      </ul>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          ` : ''}
        </div>

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
          <div class="mb-8 p-4 rounded-xl border ${statusClass} transition-all">
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
              ${state.message}
            </p>
          </div>
        ` : ''}

        <!-- 数据预览 -->
        ${Object.keys(state.excelData).length > 0 ? `
          <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-6 mb-8">
            <h3 class="font-semibold text-gray-900 mb-4 flex items-center gap-2">
              <div class="w-8 h-8 bg-purple-100 rounded-lg flex items-center justify-center">
                <svg class="w-5 h-5 text-purple-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 10h16M4 14h16M4 18h16"></path>
                </svg>
              </div>
              数据预览
              <span class="ml-auto text-sm text-gray-500 font-normal">共 ${Object.keys(state.excelData).length} 个字段</span>
            </h3>
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

        <!-- 占位符预览 -->
        ${state.placeholders.length > 0 ? `
          <div class="bg-white rounded-xl shadow-sm border border-gray-200 p-6 mb-8">
            <h3 class="font-semibold text-gray-900 mb-4 flex items-center gap-2">
              <div class="w-8 h-8 bg-yellow-100 rounded-lg flex items-center justify-center">
                <svg class="w-5 h-5 text-yellow-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 20l4-16m2 16l4-16M6 9h14M4 15h14"></path>
                </svg>
              </div>
              Word 模板占位符
              <span class="ml-auto text-sm text-gray-500 font-normal">共 ${state.placeholders.length} 个</span>
            </h3>
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
            ${!state.excelFile || !state.wordFile || state.status === 'loading' ? 'disabled' : ''}
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

        <!-- 底部说明 -->
        <div class="mt-12 text-center text-sm text-gray-500">
          <p>数据处理完全在本地完成，不会上传到任何服务器</p>
        </div>
      </main>
    </div>
  `;
}

// 初始化
export function initApp(): void {
  // 将函数挂载到 window 以便 HTML onclick 调用
  (window as unknown as Record<string, unknown>).handleExcelUpload = handleExcelUpload;
  (window as unknown as Record<string, unknown>).handleWordUpload = handleWordUpload;
  (window as unknown as Record<string, unknown>).generateReport = generateReport;
  (window as unknown as Record<string, unknown>).downloadSampleTemplate = downloadSampleTemplate;
  (window as unknown as Record<string, unknown>).toggleInstructions = toggleInstructions;
  
  render();
}
