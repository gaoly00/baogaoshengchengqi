# 估价报告生成器 - 桌面端打包指南

## 环境要求

- Node.js 18+ (推荐 20+)
- npm (随 Node.js 一起安装)

## 打包步骤

1. 解压此文件夹到任意目录

2. 打开命令行（CMD 或 PowerShell），进入解压后的目录

3. 安装依赖：
   ```
   npm install
   ```

4. 打包为 Windows 桌面应用：
   ```
   npm run build:win
   ```

5. 打包完成后，在 `release/win-unpacked/` 目录下找到 `估价报告生成器.exe`

6. 双击 `估价报告生成器.exe` 即可运行

## 注意事项

- 首次 `npm install` 可能需要几分钟，因为要下载 Electron（约 80MB）
- 打包后的 `release/win-unpacked/` 文件夹包含完整的运行时，不要只复制 exe 文件
- 如果杀毒软件拦截，请添加信任

## 如果想修改代码后重新打包

1. 修改 `src/` 目录下的源代码
2. 运行 `npm run build` 重新构建前端
3. 运行 `npm run build:win` 重新打包桌面应用
