# 估价报告生成器 - 桌面端打包指南

## 方式一：使用 Electron 打包（推荐 Windows）

### 步骤 1：在本机安装 Node.js
下载并安装 Node.js (LTS 版本)：https://nodejs.org/

### 步骤 2：下载项目代码
将项目代码下载到本地

### 步骤 3：安装依赖
```bash
cd 项目目录
npm install
```

### 步骤 4：安装 Electron 打包依赖
```bash
npm install electron@33.0.0 electron-builder@25.1.8 --save-dev
```

### 步骤 5：打包 Windows 安装包
```bash
npm run build:win
```

打包完成后，在 `release` 目录会生成：
- `估价报告生成器-Setup-1.0.0.exe`（安装包）
- `估价报告生成器-1.0.0.exe`（免安装版）

### 步骤 6：分发使用
- **安装版**：双击 Setup 文件安装
- **免安装版**：直接运行 exe 文件即可

---

## 方式二：使用 Tauri 打包（更小、更快）

### 步骤 1：安装 Rust
```bash
# Windows: 从 https://rustup.rs 安装
# macOS: brew install rust
# Linux: curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
```

### 步骤 2：安装 Tauri CLI
```bash
npm install -g @tauri-apps/cli
```

### 步骤 3：初始化 Tauri
```bash
npx tauri init --ci
```

### 步骤 4：修改配置
编辑 `src-tauri/tauri.conf.json`：
```json
{
  "build": {
    "distDir": "../dist"
  }
}
```

### 步骤 5：打包
```bash
npx tauri build
```

---

## 方式三：直接使用 Web 版本

如果不需要打包成 exe，也可以直接：
1. 启动 Web 服务：`npm start`
2. 或使用 VS Code 的 Live Server 插件

---

## 文件说明

```
├── dist/                    # 构建好的 Web 应用
│   ├── index.html
│   └── assets/
├── electron/
│   └── main.js             # Electron 入口
├── package.desktop.json     # 桌面端 package.json
└── README-DESKTOP.md       # 本文件
```

---

## 常见问题

### Q: 打包后文件太大？
Electron 打包后约 150MB，Tauri 打包后约 10MB。推荐使用 Tauri。

### Q: 杀毒软件报毒？
Electron 打包的应用有时会被误报，可以添加数字签名解决。

### Q: 如何自定义图标？
替换 `electron/` 目录下的图标文件：
- Windows: `icon.ico`
- macOS: `icon.icns`
- Linux: `icon.png`
