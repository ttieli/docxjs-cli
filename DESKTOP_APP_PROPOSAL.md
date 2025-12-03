# 桌面应用程序化方案 (Desktop App Conversion Plan)

## 1. 目标
将当前的 `docxjs-cli` Web 工具打包为独立的桌面安装包：
*   **Windows**: `.exe` / `.msi` (支持安装向导)
*   **macOS**: `.dmg` / `.app` (拖拽安装)

## 2. 核心挑战与解决方案

### 挑战 A: 环境依赖 (最大的难点)
当前系统依赖用户电脑上必须安装有 **Node.js** 和 **Python 3 (带 python-docx 库)**。普通桌面用户通常不具备这些环境。

### 解决方案: "双重打包" 策略

我们采用 **Electron** + **PyInstaller** 的组合方案：

1.  **Node.js 部分 (Electron)**: 
    Electron 本身就是一个自带 Node.js 运行时和 Chromium 浏览器的容器。它能完美运行我们现有的 `lib/core.js` 和 `markdown-it` 等逻辑，无需用户安装 Node.js。

2.  **Python 部分 (PyInstaller)**:
    使用 `PyInstaller` 将 `style_extractor.py` 及其所有依赖（python-docx, lxml 等）“冻结”成一个独立的可执行文件（`extractor.exe` 或 mac 二进制文件）。Electron 运行时直接调用这个可执行文件。

## 3. “三位一体”代码复用策略 (新增)

为了**在保持现有 CLI/Web 模式不受影响的前提下**增加桌面端支持，我们需要采用**“核心逻辑抽离 + 多入口适配”**的架构。

### 3.1 共享核心层 (Shared Core)
保持 `lib/` 目录下的代码为**纯逻辑**，不依赖任何特定的运行环境（CLI/Web/Electron）。

*   **`lib/core.js`**: 保持不变。接收 Markdown + Config，返回 Buffer。纯函数。
*   **`lib/template-manager.js`**: 保持不变。文件读写逻辑在 Electron 和 Node 中通用。
*   **`style_extractor.py`**: 保持不变。作为被调用的外部工具。

### 3.2 Python 调用层适配 (Adapter Pattern)
我们需要一个智能的 Python 调用器 `lib/python-bridge.js`，它能根据环境自动选择调用方式：

*   **CLI/Web 模式**: 寻找系统路径或 venv 中的 `python3 style_extractor.py`。
*   **Electron 模式**: 寻找打包在应用内部的 `style_extractor.exe` (或二进制文件)。

**实现伪代码：**
```javascript
// lib/python-bridge.js
function getExtractorPath() {
  if (process.versions.electron) {
    // Electron 模式：调用打包好的二进制
    return path.join(process.resourcesPath, 'bin', 'style_extractor');
  } else {
    // CLI/Web 模式：调用源码脚本
    return 'python3 ' + path.join(__dirname, '../style_extractor.py');
  }
}
```

### 3.3 前端适配 (Frontend Adapter)
`public/index.html` 需要进行微调，以同时支持浏览器环境和 Electron 环境。

*   **浏览器模式**: AJAX 请求 -> `fetch('/api/convert')`。
*   **Electron 模式**: IPC 通信 -> `ipcRenderer.invoke('convert', ...)`。

我们可以在 `index.html` 中注入一个垫片脚本 `renderer.js`：
```javascript
const isElectron = !!window.ipcRenderer;

async function callConvert(data) {
  if (isElectron) {
    return await window.ipcRenderer.invoke('convert', data);
  } else {
    const res = await fetch('/api/convert', ...);
    return await res.blob();
  }
}
```

### 3.4 项目结构调整建议

```
/
├── bin/                # CLI 入口 (保持不变)
├── lib/                # 共享核心逻辑 (保持不变)
├── public/             # 共享前端资源 (保持不变)
├── server/             # Web 模式入口 (Express App)
├── electron/           # [新增] Electron 模式入口
│   ├── main.js         # Electron 主进程
│   └── preload.js      # 安全桥接
├── style_extractor.py  # 共享 Python 脚本
└── package.json
```

## 4. 实施路线图 (Roadmap)

### 第一阶段：Python 独立化与桥接
目标：让代码能智能识别是调用 `.py` 脚本还是 `.exe` 文件。

1.  安装打包工具：`pip install pyinstaller`
2.  **新增** `lib/python-bridge.js`：封装 Python 调用的路径判断逻辑。
3.  **修改** `server/app.js` 和 `bin/cli.js`：将直接写死的 `exec('python3 ...')` 替换为调用 `lib/python-bridge.js`。
    *   *此时 CLI 和 Web 依然正常工作，只是调用方式变灵活了。*

### 第二阶段：引入 Electron 壳
目标：增加一个新的启动方式，不影响旧的。

1.  安装 Electron: `npm install --save-dev electron electron-builder`
2.  创建 `electron/main.js`：
    *   它**不启动** Express 服务器。
    *   它直接引用 `lib/core.js` 和 `lib/template-manager.js`。
    *   它监听 `ipcMain.handle('convert')` 事件，并在事件处理函数中直接调用核心逻辑。
3.  **修改** `public/index.html`：增加环境检测逻辑（如上文所述），如果是 Electron 环境则走 IPC，否则走 Fetch。

### 第三阶段：打包与发布
目标：生成最终安装包。

1.  配置 `package.json` 中的 `build` 字段 (electron-builder 配置)。
2.  配置 "extraResources"，确保生成的 Python 可执行文件被复制到安装包内部。
3.  运行构建：
    *   `npm run dist:mac` -> `.dmg`
    *   `npm run dist:win` -> `.exe`

## 5. 多架构支持分析 (Windows x86 vs ARM)

Windows 生态目前分为 **x64 (Intel/AMD)** 和 **ARM64 (Snapdragon)** 两个主要阵营。

### Electron 支持
Electron 官方完美支持两种架构，可以通过配置 `electron-builder` 轻松构建出两个版本的安装包，或者一个通用的安装包。

### Python 依赖的挑战
PyInstaller 生成的可执行文件 (`.exe`) 通常与构建它的机器架构绑定。
*   **x64 机器**构建出的 `extractor.exe` 只能在 x64 系统或支持模拟的 ARM 系统上运行。
*   **ARM64 机器**构建出的 `extractor.exe` 只能在 ARM 系统上运行。

### 推荐策略：以 x64 为主，利用仿真兼容 ARM
鉴于 Windows 11 on ARM 的 **x64 仿真层 (Emulation)** 已经非常成熟且高效：

1.  **构建策略**：我们只需构建一个 **x64 架构** 的 `style_extractor.exe`。
2.  **兼容性**：
    *   **Intel/AMD 用户**：原生运行，性能最佳。
    *   **ARM 用户 (Surface Pro X 等)**：Windows 会自动模拟 x64 环境来运行这个 `.exe`。对于这种非计算密集型的脚本（仅提取样式 JSON），性能损耗几乎可以忽略不计。
3.  **如果必须追求极致原生 ARM 性能**：
    *   您需要找一台 Windows ARM 电脑（或云主机），在其上运行 `pyinstaller`，生成一个 `style_extractor-arm64.exe`。
    *   然后在 Electron 打包时，根据目标架构把对应的 `.exe` 塞进去。但通常**没必要**这么折腾。

**结论**：**优先只构建 x64 版本**。它能覆盖 99% 的 Windows 市场，并且能通过仿真完美运行在 ARM 设备上。

## 6. 优缺点评估

| 维度 | Web 版 (现状) | Electron 桌面版 |
| :--- | :--- | :--- |
| **交付形式** | 源码/NPM 包 | 安装文件 (.exe/.dmg) |
| **用户门槛** | 高 (需懂命令行、配环境) | **极低** (双击安装) |
| **文件体积** | < 5MB | **> 150MB** (含浏览器核+Python运行时) |
| **系统权限** | 受限 | **高** (可直接读写本地文件，无需上传下载) |
| **开发成本** | 现有 | 需增加约 2-3 天的封装工作量 |

## 7. 结论

采用 **“核心逻辑共享 + 环境适配器”** 的模式，您可以做到**一份代码，三端运行（CLI, Web, Desktop）**。
*   **CLI/Web** 继续作为轻量级方案，依赖用户环境。
*   **Desktop** 作为“全包”方案，内置环境，面向普通用户。
*   针对 Windows 架构，**主攻 x64** 即可满足绝大多数需求。
