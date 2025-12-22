# HTML 预览与导出 PNG/PDF 方案设计

目标：在现有（Electron + Web）应用中，支持上传/粘贴 HTML 文件内容，提供预览，并导出 PNG 与 PDF。

## 需求拆分
- 输入：`.html` 文件（包含内联/外链资源），或直接粘贴 HTML 字符串。
- 预览：在右侧预览区域渲染 HTML（隔离、消毒、适配 A4 纸宽）。
- 导出：将当前预览内容导出为长图（PNG）与分页 PDF。
- 兼容：Electron 桌面模式与 Web 服务模式共用一套前端逻辑。
- 安全：阻断脚本、内联事件，限制外部资源加载。

## 技术路线
1) **HTML 渲染容器**
   - 使用隐藏/隔离的 `<iframe>` 或 `Shadow DOM` 容器渲染 HTML，避免污染主页面样式。
   - 对外链样式/字体：可选策略 A) 禁止；B) 允许同源/HTTPS，提供“加载外链资源”开关。
   - 图片：优先内联 base64；外链图片允许 `https`，失败时显示占位符。

2) **安全处理**
   - 客户端使用 DOMPurify（或同类库）对白名单标签/属性净化；移除 `<script>`、事件属性、内联 JS。
   - 仅保留常用块级/行内标签、表格、列表、图片、链接、行内样式（可选）。

3) **版式与尺寸**
   - 预览区域宽度锁定为 A4 可用宽（~794px @96dpi），背景白色，边距可配置（或固定 2.54cm）。
   - 对 HTML 的 `body` 添加默认排版样式（字体、行距、表格边框）以保持一致输出。

4) **导出 PNG**
   - 复用 `html2canvas`，目标元素为 HTML 预览容器。
   - 选项：`scale: 2`（提高清晰度）、`useCORS: true`、`backgroundColor: '#fff'`。
   - 下载长图 `export_html_<timestamp>.png`。

5) **导出 PDF**
   - 生成全高 canvas -> 转换为图片，按 A4 尺寸分页加入 `jspdf`。
   - 分页算法：按 `pageHeight` 分段，循环 `addPage` + `addImage` 直到耗尽高度。
   - 文件名 `export_html_<timestamp>.pdf`。
   - Electron 模式无需调用 `printToPDF`，直接使用前端同样逻辑，保证多页完整。

6) **与现有逻辑的集成**
   - UI：在“导入”处新增 “导入 HTML” 按钮/文件选择；在编辑区域增加 “HTML/Markdown” 切换模式。
   - 状态：新增 `currentHtmlContent` 状态；预览函数根据模式决定渲染 Markdown->Docx 预览或 HTML 预览。
   - 导出按钮：当模式为 HTML 时调用 `exportHtmlAsPng/Pdf`；Markdown 模式保持原有 Docx 导出。
   - 不修改后端：HTML 预览与导出完全前端完成，无需服务器/Node 参与。

7) **资源处理策略**
   - CSS：允许 `<style>` 内联；外链 `<link rel="stylesheet">` 默认禁止，提供可选允许开关。
   - 图片：内联 `data:` 直接渲染；远程图片需要 `https`，失败时显示占位符并提示。
   - 字体：提供默认字体栈；外链字体需用户同意后加载。

8) **错误与提示**
   - 解析/净化失败、资源加载失败，需在状态栏显示具体原因。
   - 导出时如跨域图片导致 canvas 污染，提示用户关闭外链或启用 `useCORS` 且服务器允许 CORS。

9) **测试要点**
   - 含表格/列表/图片/链接的 HTML 正确渲染。
   - 跨页内容导出 PDF 完整分页，不截断。
   - 外链图片正常/失败两种场景。
   - 恶意脚本、事件处理被净化。
   - Electron 与浏览器模式一致性。

10) **后续可选增强**
   - 提供“打印留白/页眉页脚”配置。
   - 自动收集外链资源并内联，提升离线导出可靠性。
   - 允许将 HTML 转为 Docx（可复用现有 Markdown->Docx 流程：HTML -> Markdown via Turndown -> Docx）。
