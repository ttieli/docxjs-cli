# 前端模板样式可视化预览与 UI 布局优化方案

本文档描述了如何在 docxjs-cli 的 Web/App 界面中实现高质量的模板样式可视化预览功能，重点阐述优化的 UI 布局与交互设计。

## 1. 设计目标
- **所见即所得 (WYSIWYG)**：用户在调整字体、边距、颜色等参数时，能即时看到其在标准 A4 纸面上的渲染效果。
- **专业感与沉浸感**：采用类似专业排版软件（如 Word, Pages）的“左侧配置 - 右侧预览”双栏布局。
- **直观易用**：将复杂的样式配置项（几十个参数）进行语义化分组，降低用户认知负担。

## 2. 核心技术选型
*(保持不变)*
- **渲染引擎**：`docx-preview` (纯前端渲染 DOCX Blob)。
- **生成策略**：后端生成包含 "Lorem Ipsum" 标准样例文本的临时 Buffer -> 前端渲染。

## 3. UI 布局与交互设计 (重点优化)

我们采用 **“侧边栏配置 + 画布预览”** 的经典工作台布局。

### 3.1 整体布局结构 (Layout)

页面被划分为三个主要区域：

1.  **顶部导航栏 (Header - 60px)**
    -   左侧：Logo 与 项目名称。
    -   中间：当前模式指示（“模板编辑器”）。
    -   右侧：全局操作按钮（“导出当前配置”、“重置默认”、“帮助”）。

2.  **左侧配置面板 (Settings Panel - 320px ~ 400px 固定宽度, 可滚动)**
    -   承载所有的输入与控制项。
    -   背景色：白色或极浅灰 (`#F9FAFB`)，与深色预览背景形成对比。
    -   **分层设计**：
        -   **L1 模板选择**：最顶层下拉框，切换基准模板（如“红头公文”、“科技报告”）。
        -   **L2 样式微调**：使用 **手风琴 (Accordion)** 或 **Tabs** 折叠不同类别的样式参数。

3.  **右侧预览画布 (Preview Stage - 剩下所有空间)**
    -   背景色：深灰色 (`#E5E7EB` 或 `#525659`)，模拟桌面/应用背景。
    -   **居中画纸**：在中央渲染一个白色的 A4 纸张容器，带有真实的阴影效果 (`box-shadow`)，模拟真实纸张质感。
    -   **悬浮工具栏 (Floating Toolbar)**：位于画布右下角或底部中央，提供缩放控制（放大/缩小/适应屏幕）。

### 3.2 详细组件设计

#### A. 左侧配置面板 (Style Editor)

为了解决参数过多的问题，我们将配置项分为以下 5 个折叠面板：

1.  **全局设置 (Page & Global)**
    -   **页边距**：四个输入框组成的矩形布局，直观对应 上/下/左/右。
    -   **页面大小**：A4 / A3 / Letter 下拉选择。
    -   **参考文档**：文件上传控件，支持拖拽。上传后显示“已提取样式：xxx.docx”的状态标签。

2.  **标题样式 (Headings)**
    -   采用 **Tabs** 切换 `H1` / `H2` / `H3`。
    -   **字体选择**：带搜索功能的下拉框（如 Select2），支持中文字体预览。
    -   **字号与颜色**：步进器 (Stepper) 控制字号，圆形色块 (Color Picker) 控制颜色。
    -   **对齐方式**：图标按钮组 (左/中/右)。

3.  **正文排版 (Body Text)**
    -   **字体与字号**：同上。
    -   **行间距**：滑动条 (Slider) + 输入框，范围 1.0 - 3.0，实时反馈疏密。
    -   **首行缩进**：开关 (Toggle switch)。

4.  **公文特定 (Official Doc)**
    -   *仅在选择公文模板时显示*。
    -   **红头文字**：输入框。
    -   **发文号**：输入框。
    -   **分割线**：红色分割线粗细调整。

5.  **表格 (Tables)**
    -   **边框风格**：下拉选择（实线/虚线/无）。
    -   **表头背景**：颜色选择器。
    -   **紧凑模式**：开关。

#### B. 右侧预览区域 (Preview Area)

-   **加载状态**：在后端生成预览流时，画布上显示骨架屏 (Skeleton Screen) 或 旋转 Loading，提示“正在生成预览...”。
-   **渲染容器**：`div#preview-container`。
    -   样式：`width: 210mm; min-height: 297mm; background: white; margin: 40px auto; padding: 0; box-shadow: 0 10px 30px rgba(0,0,0,0.15);`
    -   利用 `docx-preview` 渲染到此容器内。
-   **样例文本策略**：
    -   不使用用户上传的长文档（避免渲染慢）。
    -   后端始终返回一份 **"UI测试专用文档"**，内容包含：
        -   一行一级标题（测试红头/大标题）。
        -   两行二级标题。
        -   三段正文（测试行距、首行缩进）。
        -   一个 3x3 表格（测试边框、表头）。
        -   一个无序列表。

### 3.3 交互逻辑 (Interaction Flow)

1.  **防抖更新 (Debounce)**：
    -   用户拖动“行间距”滑块时，不要每 1ms 请求一次。
    -   设置 500ms ~ 800ms 的防抖延迟，用户停止操作后自动触发预览刷新。
    -   或者提供 **"自动刷新"** 的勾选框（默认开启）。

2.  **手动刷新模式**：
    -   如果文档生成很慢（>2s），自动切换为“手动模式”：配置面板底部显示高亮的 **"应用更改"** 按钮，用户修改多项后点击一次刷新。

3.  **对比模式 (Compare Mode)**：
    -   (高级功能) 提供“原版 vs 新版”切换按钮，快速查看样式修改前后的差异。

## 4. 前端实现代码结构建议

```html
<!-- Layout Grid -->
<div class="app-container">
  <!-- Sidebar -->
  <aside class="settings-panel">
    <div class="panel-header">
      <h2>样式配置</h2>
    </div>
    
    <div class="scroll-content">
      <!-- Section: Global -->
      <details open>
        <summary>页面布局</summary>
        <div class="control-grid">
           <!-- Margins Inputs -->
        </div>
      </details>
      
      <!-- Section: Typography -->
      <details>
        <summary>标题与正文</summary>
        <!-- H1/H2/H3 Tabs -->
      </details>
      <!-- ... more sections ... -->
    </div>

    <div class="panel-footer">
      <button id="applyBtn" class="btn-primary">刷新预览</button>
    </div>
  </aside>

  <!-- Main Preview -->
  <main class="preview-stage">
    <div class="toolbar-floating">
      <button onclick="zoomIn()">+</button>
      <span id="zoomLabel">100%</span>
      <button onclick="zoomOut()">-</button>
    </div>
    
    <div class="canvas-scroll">
      <div id="paper-container">
         <!-- docx-preview renders here -->
      </div>
    </div>
  </main>
</div>
```

## 5. CSS 样式参考 (Variables)

```css
:root {
  --sidebar-width: 360px;
  --header-height: 60px;
  --bg-app: #F3F4F6;
  --bg-panel: #FFFFFF;
  --primary-color: #2563EB;
  --border-color: #E5E7EB;
  --paper-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
}

.app-container {
  display: flex;
  height: 100vh;
  background-color: var(--bg-app);
}

.settings-panel {
  width: var(--sidebar-width);
  background: var(--bg-panel);
  border-right: 1px solid var(--border-color);
  display: flex;
  flex-direction: column;
}

.preview-stage {
  flex: 1;
  overflow: auto; /* Allow scrolling the A4 paper */
  display: flex;
  justify-content: center;
  align-items: flex-start;
  padding: 40px;
}

#paper-container {
  width: 210mm; /* A4 Width */
  min-height: 297mm;
  background: white;
  box-shadow: var(--paper-shadow);
  transform-origin: top center;
  transition: transform 0.2s ease;
}
```

## 6. 总结
通过引入“侧边栏 + 画布”布局和 `docx-preview` 技术，我们将原本枯燥的 JSON 配置转化为直观的视觉反馈。这不仅提升了用户体验，也使得 docxjs-cli 从一个单纯的命令行工具进化为具备生产力的文档排版辅助工具。