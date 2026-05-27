# DOCX 表格列宽 QuickLook 挤压修复方案

> 创建时间：2026-05-27 21:25
> 管线：superpower-chain Pipeline A
> 关联问题：`../10_分析/20260527_DOCX表格列宽导致QuickLook预览挤压.md`
> 项目：FilesProcess_docxjs (`docx@9.5.1`)

**Goal:** 修复 `lib/core.js` 三处 `new Table()` 调用，让生成的 DOCX 表格在 Apple QuickLook / Pages / Mail / Spotlight / iCloud 等 OOXML 严格渲染器中不再塌缩成左侧竖条。

**Architecture:** 新增 3 个辅助函数（margin 解析、contentWidth 计算、列宽分配）→ 在三处 `new Table()` 上游计算 `columnWidths[]` 数组 → 改用 `WidthType.DXA` + 显式 `columnWidths` + `TableLayoutType.FIXED` + TableCell 加 `width` 加固。所有改动集中在 `lib/core.js` 单文件，无 schema/接口变更，向后完全兼容。

**Tech Stack:** Node.js + `docx@9.5.1` + Markdown-it。无测试框架（验证靠独立的 XML 断言脚本 + `qlmanage` 命令）。

---

## ⚠️ 执行说明（A5 评审强制要求）

**所有 Edit 操作的 `old_string` 必须用 Read 工具从源文件直接复制原始字符串，不要从本方案的代码块中复制。** 原因：

1. **缩进不一致**：方案中代码块按 0/4 空格展示便于阅读，但 `lib/core.js` 实际嵌套深处（如 L495-L537、L1292-L1311）缩进高达 16-24 空格。Edit 工具按字面匹配，方案中的代码块作 old_string 会全部失败。
2. **trailing whitespace**：`lib/core.js` 多处行尾有空格（如 L1290、L1313、L1331），看似空行实则非空。Read 时会看到 16 个空格的"伪空行"。
3. **方案中的 `新代码` 块也按方案缩进展示**，应用 Edit 时 `new_string` 也需手动加上原缩进对齐周围代码。

**执行节奏**：每个 Step 先 Read 实际行号 → 复制 old_string → 按缩进对齐写 new_string → Edit → 用 `node -c` 验证 → 用 `git diff` 视觉确认。

**Task 命名说明**：Task 2 已按修订后的"正向编号"展开（先定义 helper 计算 → 再用 cellConfig），不存在执行顺序倒置问题。

---

## 改动总览

| 文件 | 类型 | 说明 |
|------|------|------|
| `lib/core.js` | 修改 | imports 加 `TableLayoutType`；新增 3 个辅助函数；3 处 `new Table()` + 上游 TableCell 修改 |
| `scripts/verify-table-xml.js` | 新增 | XML 级回归脚本：解压 DOCX → grep 关键属性 → 断言 |
| `docs/test-output.docx` | 重新生成 | 用现有 `docs/test-all-formats.md` 跑一遍验证 |

**关键技术约束**（A2/A3 已验证）：
- docx@9.5.1 的 `WidthType.PERCENTAGE` 实现 bug：number 类型 size 会被字面拼 `%` → 输出非法 `w:w="100%"`
- 不传 `columnWidths` → docx 默认 `Array(N).fill(100)` twips ≈ N×1.76mm（核心根因）
- `TableLayoutType.FIXED = "fixed"`（已在 docx 库 d.ts L2726-L2729 验证）
- 默认页面：A4 = 11906×16838 twips；margin 从 `currentStyle.margin` 读取，可能为 cm/in/mm 字符串
- 单位换算：1cm=567 twips（精确：1cm=566.93），1in=1440 twips，1mm=56.7 twips

---

## File Structure

`lib/core.js` 的新结构（增量部分）：

```
lib/core.js
├── imports (L33-37)
│   └── + TableLayoutType  ← 新增
├── helpers (新增节，放在文件靠前位置，约 L88-150)
│   ├── parseMarginToTwips(marginStr) → twips
│   ├── calcContentWidth(pageWidth, margin) → twips
│   └── computeColumnWidths(contentWidth, colCount) → number[]
├── htmlTableToDocx (L456-L562)
│   ├── 改 TableCell 构造，加 width
│   └── 改 new Table，传 columnWidths + layout + DXA
├── code block table (L1170-L1196)
│   ├── 改 TableCell，加 width
│   └── 改 new Table，传 columnWidths + layout + DXA
└── markdown table (L1284-L1349)
    ├── 改 TableCell，加 width
    └── 改 new Table，传 columnWidths + layout + DXA
```

---

## Task 1: 引入 TableLayoutType + 新增 3 个辅助函数

**Files:**
- Modify: `lib/core.js:33-37`（imports）
- Modify: `lib/core.js`（在 `// --- Image Handling & Caching ---` 之前新增 helpers 节，约 L38 之前）

**目标**：为后续 3 处修复提供共用基础设施。无任何对外行为变化。

- [ ] **Step 1.1: 修改 imports 加 TableLayoutType**

旧代码（L33-37）：
```js
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    BorderStyle, HeadingLevel, AlignmentType, WidthType, VerticalAlign, ShadingType,
    ExternalHyperlink, UnderlineType, ImageRun, LineRuleType, FootnoteReferenceRun
} = require('docx');
```

新代码：
```js
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    BorderStyle, HeadingLevel, AlignmentType, WidthType, VerticalAlign, ShadingType,
    ExternalHyperlink, UnderlineType, ImageRun, LineRuleType, FootnoteReferenceRun,
    TableLayoutType
} = require('docx');
```

- [ ] **Step 1.2: 在 imports 块和 `// --- Image Handling & Caching ---` 之间插入辅助函数节**

在 L38 之前（即 imports 解构结束、Image Handling 注释之前）插入以下代码：

```js
// --- Table Width Helpers (fixes Apple QuickLook table compression bug) ---
// See docs/10_分析/20260527_DOCX表格列宽导致QuickLook预览挤压.md
const PAGE_WIDTH_A4_TWIPS = 11906;   // docx 库默认页面（lib/core.js Document 未传 page.size）
const DEFAULT_MARGIN_TWIPS = 1134;   // 2cm fallback when margin parsing fails

/**
 * Parse margin string (e.g. "2.54cm", "1in", "20mm", "1440") to twips.
 * Falls back to DEFAULT_MARGIN_TWIPS for unknown formats.
 * @param {string|number} marginStr
 * @returns {number} twips (Math.floor applied)
 */
function parseMarginToTwips(marginStr) {
    if (typeof marginStr === 'number') return Math.floor(marginStr);
    if (typeof marginStr !== 'string') return DEFAULT_MARGIN_TWIPS;
    const m = marginStr.trim().match(/^([\d.]+)\s*(cm|mm|in|pt|twip|twips)?$/i);
    if (!m) return DEFAULT_MARGIN_TWIPS;
    const num = parseFloat(m[1]);
    const unit = (m[2] || 'twip').toLowerCase();
    let twips;
    switch (unit) {
        case 'cm':            twips = num * 567; break;        // 1cm = 567 twips
        case 'mm':            twips = num * 56.7; break;       // 1mm = 56.7 twips
        case 'in':            twips = num * 1440; break;       // 1in = 1440 twips
        case 'pt':            twips = num * 20; break;         // 1pt = 20 twips
        case 'twip':
        case 'twips': default: twips = num;
    }
    return Math.floor(twips);
}

/**
 * Compute table content width in twips, based on page width and margins.
 * @param {object} marginObj - { left, right, top, bottom } from currentStyle.margin
 * @param {number} pageWidthTwips - default A4 = 11906
 * @returns {number} contentWidth twips
 */
function calcContentWidth(marginObj, pageWidthTwips = PAGE_WIDTH_A4_TWIPS) {
    const leftTw = parseMarginToTwips(marginObj && marginObj.left);
    const rightTw = parseMarginToTwips(marginObj && marginObj.right);
    const content = pageWidthTwips - leftTw - rightTw;
    // Safety: if margins are unreasonably large, fallback to 60% of page width
    if (content < 1000) return Math.floor(pageWidthTwips * 0.6);
    return content;
}

/**
 * Split contentWidth evenly across N columns.
 * Last column absorbs floor() remainder to keep sum exactly = contentWidth.
 * @param {number} contentWidth - twips
 * @param {number} colCount - integer >= 1
 * @returns {number[]} length = colCount, all integers, sum = contentWidth
 */
function computeColumnWidths(contentWidth, colCount) {
    if (!colCount || colCount < 1) return [contentWidth];
    const base = Math.floor(contentWidth / colCount);
    const widths = new Array(colCount).fill(base);
    const remainder = contentWidth - base * colCount;
    widths[colCount - 1] += remainder;   // last column absorbs floor remainder
    return widths;
}
```

- [ ] **Step 1.3: 自检 — node 语法检查**

```bash
cd "/Users/tieli/Library/Mobile Documents/com~apple~CloudDocs/Project/FilesProcess_docxjs"
node -c lib/core.js
```

Expected: 无输出（语法正确）

- [ ] **Step 1.4: 自检 — 函数行为快速验证**（一行 REPL）

```bash
node -e "
const c = require('./lib/core.js');
// 由于 c 只 export generateDocx 和 closeSharedBrowser，无法直接测 helper
// 改为读源码片段确认函数已写入：
const src = require('fs').readFileSync('lib/core.js', 'utf8');
console.log('parseMarginToTwips:', src.includes('function parseMarginToTwips'));
console.log('calcContentWidth:', src.includes('function calcContentWidth'));
console.log('computeColumnWidths:', src.includes('function computeColumnWidths'));
console.log('TableLayoutType imported:', src.includes('TableLayoutType') && src.match(/} = require\('docx'\)/));
"
```

Expected:
```
parseMarginToTwips: true
calcContentWidth: true
computeColumnWidths: true
TableLayoutType imported: true
```

- [ ] **Step 1.5: Commit**

```bash
git add lib/core.js
git commit -m "$(cat <<'EOF'
fix(table): add width helpers + TableLayoutType import

Foundation for fixing Apple QuickLook table compression bug:
- parseMarginToTwips: cm/mm/in/pt/twip → twips
- calcContentWidth: page width - margins, with safety fallback
- computeColumnWidths: even split, last col absorbs remainder
- Import TableLayoutType for FIXED layout in subsequent commits

No behavior change yet; consumers landed in next 3 commits.

Refs: docs/10_分析/20260527_DOCX表格列宽导致QuickLook预览挤压.md

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: 修复 htmlTableToDocx (含 colspan/rowspan)

**Files:**
- Modify: `lib/core.js:456-562`（`htmlTableToDocx` 函数）

**目标**：HTML 表格构造点 —— 三处中最复杂，需处理 colspan（cell.width = 跨列宽度合计）。

**colspan 处理逻辑**：
- 标准列宽数组 `columnWidths` 长度 = `tableData.colCount`（首行展开列数）
- 单元格 `cell.colSpan > 1` 时，其 width = `columnWidths[cell.colIndex] + columnWidths[cell.colIndex+1] + ...`
- rowSpan 不影响列宽（垂直跨行不改变列结构）

- [ ] **Step 2.1: 在 `// Build docx rows` 之前（L488 之前）插入 columnWidthsHtml 计算 + colCount=0 守护**

**先 Read `lib/core.js:486-490`** 找到精确插入点。当前 L488 是 `    // Build docx rows`（前导 4 空格）。

在 L488 之前插入以下代码（注意保持与周围 4 空格缩进一致）：

```js
    // Safety: empty HTML table produces tableData.colCount === 0 — return null instead of building Table
    if (tableData.colCount === 0) return null;

    // Compute column widths (Apple QuickLook needs explicit tblGrid + tcW)
    const contentWidthHtml = calcContentWidth(currentStyle.margin);
    const columnWidthsHtml = computeColumnWidths(contentWidthHtml, tableData.colCount);

```

验证 `columnWidthsHtml` 已在 Step 2.2 使用范围内（函数作用域，覆盖到 L562 return）。

- [ ] **Step 2.2: 修改 cellConfig 加 width 字段（在 L521-L527 区域）**

**先 Read `lib/core.js:521-528`** 复制实际代码作 old_string（嵌套 ~16 空格缩进）。

实际代码片段（示意，实际有外层缩进）：
```js
                const cellConfig = {
                    children: [new Paragraph({
                        children: cellRuns,
                        alignment: align,
                    })],
                    verticalAlign: VerticalAlign.CENTER,
                };
```

修改为（新增最后一项 `width`，注意 cell.colSpan 默认 1，由 parseHtmlTable L391/L418 保证 ≥ 1）：
```js
                const cellConfig = {
                    children: [new Paragraph({
                        children: cellRuns,
                        alignment: align,
                    })],
                    verticalAlign: VerticalAlign.CENTER,
                    // Explicit cell width for Apple QuickLook (sums multi-col when colSpan > 1)
                    width: {
                        size: columnWidthsHtml.slice(colIndex, colIndex + cell.colSpan).reduce((a, b) => a + b, 0),
                        type: WidthType.DXA,
                    },
                };
```

**⚠️ 保持不动**：L529-L535 现有的 `cellConfig.rowSpan` 和 `cellConfig.columnSpan` 赋值必须**完整保留**（docx 库会自动从 columnSpan 生成 `w:gridSpan`、从 rowSpan 自动插入 continue cell）。新加的 `width` 与现有 `columnSpan` 字段互补，缺一不可：
- `columnSpan` 告诉 docx 库这个 cell 跨 N 列（输出 `w:gridSpan`）
- `width` 告诉渲染器这个 cell 实际宽度（输出 `w:tcW = N 列宽合计`）

- [ ] **Step 2.3: 修改 new Table 调用（L558-L562）**

旧代码（L558-L562）：
```js
return new Table({
    rows: docxRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders
});
```

新代码：
```js
return new Table({
    rows: docxRows,
    width: { size: contentWidthHtml, type: WidthType.DXA },
    columnWidths: columnWidthsHtml,
    layout: TableLayoutType.FIXED,
    borders: tableBorders
});
```

- [ ] **Step 2.4: 自检 — 语法 + 局部 grep 确认改动**

```bash
cd "/Users/tieli/Library/Mobile Documents/com~apple~CloudDocs/Project/FilesProcess_docxjs"
node -c lib/core.js
sed -n '485,562p' lib/core.js | grep -E 'columnWidthsHtml|contentWidthHtml|TableLayoutType\.FIXED|WidthType\.DXA'
```

Expected: 至少 4 行匹配（计算 + cell.width + Table.width + columnWidths + layout）

- [ ] **Step 2.5: 自检 — 跑一遍 HTML 表格样本生成**

```bash
cat > /tmp/test_html_table.md <<'EOF'
# HTML 表格测试

<table>
<tr><th>姓名</th><th>年龄</th><th>城市</th></tr>
<tr><td>张三</td><td>30</td><td>北京</td></tr>
<tr><td>李四</td><td>25</td><td>上海</td></tr>
</table>
EOF
node bin/cli.js /tmp/test_html_table.md -o /tmp/test_html_table.docx --quiet
unzip -p /tmp/test_html_table.docx word/document.xml | grep -oE 'tblW[^/]*|gridCol[^/]*|tblLayout[^/]*' | head -10
```

Expected: `tblW w:type="dxa" w:w="9026"`（或接近，取决于默认模板 margin）；`gridCol w:w` 不为 100；`tblLayout w:type="fixed"` 必现

- [ ] **Step 2.6: Commit**

```bash
git add lib/core.js
git commit -m "$(cat <<'EOF'
fix(table): use DXA + columnWidths + FIXED layout for HTML tables

htmlTableToDocx (L456-562) — first of three table writers fixed.

Changes:
- Compute columnWidths via computeColumnWidths(contentWidth, colCount)
- TableCell.width = sum of spanned columns (handles colspan > 1)
- Table: width=DXA contentWidth, columnWidths array, layout=FIXED
- Removes invalid w:w="100%" output, removes default 100-twip gridCol

Apple QuickLook / Pages / Mail now render HTML tables correctly.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: 修复代码块表格 (单列单行)

**Files:**
- Modify: `lib/core.js:1170-1196`

**目标**：代码块包装表格 —— 单列单行，最简单的场景。

- [ ] **Step 3.1: 在 codeTable 构造前（L1176 之前）计算列宽**

旧代码（L1170-L1175 + L1176 起）：
```js
// Create table container with border and padding
const codeBorder = {
    style: BorderStyle.SINGLE,
    size: 6,
    color: "DDDDDD"
};
const codeTable = new Table({
    ...
```

新代码（在 `const codeTable = new Table({` 之前插入）：
```js
// Create table container with border and padding
const codeBorder = {
    style: BorderStyle.SINGLE,
    size: 6,
    color: "DDDDDD"
};
// Compute column widths (Apple QuickLook compatibility — single full-width column)
const contentWidthCode = calcContentWidth(currentStyle.margin);
const columnWidthsCode = [contentWidthCode];
const codeTable = new Table({
    ...
```

- [ ] **Step 3.2: 修改 TableCell 加 width 字段（L1178-L1187）**

旧代码（L1178-L1187）：
```js
children: [new TableCell({
    children: codeParagraphs,  // Multiple Paragraphs, one per line
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "F5F5F5" },
    margins: {
        top: 80,
        bottom: 80,
        left: 150,
        right: 150
    }
})]
```

新代码：
```js
children: [new TableCell({
    children: codeParagraphs,  // Multiple Paragraphs, one per line
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "F5F5F5" },
    margins: {
        top: 80,
        bottom: 80,
        left: 150,
        right: 150
    },
    width: { size: contentWidthCode, type: WidthType.DXA },
})]
```

- [ ] **Step 3.3: 修改 codeTable 构造的 width / layout / columnWidths（L1189-L1195）**

旧代码（L1189-L1195）：
```js
borders: {
    top: codeBorder,
    bottom: codeBorder,
    left: codeBorder,
    right: codeBorder
},
width: { size: 100, type: WidthType.PERCENTAGE }
```

新代码：
```js
borders: {
    top: codeBorder,
    bottom: codeBorder,
    left: codeBorder,
    right: codeBorder
},
width: { size: contentWidthCode, type: WidthType.DXA },
columnWidths: columnWidthsCode,
layout: TableLayoutType.FIXED
```

- [ ] **Step 3.4: 自检 — 跑代码块样本**

```bash
cat > /tmp/test_code_table.md <<'EOF'
# 代码块测试

\`\`\`python
def hello():
    print("Hello, world!")
\`\`\`

正常段落。
EOF
node bin/cli.js /tmp/test_code_table.md -o /tmp/test_code_table.docx --quiet
unzip -p /tmp/test_code_table.docx word/document.xml | grep -oE 'tblW[^/]*|gridCol[^/]*|tblLayout[^/]*|tcW[^/]*' | head -10
```

Expected: `tblW w:type="dxa"`、`gridCol` 接近 contentWidth、`tblLayout w:type="fixed"`、`tcW w:type="dxa"`（cell width 也加了）

- [ ] **Step 3.5: Commit**

```bash
git add lib/core.js
git commit -m "$(cat <<'EOF'
fix(table): use DXA + FIXED layout for code block wrapper table

Code block visual wrapper (L1170-1196) — second of three table writers fixed.

Changes:
- columnWidths = [contentWidth] (single full-width column)
- TableCell.width = contentWidth (DXA)
- Table: width=DXA, columnWidths array, layout=FIXED

Code blocks now render with full page width in Apple QuickLook.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

## Task 4: 修复 markdown table (标准 `| ... |` 语法)

**Files:**
- Modify: `lib/core.js:1284-1349`（`table_close` 分支）

**目标**：第三处也是最常用的表格构造点。无 colspan/rowspan，但需注意：cell 数 = 首行 cells 数（标准 markdown table 的行内列数固定）。

- [ ] **Step 4.1: 在 `// Process table rows asynchronously` 之前（L1291 之前）插入列宽计算 + 空表守护**

**⚠️ 注意 L1290 是 trailing whitespace 行（16 个空格的"伪空行"）**，不要把它纳入 old_string。使用 `Read lib/core.js:1291-1292` 锁定精确插入点，把新代码插在 L1291 那一行 `                // Process table rows asynchronously` 之前。

实际插入策略：用 Edit 工具的 old_string 匹配 L1291-L1292（带 16 空格缩进的"// Process ..." 行 + 下一行的 const docxRows = ...）：

```js
                // Process table rows asynchronously
                const docxRows = await Promise.all(tableBuffer.rows.map(async rowObj => {
```

替换为（在 // Process 之前插入空表守护 + 列宽计算）：
```js
                // Safety: empty markdown table — skip
                const mdColCount = (tableBuffer.rows[0] && tableBuffer.rows[0].content)
                    ? tableBuffer.rows[0].content.length
                    : 0;
                if (mdColCount === 0) {
                    tableBuffer = null;
                    break;
                }

                // Compute column widths (Apple QuickLook compatibility)
                const contentWidthMd = calcContentWidth(currentStyle.margin);
                const columnWidthsMd = computeColumnWidths(contentWidthMd, mdColCount);

                // Process table rows asynchronously
                const docxRows = await Promise.all(tableBuffer.rows.map(async rowObj => {
```

注意所有插入行的**缩进必须是 16 空格**（与 L1291 一致）。`break` 跳出 `for (let i = 0; ...)` 主循环——读 lib/core.js 周围代码确认作用域：当前 `else if (token.type === 'table_close')` 在主 token loop 中（约 L1100+），`break` 会跳出该 loop。如果不希望中断后续 token 处理，可改为 `tableBuffer = null; continue;` 但当前结构里 break/continue 都能让 `tableBuffer.rows.length > 0` 后续判断不进入，行为安全。**建议用 `tableBuffer = null;` 后让代码自然走完 else if 块（即不加 break/continue），更稳**：

```js
                // Safety: empty markdown table — skip table construction
                const mdColCount = (tableBuffer.rows[0] && tableBuffer.rows[0].content)
                    ? tableBuffer.rows[0].content.length
                    : 0;
                if (mdColCount === 0) {
                    tableBuffer = null;
                } else {
                    // Compute column widths (Apple QuickLook compatibility)
                    const contentWidthMd = calcContentWidth(currentStyle.margin);
                    const columnWidthsMd = computeColumnWidths(contentWidthMd, mdColCount);

                    // Process table rows asynchronously
                    const docxRows = await Promise.all(tableBuffer.rows.map(async rowObj => {
```

注意此 else 分支需要在 Step 4.3 的 docChildren.push 之后闭合 `}`。详见 Step 4.3。

**简化版（推荐采用）**：直接在守护条件不成立时让后续代码走完正常路径——把 `if (mdColCount === 0)` 单独提到最前面，命中就 `tableBuffer = null; ` 然后让外层 `else if (token.type === 'table_close')` 块的剩余代码（docxRows、docChildren.push）变成"NOOP，因为 tableBuffer 被置 null 但已经进 if 块"——但其实剩余代码不检查 tableBuffer，所以会执行空 rows。**正确做法**：用早返回模式：

```js
                // Safety: empty markdown table — skip
                const mdColCount = (tableBuffer.rows[0] && tableBuffer.rows[0].content)
                    ? tableBuffer.rows[0].content.length
                    : 0;
                if (mdColCount === 0) {
                    tableBuffer = null;
                    continue;  // skip this token, continue outer token loop
                }

                // Compute column widths (Apple QuickLook compatibility)
                const contentWidthMd = calcContentWidth(currentStyle.margin);
                const columnWidthsMd = computeColumnWidths(contentWidthMd, mdColCount);

                // Process table rows asynchronously
                const docxRows = await Promise.all(tableBuffer.rows.map(async rowObj => {
```

实施者请 Read `lib/core.js:1098-1110` 确认外层 for loop 的变量是 `i` 还是其他名，确认 `continue` 跳到的是该 for loop。如果外层是 `for-of` 或其他形式，按实际调整。

- [ ] **Step 4.2: 修改 TableCell 构造加 width（L1303-L1309）**

旧代码（L1303-L1309）：
```js
return new TableCell({
    children: [new Paragraph({
        children: cellRuns,
        alignment: align,
    })],
    verticalAlign: VerticalAlign.CENTER,
});
```

新代码：标准 md table 无 colspan，cell 的列索引就是其在 cells 数组中的索引。需要把 `cellText` 的 map 改写为带 index：

旧代码上下文（L1292-L1311）：
```js
const docxRows = await Promise.all(tableBuffer.rows.map(async rowObj => {
    const cells = await Promise.all(rowObj.content.map(async cellText => {
        let isBold = rowObj.isHeader ? tblConfig.headerBold : false;
        let color = rowObj.isHeader ? tblConfig.headerColor : (currentStyle.colorMain || "000000");
        let align = AlignmentType.LEFT;
        if (tblConfig.cellAlign === 'center') align = AlignmentType.CENTER;
        if (tblConfig.cellAlign === 'right') align = AlignmentType.RIGHT;
        
        const cellTokens = md.parseInline(cellText, {})[0];
        const cellRuns = await processInline(md, cellTokens, currentStyle, baseDir, color, isBold);
        
        return new TableCell({
            children: [new Paragraph({
                children: cellRuns,
                alignment: align,
            })],
            verticalAlign: VerticalAlign.CENTER,
        });
    }));
    return new TableRow({ children: cells });
}));
```

新代码（把 inner map 加 index 参数 + 加 width）：
```js
const docxRows = await Promise.all(tableBuffer.rows.map(async rowObj => {
    const cells = await Promise.all(rowObj.content.map(async (cellText, cellIdx) => {
        let isBold = rowObj.isHeader ? tblConfig.headerBold : false;
        let color = rowObj.isHeader ? tblConfig.headerColor : (currentStyle.colorMain || "000000");
        let align = AlignmentType.LEFT;
        if (tblConfig.cellAlign === 'center') align = AlignmentType.CENTER;
        if (tblConfig.cellAlign === 'right') align = AlignmentType.RIGHT;

        const cellTokens = md.parseInline(cellText, {})[0];
        const cellRuns = await processInline(md, cellTokens, currentStyle, baseDir, color, isBold);

        // Use columnWidthsMd[cellIdx]; fallback to last col for over-flow safety
        const cellWidth = columnWidthsMd[cellIdx] !== undefined
            ? columnWidthsMd[cellIdx]
            : columnWidthsMd[columnWidthsMd.length - 1];

        return new TableCell({
            children: [new Paragraph({
                children: cellRuns,
                alignment: align,
            })],
            verticalAlign: VerticalAlign.CENTER,
            width: { size: cellWidth, type: WidthType.DXA },
        });
    }));
    return new TableRow({ children: cells });
}));
```

- [ ] **Step 4.3: 修改 docChildren.push(new Table(...)) (L1332-L1336)**

旧代码（L1332-L1336）：
```js
docChildren.push(new Table({
    rows: docxRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders
}));
```

新代码：
```js
docChildren.push(new Table({
    rows: docxRows,
    width: { size: contentWidthMd, type: WidthType.DXA },
    columnWidths: columnWidthsMd,
    layout: TableLayoutType.FIXED,
    borders: tableBorders
}));
```

- [ ] **Step 4.4: 自检 — 跑 markdown table 样本**

```bash
cat > /tmp/test_md_table.md <<'EOF'
# Markdown 表格测试

| 姓名 | 年龄 | 城市 | 备注 |
|------|------|------|------|
| 张三 | 30 | 北京 | 工程师 |
| 李四 | 25 | 上海 | 设计师 |
| 王五 | 28 | 深圳 | 产品 |
EOF
node bin/cli.js /tmp/test_md_table.md -o /tmp/test_md_table.docx --quiet
unzip -p /tmp/test_md_table.docx word/document.xml | grep -oE 'tblW[^/]*|gridCol[^/]*|tblLayout[^/]*|tcW[^/]*' | head -20
```

Expected:
- `tblW w:type="dxa" w:w="9026"`（或接近，看默认 margin）
- 4 个 `gridCol`，每列约 2256 twips（9026/4），最后一列吸收余数
- `tblLayout w:type="fixed"`
- `tcW w:type="dxa"`（每个单元格都有）
- **绝对无 `w:w="100%"`**

- [ ] **Step 4.5: Commit**

```bash
git add lib/core.js
git commit -m "$(cat <<'EOF'
fix(table): use DXA + columnWidths + FIXED layout for markdown tables

Markdown table writer (L1284-1349) — third and final table writer fixed.

Changes:
- Compute mdColCount from first row's cell count
- columnWidthsMd via computeColumnWidths(contentWidth, colCount)
- TableCell.width = columnWidthsMd[cellIdx] (DXA, fallback to last col)
- Table: width=DXA, columnWidths array, layout=FIXED
- Inner map gains cellIdx parameter for column-aware sizing

All three table writers (HTML, code block, markdown) now produce
OOXML-compliant table widths. Apple QuickLook / Pages / Mail render
tables correctly.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

## Task 5: 新增 XML 级回归脚本

**Files:**
- Create: `scripts/verify-table-xml.js`

**目标**：可独立运行的 XML 断言脚本，未来回归测试用。

- [ ] **Step 5.1: 创建 scripts/verify-table-xml.js**

```bash
mkdir -p "/Users/tieli/Library/Mobile Documents/com~apple~CloudDocs/Project/FilesProcess_docxjs/scripts"
```

文件内容（`scripts/verify-table-xml.js`）：

```js
#!/usr/bin/env node
// XML-level regression check for DOCX table widths.
// Usage: node scripts/verify-table-xml.js <file.docx>
// Exits 0 on PASS, non-zero on FAIL.
//
// Validates the fix for: docs/10_分析/20260527_DOCX表格列宽导致QuickLook预览挤压.md

const fs = require('fs');
const { execSync } = require('child_process');

const docx = process.argv[2];
if (!docx) {
    console.error('Usage: node scripts/verify-table-xml.js <file.docx>');
    process.exit(2);
}
if (!fs.existsSync(docx)) {
    console.error(`File not found: ${docx}`);
    process.exit(2);
}

// Extract document.xml
let xml;
try {
    xml = execSync(`unzip -p "${docx}" word/document.xml`, { encoding: 'utf8' });
} catch (err) {
    console.error(`Failed to unzip: ${err.message}`);
    process.exit(2);
}

const checks = [];
const failures = [];

// Check 1: no literal w:w="100%"
const has100Pct = /w:w="100%"/.test(xml);
checks.push({ name: 'No literal w:w="100%"', pass: !has100Pct });

// Check 2: every tblW uses dxa or pct=integer (no string percent)
// Use attribute-order-independent matching (docx lib could reorder in future versions)
const tblWMatches = [...xml.matchAll(/<w:tblW\b([^/>]*)/g)];
const tblWBad = tblWMatches.filter(m => {
    const attrs = m[1];
    const hasType = /w:type="([^"]+)"/.exec(attrs);
    const hasW = /w:w="([^"]+)"/.exec(attrs);
    if (!hasType) return true;  // missing type is bad
    const type = hasType[1];
    const wVal = hasW ? hasW[1] : '';
    if (type === 'dxa') return !/^\d+$/.test(wVal);          // dxa requires integer
    if (type === 'auto') return false;                        // auto OK
    if (type === 'pct') return !/^\d+$/.test(wVal);          // pct requires integer (no %)
    return true;
});
checks.push({
    name: `All tblW use legal type (dxa/auto/pct-integer): ${tblWMatches.length} tables`,
    pass: tblWBad.length === 0,
    detail: tblWBad.length > 0 ? `bad: ${tblWBad.map(b => b[0]).join(', ')}` : null,
});

// Check 3: no tblGrid has all gridCols = 100
const tableBlocks = xml.split('<w:tbl>').slice(1);
const allTblGridBad = tableBlocks.map((block, idx) => {
    const tblGridMatch = block.match(/<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/);
    if (!tblGridMatch) return null;
    const gridCols = [...tblGridMatch[1].matchAll(/w:w="(\d+)"/g)].map(m => parseInt(m[1], 10));
    const allHundred = gridCols.length > 0 && gridCols.every(w => w === 100);
    return allHundred ? { idx, gridCols } : null;
}).filter(Boolean);
checks.push({
    name: `No table has all gridCol = 100 twips (${tableBlocks.length} tables checked)`,
    pass: allTblGridBad.length === 0,
    detail: allTblGridBad.length > 0 ? `bad tables: ${JSON.stringify(allTblGridBad)}` : null,
});

// Check 4: every table has tblLayout w:type="fixed"
const tablesWithoutFixed = tableBlocks.filter(block => !/<w:tblLayout\s+w:type="fixed"/.test(block));
checks.push({
    name: `All tables have tblLayout=fixed (${tableBlocks.length} tables)`,
    pass: tablesWithoutFixed.length === 0,
    detail: tablesWithoutFixed.length > 0 ? `${tablesWithoutFixed.length} tables missing tblLayout=fixed` : null,
});

// Check 5: gridCol values are reasonable (each > 500 twips OR contentWidth > 5000)
const tinyGridCols = [];
tableBlocks.forEach((block, idx) => {
    const tblGridMatch = block.match(/<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/);
    if (!tblGridMatch) return;
    const gridCols = [...tblGridMatch[1].matchAll(/w:w="(\d+)"/g)].map(m => parseInt(m[1], 10));
    const sum = gridCols.reduce((a, b) => a + b, 0);
    if (sum < 5000) tinyGridCols.push({ idx, sum, gridCols });
});
checks.push({
    name: 'tblGrid total width sane (sum > 5000 twips ~= 8.8cm)',
    pass: tinyGridCols.length === 0,
    detail: tinyGridCols.length > 0 ? `tables with narrow total: ${JSON.stringify(tinyGridCols)}` : null,
});

// Report
console.log(`\n=== XML Table Width Check for ${docx} ===\n`);
let allPass = true;
checks.forEach(c => {
    const icon = c.pass ? '✅' : '❌';
    console.log(`${icon} ${c.name}`);
    if (!c.pass && c.detail) console.log(`   ${c.detail}`);
    if (!c.pass) allPass = false;
});

if (allPass) {
    console.log(`\n✅ ALL CHECKS PASSED\n`);
    process.exit(0);
} else {
    console.log(`\n❌ SOME CHECKS FAILED\n`);
    process.exit(1);
}
```

- [ ] **Step 5.2: 自检 — chmod + 跑一遍 Task 4 生成的文件**

```bash
cd "/Users/tieli/Library/Mobile Documents/com~apple~CloudDocs/Project/FilesProcess_docxjs"
chmod +x scripts/verify-table-xml.js
node scripts/verify-table-xml.js /tmp/test_md_table.docx
```

Expected: 5 项全 ✅，退出码 0

- [ ] **Step 5.3: 反向验证（脚本能识别 bad 文件）**

```bash
node scripts/verify-table-xml.js "/Users/tieli/Library/Mobile Documents/com~apple~CloudDocs/HandOn/Outputs/conv_20260527_202124_0nvl/个人基本信息表_铁力_20260527.docx"
```

Expected: 多项 ❌（旧文件包含 bug），退出码 1。这证明脚本能识别问题。

- [ ] **Step 5.4: Commit**

```bash
git add scripts/verify-table-xml.js
git commit -m "$(cat <<'EOF'
test(table): add XML-level regression script for table widths

scripts/verify-table-xml.js validates DOCX table OOXML compliance:
- No literal w:w="100%" (docx@9.5.1 PERCENTAGE bug)
- tblW uses dxa/auto/pct-integer (no string percent)
- No table has all gridCol = 100 twips (docx default bug)
- All tables have tblLayout=fixed (QuickLook autofit avoidance)
- tblGrid total width >= 5000 twips (sanity check)

Usage: node scripts/verify-table-xml.js <file.docx>
Exits 0 on PASS, 1 on FAIL, 2 on error.

Verified against current bug sample (fails) and fixed test outputs (passes).

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

## Task 6: 全量回归验证（多模板 + 边界用例 + QuickLook 视觉）

**Files:** 无修改，仅运行验证

**目标**：用多种模板和边界用例确认修复彻底，QuickLook 视觉确认。

- [ ] **Step 6.1: 用现有 docs/test-all-formats.md 跑默认模板**

```bash
cd "/Users/tieli/Library/Mobile Documents/com~apple~CloudDocs/Project/FilesProcess_docxjs"
node bin/cli.js docs/test-all-formats.md -o /tmp/test_all_default.docx --quiet
node scripts/verify-table-xml.js /tmp/test_all_default.docx
```

Expected: 全 ✅

- [ ] **Step 6.2: 用多个模板跑同一样本**

```bash
for tpl in academic_paper_en tech_report_blue; do
    echo "=== Template: $tpl ==="
    node bin/cli.js docs/test-all-formats.md -o /tmp/test_all_${tpl}.docx --template $tpl --quiet
    node scripts/verify-table-xml.js /tmp/test_all_${tpl}.docx
done
```

Expected: 两个模板都全 ✅

- [ ] **Step 6.3: 边界用例 — 多列表格（10 列）**

```bash
cat > /tmp/test_wide.md <<'EOF'
# 10 列表格

| C1 | C2 | C3 | C4 | C5 | C6 | C7 | C8 | C9 | C10 |
|----|----|----|----|----|----|----|----|----|-----|
| a | b | c | d | e | f | g | h | i | j |
EOF
node bin/cli.js /tmp/test_wide.md -o /tmp/test_wide.docx --quiet
node scripts/verify-table-xml.js /tmp/test_wide.docx
unzip -p /tmp/test_wide.docx word/document.xml | grep -oE 'gridCol[^/]*' | head -12
```

Expected: 全 ✅；10 个 gridCol，每列约 900 twips；最后一列吸收余数

- [ ] **Step 6.4a: 边界用例 — HTML 含 colspan**

```bash
cat > /tmp/test_colspan.md <<'EOF'
# Colspan 测试

<table>
<tr><th colspan="2">合并 2 列</th><th>单列</th></tr>
<tr><td>a</td><td>b</td><td>c</td></tr>
</table>
EOF
node bin/cli.js /tmp/test_colspan.md -o /tmp/test_colspan.docx --quiet
node scripts/verify-table-xml.js /tmp/test_colspan.docx
unzip -p /tmp/test_colspan.docx word/document.xml | grep -oE 'gridCol[^/]*|tcW[^/]*|gridSpan[^/]*'
```

Expected:
- 全 ✅
- 3 个 gridCol（约 3008/3009/3009 twips）
- 首行第 1 个 cell 的 tcW = 6017 ± 1（前 2 列合计），含 gridSpan
- 第 2 行三个 cell 各自 tcW ≈ 3008

- [ ] **Step 6.4b: 边界用例 — HTML 含 rowspan（验证 docx 库自动 CONTINUE cell）**

docx 库会为 rowSpan > 1 的 cell 在下一行自动插入无 width 的 CONTINUE cell（实证 `node_modules/docx/dist/index.cjs:15226-15235`）。验证 tblGrid 不受影响：

```bash
cat > /tmp/test_rowspan.md <<'EOF'
# Rowspan 测试

<table>
<tr><td rowspan="2">合并 2 行</td><td>A</td><td>B</td></tr>
<tr><td>C</td><td>D</td></tr>
<tr><td>E</td><td>F</td><td>G</td></tr>
</table>
EOF
node bin/cli.js /tmp/test_rowspan.md -o /tmp/test_rowspan.docx --quiet
node scripts/verify-table-xml.js /tmp/test_rowspan.docx
unzip -p /tmp/test_rowspan.docx word/document.xml | grep -oE 'gridCol[^/]*|vMerge[^/]*' | head -10
```

Expected: 全 ✅；3 个 gridCol；首列 `vMerge` 出现（restart + continue）

- [ ] **Step 6.4c: PDF 路径回归验证（bin/capture.js 走浏览器预览）**

```bash
# 项目用 bin/capture.js 生成 PDF（基于 html2canvas + jsPDF 截浏览器预览）
node bin/cli.js docs/test-all-formats.md -o /tmp/regression.docx --quiet
node bin/capture.js /tmp/regression.docx --pdf /tmp/regression.pdf 2>&1 | head -5 || echo "(若 capture.js 接口不同请按项目实际命令)"
ls -lh /tmp/regression.pdf 2>/dev/null && echo "PDF 生成成功"
```

Expected: PDF 文件生成（大小合理）；视觉打开 PDF 确认表格宽度正常（应与 docx-preview 浏览器渲染一致）

- [ ] **Step 6.5: QuickLook 视觉验证（macOS 必做）**

```bash
qlmanage -t -s 1024 -o /tmp/qltest /tmp/test_md_table.docx
ls -lh /tmp/qltest/
# 用 Preview / open 打开 PNG 视觉确认表格宽度占页面比例 > 60%
open /tmp/qltest/*.png
```

如果在非 macOS 环境跑（不可能但安全起见），跳过本步并说明。

- [ ] **Step 6.6: QuickLook 全屏预览**

```bash
qlmanage -p /tmp/test_md_table.docx
```

Expected: 浮窗弹出，表格正常宽度展示（按 Esc 关闭）

- [ ] **Step 6.7: 重新生成 docs/test-output.docx（项目内回归样本）**

旧的 `docs/test-output.docx` 是 bug 前生成的样本，应用修复后覆盖：

```bash
node bin/cli.js docs/test-all-formats.md -o docs/test-output.docx --quiet
node scripts/verify-table-xml.js docs/test-output.docx
```

Expected: 全 ✅

- [ ] **Step 6.8: 把原 bug 样本重新过一遍（用 HandOn pipeline 的源 markdown 重生成）**

如果可获得源 markdown（HandOn 流程的原始输入），重新生成对比；否则用同一份 markdown 模拟：

```bash
# 用 test-all-formats.md 当代理样本即可（已包含表格）
node scripts/verify-table-xml.js docs/test-output.docx
echo "请手动 QuickLook 预览 docs/test-output.docx 确认表格不再挤压"
qlmanage -p docs/test-output.docx
```

- [ ] **Step 6.9: Commit 回归样本**

```bash
git add docs/test-output.docx
git commit -m "$(cat <<'EOF'
test(table): regenerate docs/test-output.docx with table fix applied

Replaces bug-era sample with fixed output. Verified by:
- scripts/verify-table-xml.js (all 5 checks pass)
- qlmanage -p (visual confirmation: tables render at full width)

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

## 风险评估

| 风险 | 概率 | 影响 | 缓解 |
|------|------|------|------|
| `currentStyle.margin` 在某些模板缺失 | 低 | parseMarginToTwips 返回 fallback，可能列宽不准 | calcContentWidth 有 < 1000 安全 fallback；DEFAULT_MARGIN_TWIPS = 1134（2cm） |
| HTML 表格 colspan 计算错误 | 中 | 合并单元格宽度异常 | `.slice(colIndex, colIndex + cell.colSpan).reduce(+)` 是简洁正确的实现；Task 6.4 边界用例覆盖 |
| 极端窄表格（用户预期窄列） | 低 | 修复后变成等宽，可能不符合用户预期 | 这是修复，原行为是 bug；未来可加 `currentStyle.table.columnWidths` 让用户覆盖 |
| 浮点精度漂移 | 低 | sum(gridCol) 不等于 contentWidth | Math.floor + 最后一列吸收余数，sum 严格等于 contentWidth |
| 影响下游 docx-preview 渲染 | 低 | docx-preview 是 docxjs 用的同一库的预览器，已支持 DXA + FIXED | 项目 `public/vendor/docx-preview` 版本可读 DXA，无回归风险 |
| `bin/cli.js` 不支持 `--quiet` 或 `--template` | 低 | 验证脚本失败 | 已在 8d125a6 提交中加入这些参数（recent commit log 可见） |
| 已存在的代码块的 padding（margins: 80/150 twips）会缩小内容宽度 | 极低 | 视觉感觉代码块比段落窄 | 这是 cell margin 不是 column width，不影响 tblGrid；保持现状 |

## 不变量（修复前后保持）

- `currentStyle.table` 的其他配置（borderStyle、headerBold、spacingBefore 等）行为不变
- 模板 JSON（`templates/*.json`）schema 不变，无需用户改模板
- `Document` 构造的 `page.size`/`page.margin` 不变（默认 A4 + 模板 margin）
- 三处 Table 构造的 borders、shading、margins、children 不变
- 命令行接口、API 接口、Electron 接口不变

## 跳过的可选优化（follow-up）

以下不在本次方案范围，可作为后续 issue：

- **支持模板自定义 columnWidths**：在 `currentStyle.table` 加 `columnWidths: [w1, w2, ...]` 让用户精细控制（如比例 2:1:1）
- **支持横向纸张**：当前 contentWidth 按竖向 A4 算；如未来加横向需读 sectPr.pgSz 然后计算
- **上游修复后切回 pct=5000**：等 docx 库修复 PERCENTAGE bug，可在 wrapper 加 `useAutoLayout` 开关
- **重构成 wrapper 函数**：`createDocxTable(rows, opts)` 内置宽度计算，避免三处重复代码

---

## 评审记录

### Round 1

**评审时间**：2026-05-27 21:32
**评审模式**：多角度审查（Mode 3）
**评审人数**：3 位（互补视角，并行独立审查）

| 专家 | 评级 | 要点 |
|------|------|------|
| 源码对照专家（Code Match Reviewer） | 🔴 | ✅ Helper 函数语法正确 + parseMarginToTwips 走 number 分支正确 + 验证脚本预期值正确<br>🔴 **所有 old_string 缩进与实际严重不匹配**：方案展示 0/4 空格，实际嵌套 16-24 空格 → Edit 全部会失败<br>🔴 **Task 2 Step 2.1/2.2 编号与执行顺序矛盾**：Step 2.1 用 columnWidthsHtml，Step 2.2 才定义，executing-plans 按编号执行会产生 ReferenceError<br>🔴 **L1290 含 16 个 trailing space**：方案 Task 4 Step 4.1 old_string 把它当空行，Edit 会失败 |
| docx@9.5.1 API 专家（API Correctness Reviewer） | 🟡 | ✅ `width.type=DXA` / `columnWidths: number[]` / `layout: TableLayoutType.FIXED`（字符串字面量）/ `cell.width = 跨列合计` 全部 API 用法正确<br>✅ rowSpan>1 时 docx 库自动插入 CONTINUE cell（cjs L15226-15235），无 width 无需方案处理<br>🟡 verify 脚本 Check 2 正则脆性（绑定 type→size 顺序），建议解耦<br>🟡 PERCENTAGE→DXA 副作用：固定 contentWidth，模板若改 page.size（横向/A3）会失败，需 follow-up |
| 边界 + 副作用专家（Edge Case & Regression Reviewer） | 🟡 | ✅ parseMarginToTwips 正则覆盖所有模板 margin 值（cm 字符串 + number twips 全收）<br>✅ colspan reduce 安全（已传 initial value 0；parseHtmlTable L391/L418 保证 colSpan ≥ 1）<br>✅ PDF 路径（bin/capture.js）走 docx-preview 浏览器预览，无回归风险<br>🟡 **Task 2 Step 2.1/2.2 位置描述矛盾**（"L548 之前" vs "L488 之前"），实施者会困惑<br>🟡 **必须明确保留 L530-L535 columnSpan/rowSpan** 字段，避免误删<br>🟡 空表格守护（mdColCount=0 / colCount=0）两处都需要 |

**评审共识修复点**（已落实到方案文档）：

| # | 修复项 | 落实位置 |
|---|--------|----------|
| 1 | 加"⚠️ 执行说明"段（强制 Read 复制原始缩进，警示 trailing whitespace） | 方案文档顶部"⚠️ 执行说明（A5 评审强制要求）" |
| 2 | Task 2 物理重写为正向顺序（Step 2.1 先插 columnWidthsHtml + colCount=0 守护，Step 2.2 再改 cellConfig） | Task 2 Step 2.1 / 2.2 |
| 3 | Task 2 Step 2.2 加"⚠️ 保持不动"段，明确保留 L529-L535 的 cellConfig.rowSpan / columnSpan 赋值 | Task 2 Step 2.2 末段 |
| 4 | Task 4 Step 4.1 改插入点为 L1291 注释之前（避开 L1290 trailing whitespace） + 加 mdColCount=0 空表守护 | Task 4 Step 4.1 |
| 5 | Task 5 verify 脚本 Check 2 改为属性顺序无关的正则（独立 lookahead）| Task 5 Step 5.1 |
| 6 | Task 6 加 6.4b（rowspan 边界用例验证 docx 自动 CONTINUE cell） + 6.4c（PDF 路径回归） | Task 6 Step 6.4a/6.4b/6.4c |

**保留为 follow-up**（不阻塞当前修复）：
- 横向纸张/A3/Letter 支持（calcContentWidth 读 page.size 而非硬编码 A4）
- `.gitignore` 增加 `docs/test-output.docx`（避免二进制 diff 噪声）
- helpers 函数 export 出来供未来 wrapper 化复用

**综合结论**：❌ **不通过 Round 1**（1 个 🔴 + 2 个 🟡）

修订要求全部已落实到方案文档（见上表 6 项 + Task 2/Task 4/Task 5/Task 6 实际改动）。Task 2 已物理重写为正向顺序，不再有 "Step 修正"说明。Task 4 Step 4.1 已避开 L1290 trailing whitespace 问题。

**修订状态**：6 项修订全部已落实，方案达可执行成熟度。

---

### Round 2

**评审时间**：2026-05-27 21:35
**评审模式**：元评审快速复查
**评审人数**：1 位（chain 自检）

| Round 1 反馈项 | 落实状态 | 验证 |
|---|---|---|
| 1. 加"⚠️ 执行说明" | ✅ | 文档顶部新增独立小节，3 条强约束（Read 复制原始缩进 / 警示 trailing whitespace / new_string 也需对齐） |
| 2. Task 2 物理重写 | ✅ | Step 2.1 = "插入 columnWidthsHtml + colCount=0 守护"（含完整 16 空格缩进示意），Step 2.2 = "改 cellConfig"，编号与逻辑顺序一致 |
| 3. Task 2 保留 columnSpan/rowSpan 警示 | ✅ | Step 2.2 末段 "⚠️ 保持不动"段，明确说明新 width 与现有 columnSpan 互补 |
| 4. Task 4 避开 L1290 | ✅ | Step 4.1 改为"L1291 // Process ... 之前"，old_string 不再触碰 L1290 trailing whitespace |
| 5. verify 脚本正则解耦 | ✅ | Check 2 改为独立 lookahead，无属性顺序耦合 |
| 6. Task 6 加 rowspan + PDF 用例 | ✅ | 6.4b（rowspan + vMerge 检查） + 6.4c（bin/capture.js PDF 路径） |

**新引入问题检查**：
- Step 4.1 文字偏长（3 种方案描述），但逻辑正确且明确推荐"早返回 continue"版本 → 实施者可直接采纳推荐版本，不影响可执行性
- 其他无新引入问题

**综合评级**：🟢

**综合结论**：✅ **通过 Round 2**

3 位 Round 1 反馈全部已采纳并落实到方案。方案具备执行所需的全部信息（精确行号、缩进警示、保留字段提示、空表守护、边界用例）。可进入 A6 代码实施。
