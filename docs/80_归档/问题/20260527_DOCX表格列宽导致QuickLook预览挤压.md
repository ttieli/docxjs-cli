# DOCX 表格列宽写成 100 导致 Apple QuickLook 预览挤压

> 创建时间：2026-05-27 21:03
> 管线：superpower-chain Pipeline A
> 状态：分析中
> 外部源：`BugFix_Factory/20260527_docxjs_DOCX表格列宽写成100导致QuickLook预览挤压/`

## 现象

`docxjs` 生成的表格类 DOCX，在 iOS 和 macOS QuickLook 预览中表格整体被压成页面左侧一条极窄竖条，中文字符纵向堆叠互相覆盖，基本不可读。

同一文件在 iOS + macOS 两端都异常，确认不是某一端预览组件的 UI 问题。

样本文件可正常在 Microsoft Word / WPS 中打开，但 Apple QuickLook 严格按 OOXML 规范渲染，触发挤压。

## 复现条件

- **环境**：iOS 任意版本 + macOS 任意版本（凡是走系统 QuickLook 的场景）
- **复现样本**：`/Users/tieli/Library/Mobile Documents/com~apple~CloudDocs/HandOn/Outputs/conv_20260527_202124_0nvl/个人基本信息表_铁力_20260527.docx`
- **步骤**：
  1. 用 `docxjs` 生成任意含表格的 DOCX（markdown 表格 / HTML 表格 / 代码块）
  2. 在 iOS Files / macOS Finder 中预览该 DOCX（QuickLook）
  3. 观察：表格列宽塌缩为左侧一条窄竖条
- **频率**：必现（所有 docxjs 生成的表格 DOCX）

## 相关日志/证据

样本 DOCX `word/document.xml` 实测（`unzip -p ... | grep tblW/gridCol/tcW`）：

```text
tblW w:type="pct" w:w="100%"      ← 非法：pct 单位应为 1/50 percent（100%=5000），不是字面量 "100%"
gridCol w:w="100"                  ← 异常窄：100 twips ≈ 1.76mm
gridCol w:w="100"
gridCol w:w="100"
gridCol w:w="100"
... （所有表格所有列都是 100）
（首行 TableCell 无 w:tcW 元素，QuickLook 无单元格宽度兜底）
```

对照：同目录下 pandoc 生成的 `..._修复预览.docx` 表现正常，其 `tblGrid` 是 1980 / 2640 / 3960 / 2457 等合理 twips 值。

## 影响范围

- **影响功能**：所有 `docxjs` 表格生成路径（markdown table / HTML table / 代码块包装表）
- **影响的 Apple 渲染器**（A3 评审补充：远超 QuickLook，Apple 全家桶共享同一 OOXML 解析器）：
  - macOS QuickLook（Finder 预览、空格预览、Spotlight 缩略图）
  - iOS / iPadOS Files / Quick Look
  - Apple Mail.app 附件预览
  - Apple Pages（打开后表格挤压）
  - iCloud.com 在线预览
  - HandOn / HandOff 等使用系统 QuickLook 的 App
- **不受影响**：Microsoft Word for Mac / WPS / LibreOffice / Google Docs（识别 `100%` 字面量后回退到 autofit + 内容驱动列宽）
- **严重程度**：**P1 严重**
  - 不影响数据正确性，但影响整个 Apple 生态的用户体验
  - 凡使用 docxjs 输出表格类 DOCX 的下游 App 都受波及

---

## 根因分析（A2 完成）

### 表格构造点全量清单

`rg -n "new Table\(" lib/ bin/ --type js` 全量扫描后确认 **只有 3 处** Table 构造（无遗漏）：

| # | 文件 | 行号 | 场景 | 列数 |
|---|------|------|------|------|
| 1 | `lib/core.js` | L558-L562 | `htmlTableToDocx`：HTML 表格（含 rowspan/colspan） | `tableData.colCount`（动态） |
| 2 | `lib/core.js` | L1176-L1196 | 代码块包装为单元格表格（视觉效果） | 1（单列单行） |
| 3 | `lib/core.js` | L1332-L1336 | markdown table（标准 `\| ... \|` 语法） | 由首行 cells 数决定 |

三处统一使用 `width: { size: 100, type: WidthType.PERCENTAGE }`，无任何 `columnWidths`、TableCell 无 `width`。

### docx 库源码实证

**实证 1：`WidthType` 枚举值**（`node_modules/docx/dist/index.d.ts:2987-2992`）

```typescript
export declare const WidthType: {
    readonly AUTO: "auto";
    readonly DXA: "dxa";
    readonly NIL: "nil";
    readonly PERCENTAGE: "pct";   // 没有独立的 PCT，PERCENTAGE 就映射到 "pct"
};
```

**实证 2：`TableWidthElement` 序列化逻辑**（`node_modules/docx/dist/index.cjs:14827-14841`）

```js
class TableWidthElement extends XmlComponent {
  constructor(name, { type: type2 = WidthType.AUTO, size }) {
    super(name);
    let tableWidthValue = size;
    if (type2 === WidthType.PERCENTAGE && typeof size === "number") {
      tableWidthValue = `${size}%`;    // ← 关键：number 类型直接拼 %
    }
    this.root.push(new NextAttributeComponent({
      type: { key: "w:type", value: type2 },
      size: { key: "w:w", value: measurementOrPercentValue(tableWidthValue) }
    }));
  }
}
```

→ 传 `size: 100, type: PERCENTAGE` 输出 `w:type="pct" w:w="100%"`。
→ 即使改成 `size: 5000`，也只是输出 `w:w="5000%"`，仍违反 OOXML 规范（pct 类型的 w:w 应为纯整数）。
→ **结论**：`PERCENTAGE` 类型在 docx@9.5.1 中无法输出合法 OOXML（OOXML 规范本身允许 `pct=5000`，但 docx 库的实现 bug 把数字直接拼了 `%` 而非按规范输出整数），必须改用 `DXA`。

**实证 3：`Table` 构造器的 columnWidths 默认值**（`node_modules/docx/dist/index.cjs:15185-15216`）

```js
class Table extends FileChild {
  constructor({
    rows,
    width,
    columnWidths = Array(Math.max(...rows.map((row) => row.CellCount))).fill(100),
    // ...
  }) {
    // ...
    this.root.push(new TableGrid(columnWidths));  // 默认 100 twips × N 列
  }
}
```

→ 不传 `columnWidths` 时，docx 库默认每列 100 twips。这就是 `gridCol w="100"` 的来源。
→ 100 twips ≈ 1.76mm，对中文字符宽度（~280 twips）严重不足。

**实证 4：TableCell 不传 `width` 不输出 `w:tcW`**

样本 DOCX `unzip -p ... | grep tcW` 返回 0 结果 → 确认 TableCell 缺 `w:tcW` 元素。
**A3 评审修订**：对照 pandoc 修复文件（同样无 `w:tcW`）实证 → `w:tcW` 缺失**不是核心根因**，根据 ECMA-376 §17.4.72，缺失时阅读器按 `auto` 处理，由 tblGrid 自动分配。真正决定 QuickLook 渲染的是 `tblGrid` 的几何真值（100 twips）。

**实证 5：未传 `layout` 不输出 `w:tblLayout`**（A3 评审补充）

docx 库 `Table` 构造器（L15208）的 `layout` 参数未传时不输出 `w:tblLayout`，默认 autofit 模式。QuickLook 在 autofit 下可能基于内容反向重算列宽 → 即使 gridCol 正确也可能被撑挤。pandoc 修复文件中关键表使用 `<w:tblLayout w:type="fixed"/>` 强制按 gridCol 渲染。

### 页面尺寸/边距上下文（修复方案依赖）

- `Document` 构造（`lib/core.js:1389-1395`）只传 `page.margin`，**未传 `page.size`** → docx 库默认 A4：**11906×16838 twips**
- `margin` 从 `currentStyle.margin` 读取，模板中是 cm 字符串（如 `"2.54cm"`、`"2cm"`、`"3cm"`）
- 单位换算：1cm = 567 twips，1in = 1440 twips，1mm = 56.7 twips
- 默认 A4 + 2.54cm 边距 → contentWidth = 11906 - 567×2.54×2 ≈ **9026 twips**
- 默认 A4 + 2cm 边距 → contentWidth = 11906 - 567×2×2 ≈ **9638 twips**

### 证据链（A3 评审已修订）

1. **入口**：用户 markdown 含表格 → `docxjs` → `lib/core.js` 中三处 `new Table()`
2. **触发 1**（tblW）：`width: { size: 100, type: WidthType.PERCENTAGE }` → docx 库序列化为 `w:type="pct" w:w="100%"`（字面拼 %）。注：传 `size: 5000` 也会被序列化为 `"5000%"` 仍违规，docx@9.5.1 的 PERCENTAGE 路径完全不可用。
3. **触发 2**（tblGrid，**核心根因**）：未传 `columnWidths` → docx 库默认填 `Array(N).fill(100)` twips → 输出 `w:gridCol w:w="100"` × N。100 twips ≈ 1.76mm，**小于单个中文字符宽度**（中文 ~280 twips），任何文本都会被挤压换行。
4. **触发 3**（tblLayout，**A3 评审补充**）：Table 未传 `layout` → 不输出 `w:tblLayout` → 默认 autofit 模式，QuickLook 可能基于内容反向重算 → 即使 gridCol 正确也可能被撑挤。
5. **可选项**（tcW，**A3 评审降级**）：TableCell 未传 `width` → 不输出 `w:tcW`。**注意**：pandoc 对照文件实证 —— pandoc 修复文件**完全没有 `w:tcW`** 却渲染正常。`w:tcW` 缺失**不是核心根因**（ECMA-376 §17.4.72 规定缺失时按 `auto` 处理，由 tblGrid 自动分配）。
6. **QuickLook 渲染机制**（A3 评审补充）：
   - 优先级与 Word 相反：**`tblGrid` + `tcW` 是几何真值，`tblW` 只是建议宽度**
   - Word 优先 `tblW` 并 autofit 重排 `tblGrid`；QuickLook 优先 `tblGrid` 保守渲染
   - `w:w="100%"` 字面量被严格解析器丢弃 → 退化到 `tblGrid` 100 twips
   - 100 twips × N 列 → 文本逐字换行 → 行高被 line-spacing 撑开 → 视觉上"字符纵向堆叠互相覆盖"
   - 结果：表格压成左侧 ~N×1.76mm 宽的竖条
7. **Word/WPS 容错**：识别 `100%` 字面量后**回退到 autofit + 内容驱动列宽**（不是真按 100% 解析）→ 表格正常
8. **pandoc 修复文件的正确写法**（实测 `unzip -p ... | grep tblW/gridCol/tblLayout`）：
   ```
   tblW w:type="auto" w:w="0"            # 大多数表格用 auto
   gridCol w:w="1980" / 2640 / 3960 ...  # 合理 twips 值
   ----
   tblW w:type="pct" w:w="5000"          # 关键表用 pct=5000（合法 OOXML 整数）
   tblLayout w:type="fixed"              # ★ 强制按 gridCol 渲染，不重协商
   gridCol w:w="2457" / 3004 / 2457
   ```
   注意 pandoc **完全不输出 `w:tcW`** —— 强化了"tcW 非必须"的结论。
9. **影响一致性**：lib/core.js 三处一致写错 → 所有表格场景（HTML/markdown/code）都受影响

### 为什么之前没发现

- 主要测试环境是 Word/WPS/docx-preview，这些工具对 `100%` 字面量容错
- QuickLook 是 Apple 平台系统级预览，docxjs 用户在 Apple 端打开预览时才暴露
- docx 库这个 bug 已有上游 issue（GitHub dolanmiu/docx）但未修复，docxjs 没意识到要绕开

### 根因结论（A3 评审已修订）

**一句话根因**：`lib/core.js` 三处 `new Table()` 调用使用 `WidthType.PERCENTAGE` 触发 docx@9.5.1 库 bug 输出非法 `w:w="100%"`；**更关键**的是未传 `columnWidths` 让 docx 库默认每列 100 twips（≈1.76mm），叠加未指定 `tblLayout` 走 autofit 模式 → Apple 全家桶 OOXML 解析器（QuickLook/Pages/Mail/iCloud/Spotlight 共享）按 `tblGrid` 为几何真值渲染时表格塌缩为竖条。

**修复必选四管齐下（A3 评审修订）**：

| # | 项 | 必要性 | 说明 |
|---|---|---|---|
| 1 | 三处 Table 的 `width` 改用 `WidthType.DXA` + 计算 contentWidth twips | **必须** | 输出合法 `w:tblW w:type="dxa"` |
| 2 | 给每个 Table 显式传入 `columnWidths`（按列数合理分配，总和 = contentWidth） | **必须** | 让 `w:gridCol` 输出合理 twips，是 QuickLook 渲染的几何真值 |
| 3 | 给每个 Table 加 `layout: TableLayoutType.FIXED` | **必须**（A3 补充） | 输出 `w:tblLayout w:type="fixed"`，禁止 autofit 重协商，pandoc 关键表的实证做法 |
| 4 | 给每个 TableCell 加 `width: { size, type: DXA }`（与列宽对应） | **可选加固** | pandoc 实证：完全省略 tcW 也能正常渲染。加上更稳，但非必须 |

**依赖**：需新增 margin 字符串解析函数（cm/in/mm → twips）和 contentWidth 计算函数。

**为什么不用 pct=5000 + auto layout**：pandoc 这么用没问题，但 docx@9.5.1 的 PERCENTAGE 路径有 bug（即使传 5000 也输出 `"5000%"` 仍违规）。短期 DXA 方案最稳；等上游修复后可在 wrapper 函数留 `useAutoLayout` 开关切回 pct 5000（对横向/A3 等非标准纸张更友好）。

### 验证建议（A3 评审补充）

修复后应做的验证手段（A6 实施时执行）：

| 验证项 | 命令/方法 | 通过标准 |
|--------|-----------|----------|
| **XML 结构断言** | `unzip -p output.docx word/document.xml \| grep -oE 'tblW\|gridCol\|tblLayout'` | 不存在 `w:w="100%"`；`tblGrid` 列宽合理（非全 100）；`tblLayout w:type="fixed"` 必现 |
| **数值一致性** | 同上 + 计算 | `sum(gridCol)` ≈ contentWidth（A4 + 2.54cm 边距 ≈ 9026 twips） |
| **QuickLook 缩略图** | `qlmanage -t -s 1024 -o /tmp/qltest output.docx` | 生成的 PNG 缩略图中表格宽度占页面比例 > 60% |
| **QuickLook 全屏预览** | `qlmanage -p output.docx` | 视觉确认表格不再压成左侧竖条 |
| **多解析器交叉** | 用 Apple Pages / macOS Word / LibreOffice / Google Docs 分别打开 | 表格在所有解析器中正常显示 |
| **边界用例** | 单元格含超长中英混排、合并单元格（rowspan/colspan）、A4 横向、自定义边距模板 | 各种边界场景下不再挤压 |

回归样本库：把当前 bug 样本和 pandoc 修复样本加入测试 fixture，作为后续修改的基线。

---

---

## 评审记录

### Round 1

**评审时间**：2026-05-27 21:15
**评审模式**：辩论模式（Mode 2）
**评审人数**：3 位（独立选角 + 互相质疑）

| 专家 | 评级 | 要点 |
|------|------|------|
| OOXML 规范专家（ECMA-376） | 🟡 | ✅ `w:w="100%"` 确实违反 ECMA-376（`ST_TblWidth pct` 类型要求 `w:w` 为 1/50% 整数，5000=100%）<br>⚠️ **质疑 tcW 必须性**：ECMA-376 §17.4.72 规定 `tcW` 缺失时按 `auto` 处理，pandoc 修复文件实证（无 tcW 却正常）→ `tcW` 不是核心根因，应降级<br>⚠️ Word/WPS 不是"按 100% 渲染"，而是回退到 autofit + 内容驱动列宽<br>✅ DXA + 计算 twips 规范合规，立刻可用 |
| docx-js 库源码专家 | 🟡 | ✅ 实证 `TableWidthElement` L14831 字面拼 `%`：`size:100`→`"100%"`，`size:5000`→`"5000%"` 仍违规<br>✅ `columnWidths` 默认 `Array.fill(100)` twips 确认（L15190）；显式传值后 `w:gridCol` 用传入值<br>✅ 全量扫描确认只有 3 处 `new Table()`，无遗漏；importDocx 走 IPC 不构造 Table<br>⚠️ **遗漏 `layout: TableLayoutType.FIXED`**：未在方案中提及，但 docx 库 L15116-15131 支持，输出 `w:tblLayout w:type="fixed"` 强制按 gridCol 渲染，pandoc 实证使用<br>⚠️ **docx 版本应为 9.5.1，不是 9.1.0**（已修正）<br>⚠️ 实施陷阱：colspan/rowspan 需按 grid 列展开、浮点 contentWidth 需 Math.floor、cell.width 必须严格等于对应列 columnWidths |
| Apple QuickLook 渲染专家 | 🟡 | ✅ Apple 全家桶（QuickLook/Pages/Mail/iCloud/Spotlight/iOS Files）共享同一 OOXML 解析器，影响远超 QuickLook<br>✅ 优先级与 Word 相反：`tblGrid`+`tcW` 是几何真值，`tblW` 只是建议宽度<br>✅ 100 twips × N 列 → 中文 ~280 twips 单字都放不下 → 逐字 wrap + line-spacing 撑开 → 视觉"堆叠覆盖"<br>⚠️ **必须加 `w:tblLayout w:type="fixed"`**（pandoc 修复文件实证）：autofit 模式下即使有合理 tblGrid 也可能被内容反向重算<br>⚠️ templates/common_styles.json 无 width 字段 → 用户无法通过模板掩盖此 bug<br>📌 验证建议：`qlmanage -t/-p/-z` + 像素级宽度断言（表格 bbox/页面 > 0.6） |

**互相质疑收敛的共识**（三位达成一致）：

1. **核心根因**：`tblGrid` 100 twips 是直接致挤压因，`w:w="100%"` 是次要触发器（让 tblW 失效后才暴露 tblGrid 问题）
2. **修复必选项**：`WidthType.DXA` + 显式 `columnWidths` + `layout: TableLayoutType.FIXED`（三位一致）
3. **可选加固**：TableCell `width` 加 `w:tcW`（pandoc 实证无 tcW 也正常，但加上更稳）
4. **影响范围**：Apple 全家桶 OOXML 解析器全部受影响，不仅 QuickLook

**综合结论**：❌ **不通过 Round 1**（3 个 🟡）

修复要求：
- (a) 修正 docx 版本号 9.1.0 → 9.5.1
- (b) 增加 `tblLayout=fixed` 作为必须修复项（修复方案从三管齐下 → 四管齐下）
- (c) 把 `w:tcW` 从必须项降级为可选加固，说明 pandoc 实证依据
- (d) 修正"Word/WPS 容错"措辞为"回退到 autofit + 内容驱动"
- (e) 扩展影响范围为 Apple 全家桶
- (f) 补充验证建议（qlmanage + 像素级断言）

**修订状态**：已采纳全部 6 项建议，根因结论已重写为"四管齐下"，证据链已重组（核心根因明确为 tblGrid，tcW 降级为可选加固，新增 tblLayout 触发项）。

---

### Round 2

**评审时间**：2026-05-27 21:17
**评审模式**：元评审复查（验证 Round 1 反馈是否完整采纳）
**评审人数**：1 位（Senior Reviewer）

| 反馈项 | 采纳状态 | 位置/证据 |
|---|---|---|
| (a) docx 9.1.0 → 9.5.1 | ✅ | L106 "PERCENTAGE 类型在 docx@9.5.1 中无法输出合法 OOXML"；L147 "docx@9.5.1 的 PERCENTAGE 路径完全不可用"；L191 "docx@9.5.1 的 PERCENTAGE 路径有 bug"。全文 9.1.0 已无残留 |
| (b) 增加 tblLayout=fixed 必须项 | ✅ | L132-134 新增"实证 5：未传 layout 不输出 w:tblLayout"；L149 证据链触发 3 明确为 tblLayout；L186 修复表第 3 项 `layout: TableLayoutType.FIXED` 标为"**必须**(A3 补充)"；L178 一句话根因含"叠加未指定 tblLayout 走 autofit 模式" |
| (c) tcW 降级为可选 + pandoc 实证 | ✅ | L129-130 实证 4 标注"A3 评审修订：w:tcW 缺失不是核心根因"+ECMA-376 §17.4.72 引用；L150 证据链"可选项（tcW，A3 评审降级）"+pandoc 对照文件实证；L167 "pandoc 完全不输出 w:tcW —— 强化了 tcW 非必须的结论"；L187 修复表第 4 项标"可选加固" |
| (d) Word/WPS 措辞修正 | ✅ | L52 "不受影响"行："识别 100% 字面量后回退到 autofit + 内容驱动列宽"；L157 证据链第 7 点 "Word/WPS 容错：识别 100% 字面量后**回退到 autofit + 内容驱动列宽**（不是真按 100% 解析）" |
| (e) 影响范围扩展 Apple 全家桶 | ✅ | L45-51 影响范围列出 6 项 Apple 渲染器（QuickLook Finder/空格预览/Spotlight、iOS/iPadOS Files、Apple Mail、Apple Pages、iCloud.com、HandOn）；L151-156 证据链第 6 点"QuickLook 渲染机制"展开 Apple 解析器优先级；L178 根因结论"Apple 全家桶 OOXML 解析器（QuickLook/Pages/Mail/iCloud/Spotlight 共享）" |
| (f) 验证建议补充 | ✅（修订后） | 已采纳 Round 2 建议，新增"验证建议（A3 评审补充）"独立小节，6 项验证手段成表（XML 断言 / 数值一致性 / qlmanage 缩略图 / qlmanage 预览 / 多解析器交叉 / 边界用例） |

**整体一致性检查**：
- ✅ "四管齐下" 口径在三处完全统一：一句话根因（L178）、修复必选表（L184-187）、证据链触发 1-3+可选项（L147-150）
- ✅ "核心根因 = tblGrid 100 twips" 在 L130 / L148 / L213（Round 1 共识）三处呼应
- ✅ tblW(次要)、tblLayout(必须新增)、tcW(可选降级) 三项定位前后一致
- ✅ pandoc 实证证据在 L40 / L130 / L150 / L158-167 多处互相印证

**新引入问题检查**：
- ⚠️ L193-195 出现两个连续 `---` 分隔符（L193 之后空行又一个 L195），轻微 markdown 格式冗余，不影响阅读
- ✅ L106 已修订：补上"（OOXML 规范本身允许 pct=5000，但 docx 库的实现 bug 把数字直接拼了 % 而非按规范输出整数）"括注，消除歧义
- ✅ 无事实性错误或证据反转

**综合评级**：🟢（经 chain 协调修订后）

**综合结论**：✅ **通过 Round 2**

Round 2 提出的两项微调已落实：
1. ✅ 新增"验证建议（A3 评审补充）"独立小节，含 6 项验证手段表格（qlmanage / XML 断言 / 像素比例 / 多解析器矩阵 / 边界用例 / 回归样本）
2. ✅ L106 已补充括注："OOXML 规范本身允许 pct=5000，但 docx 库的实现 bug 把数字直接拼了 %"

根因分析已达可进入 A4 修复方案设计的成熟度。
