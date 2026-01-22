# DocxJS TODO

---

## 历史待办
- 完成 HTML 预览稳健性验证：导入多样 HTML（含外链图片/样式、纯文本），确认无报错、可导出 PNG/PDF。
- CLI 无头截图 (`docxjs-capture`) 增强：增加错误提示（缺少浏览器/端口占用）和示例文档。
- 签名与发布：准备 macOS 签名/公证证书，重打带签名的 DMG。
- fsevents 依赖：在非 iCloud 路径验证正常安装；必要时改为可选依赖或文档说明。
- 回归测试：Markdown→Docx、HTML 预览/导出、CLI Capture 端到端脚本化验证。

---

# 完整 Markdown 支持开发计划

> 分支: `feature/full-markdown-support`
> 创建时间: 2026-01-21
> 目标: 实现完整的 Markdown 格式支持

---

## 当前支持状态

| 格式 | 状态 | 说明 |
|------|------|------|
| 标题 H1-H6 | ✅ 已支持 | 完整支持 |
| 粗体/斜体 | ✅ 已支持 | `**bold**` `*italic*` |
| 链接 | ✅ 已支持 | `[text](url)` |
| 图片 | ✅ 已支持 | 包括本地和网络图片 |
| 表格 | ✅ 已支持 | GFM 表格语法 |
| 无序列表 | ✅ 已支持 | 支持多层级嵌套 |
| 代码块 | ✅ 已支持 | 包括语法高亮标记 |
| 行内代码 | ✅ 已支持 | `` `code` `` |
| Mermaid 图表 | ✅ 已支持 | 渲染为图片 |
| 数学公式 | ✅ 已支持 | `$...$` `$$...$$` (KaTeX) |
| 有序列表 | ✅ 已支持 | `1. 2. 3.` 支持嵌套 |
| 嵌套列表 | ✅ 已支持 | 多层级列表 |
| 引用块 | ✅ 已支持 | `> quote` 支持嵌套 |
| 分隔线 | ✅ 已支持 | `---` |
| 删除线 | ✅ 已支持 | `~~text~~` |
| 任务列表 | ✅ 已支持 | `- [ ] task` |
| 脚注 | ✅ 已支持 | `[^1]` 显示为上标 |
| 上下标 | ✅ 已支持 | `H~2~O` `x^2^` |
| 表格对齐 | ⚠️ 部分 | 解析了但未完全应用 |

---

## 开发阶段

### 阶段 1: 基础扩展
**预计工时**: 1-2 小时
**状态**: [x] 已完成 (2026-01-21)

#### 1.1 安装依赖包
- [x] `markdown-it-texmath` - 数学公式解析
- [x] `markdown-it-footnote` - 脚注解析
- [x] `markdown-it-sub` - 下标解析
- [x] `markdown-it-sup` - 上标解析
- [x] `markdown-it-task-lists` - 任务列表解析
- [x] `katex` - 数学公式渲染

```bash
npm install markdown-it-texmath markdown-it-footnote markdown-it-sub markdown-it-sup markdown-it-task-lists katex
```

#### 1.2 配置 markdown-it 插件链
- [x] 修改 `lib/core.js` 初始化代码
- [x] 添加插件加载逻辑

#### 1.3 实现分隔线 (hr)
- [x] 处理 `hr` token
- [x] 添加底部边框段落样式

#### 1.4 实现删除线 (strikethrough)
- [x] 在 `processInline` 中添加 `s_open/s_close` 处理
- [x] TextRun 添加 `strike: true` 属性

---

### 阶段 2: 列表系统
**预计工时**: 2-3 小时
**状态**: [x] 已完成 (2026-01-21)

#### 2.1 重构列表状态管理
- [x] 创建 `listStack` 跟踪列表层级
- [x] 区分有序/无序列表类型

#### 2.2 实现有序列表
- [x] 处理 `ordered_list_open/close` token
- [x] 配置 Document numbering
- [x] 支持起始数字 (start 属性)

#### 2.3 实现嵌套列表
- [x] 跟踪列表嵌套深度
- [x] 支持多层级缩进样式
- [x] 混合有序/无序嵌套

#### 2.4 实现任务列表
- [x] 检测任务列表项
- [x] 渲染 ☐/☑ 符号 (通过 markdown-it-task-lists 插件)

---

### 阶段 3: 引用与格式
**预计工时**: 1-2 小时
**状态**: [x] 已完成 (2026-01-21)

#### 3.1 实现引用块 (blockquote)
- [x] 处理 `blockquote_open/close` token
- [x] 添加左侧边框样式
- [x] 添加背景色
- [x] 支持嵌套引用

#### 3.2 实现上下标
- [x] 处理 `sub_open/close` token
- [x] 处理 `sup_open/close` token
- [x] TextRun 添加 `subScript/superScript` 属性

#### 3.3 修复表格对齐
- [ ] 从 token.attrs 提取对齐信息 (待完善)
- [ ] 应用到单元格 Paragraph alignment (待完善)

---

### 阶段 4: 数学公式
**预计工时**: 2-3 小时
**状态**: [x] 已完成 (2026-01-21)

#### 4.1 集成 KaTeX 渲染器
- [x] 创建 `renderMathLocally` 函数
- [x] 使用 Playwright 截图
- [x] CDN KaTeX 集成

#### 4.2 实现行内公式
- [x] 处理 `math_inline` token
- [x] 图片插入到行内
- [x] 调整图片垂直对齐

#### 4.3 实现块级公式
- [x] 处理 `math_block` token
- [x] 居中显示
- [x] 添加上下间距

#### 4.4 优化渲染性能
- [x] 实现公式缓存 (mathCache Map)
- [ ] 批量渲染优化 (待完善)

---

### 阶段 5: 脚注系统
**预计工时**: 1-2 小时
**状态**: [x] 已完成 (2026-01-21)

#### 5.1 实现脚注收集
- [x] 第一遍遍历收集脚注定义
- [x] 存储脚注 ID 和内容映射

#### 5.2 实现脚注引用与渲染
- [x] 处理 `footnote_ref` token
- [x] 使用上标格式显示脚注引用
- [x] 构建 Document footnotes 配置

---

### 阶段 6: 测试与优化
**预计工时**: 1-2 小时
**状态**: [x] 已完成 (2026-01-21)

#### 6.1 创建测试文档
- [x] 创建 `test-all-formats.md` 包含所有格式
- [x] 测试边界情况

#### 6.2 性能优化
- [ ] Playwright 浏览器复用
- [ ] 图片缓存机制

#### 6.3 错误处理
- [ ] 添加格式渲染失败的 fallback
- [ ] 完善日志输出

---

## 测试验证清单

### 基础格式
- [ ] 粗体 + 斜体 + 删除线组合
- [ ] 行内代码 + 链接组合

### 列表
- [ ] 纯有序列表 (1-10项)
- [ ] 纯无序列表
- [ ] 三层嵌套混合列表
- [ ] 任务列表 (已完成/未完成)

### 引用
- [ ] 单层引用
- [ ] 嵌套引用 (2-3层)
- [ ] 引用内包含列表

### 数学公式
- [ ] 简单行内公式: `$E=mc^2$`
- [ ] 复杂行内公式: `$\sum_{i=1}^{n} x_i$`
- [ ] 块级公式: 积分、矩阵
- [ ] 连续多个公式

### 表格
- [ ] 左/中/右对齐列
- [ ] 表格内粗体/斜体
- [ ] 表格内链接

### 脚注
- [ ] 单个脚注
- [ ] 多个脚注
- [ ] 脚注内格式化文本

### 综合
- [ ] 所有格式混合文档
- [ ] 大型文档 (100+ 段落)

---

## 文件修改清单

| 文件 | 修改类型 | 说明 |
|------|----------|------|
| `package.json` | 修改 | 添加新依赖 |
| `lib/core.js` | 重构 | 主要实现文件 |
| `public/vendor/katex/` | 新增 | KaTeX 本地资源 |
| `docs/test-all-formats.md` | 新增 | 测试文档 |

---

## 注意事项

1. **向后兼容**: 确保现有功能不受影响
2. **性能考虑**: 数学公式渲染可能较慢，需要缓存
3. **错误处理**: 渲染失败时显示原文本而非崩溃
4. **测试覆盖**: 每个阶段完成后进行测试

---

## 参考资源

- [markdown-it 文档](https://github.com/markdown-it/markdown-it)
- [docx 库文档](https://docx.js.org/)
- [KaTeX 文档](https://katex.org/docs/api.html)
- [markdown-it 插件列表](https://www.npmjs.com/search?q=keywords:markdown-it-plugin)
