# 项目状态

## [2026-05-27] Pipeline A: Bug 修复

- **标题**：DOCX 表格列宽写成 100 导致 Apple QuickLook 预览挤压
- **版本**：1.3.28 → 1.3.29
- **提交数**：7 个（5856ac8 / 4b32a4e / f4fca33 / ee2dca8 / c7762eb / 45ee587 / cf7638e）
- **文档**：已归档至 `docs/80_归档/{问题,设计}/20260527_*.md`
- **摘要**：修复 `lib/core.js` 三处 `new Table()` 调用，使用 `WidthType.DXA` + 显式 `columnWidths` + `TableLayoutType.FIXED` + TableCell `width`，让 docxjs 生成的 DOCX 表格在 Apple QuickLook / Pages / Mail / iCloud / Spotlight 等 OOXML 严格渲染器中不再压成左侧竖条。验证：6 个 Task 全部通过，3 模板 + 边界用例（10 列宽表 / colspan / rowspan）+ QuickLook 视觉对比全部 ✅。
