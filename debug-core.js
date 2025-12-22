const { generateDocx } = require('./lib/core');
const fs = require('fs');

const sampleMD = `# 示例文档 (一级标题)

这是文档的正文部分。此处展示了**粗体**、*斜体*以及普通文本的行距和字体效果。通过修改左侧的样式配置，您可以实时预览正文的字体、字号、颜色和行间距的变化。

## 二级标题样例
二级标题通常用于章节划分。

### 三级标题样例
三级标题用于更细致的内容分组。

## 列表与表格展示

下面是一个无序列表：
- 列表项 1：测试缩进和对齐
- 列表项 2：测试字体一致性
- 列表项 3：包含 **粗体** 的列表项

下面是一个表格样式的展示：

| 序号 | 项目名称 | 状态 | 备注 |
| :--- | :--- | :--- | :--- |
| 1 | 系统架构设计 | 完成 | 核心模块 |
| 2 | 前端界面开发 | 进行中 | 样式调整 |
| 3 | 后端接口联调 | 待开始 | API 对接 |`;

(async () => {
    try {
        console.log("Starting generation...");
        const buffer = await generateDocx(sampleMD, {}, process.cwd());
        console.log("Buffer generated, size:", buffer.length);
        fs.writeFileSync('debug_output.docx', buffer);
    } catch (e) {
        console.error("Error:", e);
    }
})();
