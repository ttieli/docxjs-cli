const core = require('./core');

// 固定的样例文本，用于展示各种样式效果
const SAMPLE_MARKDOWN = `
# 样式预览示例文档 (一级标题)

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
| 3 | 后端接口联调 | 待开始 | API 对接 |

## 页面布局测试
此段落用于测试页面的左右边距。请尝试调整左侧配置面板中的页边距数值，观察文字距离页面边缘的距离变化。
`;

/**
 * 根据提供的样式配置生成预览文档的 Buffer
 * @param {Object} styleConfig 样式配置对象
 * @param {string} [customMarkdown] 可选的自定义 Markdown 内容
 * @returns {Promise<Buffer>} DOCX 文件 Buffer
 */
async function generatePreviewBuffer(styleConfig, customMarkdown) {
    // 如果提供了自定义内容，则使用自定义内容；否则使用默认样例
    const content = (customMarkdown && customMarkdown.trim().length > 0) 
        ? customMarkdown 
        : SAMPLE_MARKDOWN;

    // 复用 core.js 的生成逻辑
    // 第三个参数 baseDir 传 null 或 process.cwd()，因为样例中没有图片需要解析相对路径
    return await core.generateDocx(content, styleConfig, process.cwd());
}

module.exports = {
    generatePreviewBuffer
};
