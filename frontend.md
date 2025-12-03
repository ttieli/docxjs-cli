当前架构速览

- bin/cli.js: CLI 入口，加载 templates/*.json，支持交互选择模板、--reference-doc 时调用
  style_extractor.py 提取样式并合并，再用 generateDocx 输出。
- lib/core.js: Markdown 渲染与 docx 生成，支持标题、列表、表格、链接、行内代码，高度依赖
  currentStyle 配置。
- style_extractor.py: 用 python-docx 读取参考 docx，抽取字体/字号/边距/表格边框等，返
  回 JSON。
- lib/template-manager.js: 提供模板 CRUD，但 CLI/服务未使用这套封装。
- server/app.js: 简单 Express API (/api/convert, /api/templates)，仅用模板或显式传入的
  styleConfig 生成 docx；当前不支持参考 docx，也没有鉴权/限流/体积限制。
- 模板存放在 templates/*.json，示例输出文件在根目录。

  改进与建议

- 统一模板加载：CLI 用 loadTemplates，服务端用 loadDefaultTemplates，与 lib/template-
  manager.js 三套逻辑重复；建议收敛到 TemplateManager，避免源不一致。
- Python 依赖隔离：CLI 运行时通过多种路径查找 Python，但服务端不支持参考 docx 提取；如果要
  前端调用“参考 docx”功能，需要在服务端包装调用 style_extractor.py，同时增加安全检查（文件
  类型/大小、超时）。
- 错误处理与校验：/api/convert 未限制请求体大小或校验 Markdown/模板名称；CLI 读取文件、写入
  时也未给出更清晰的错误提示（如输出路径不可写）。
- 测试与回归：npm test 占位，缺少最小的核心单测（Markdown→docx 快照或属性检查）和端到端用例
  （常用模板、表格、参考 docx）。
- 服务化准备：没有启动脚本/配置（如 npm run server），端口硬编码为 3000；建议使用环境变量并
  补一个 start:server 脚本。
- 文件结构：核心逻辑集中在 lib/core.js，可以按职责拆分（inline 处理、表格渲染、样式合并）以
  便扩展和测试。

  如果要加一个“最简单的前端”

- 复用现有 Express：在 server/app.js 下增设静态资源目录，提供一个单页（纯 HTML/JS/CSS）
  即可。
- 前端页面功能最小集：
  1. 一个 Markdown 文本域、可选模板下拉（通过 GET /api/templates 填充），可选上传参考
     docx（若后端支持）。
  2. 点击“转换”后调用 /api/convert，返回 blob 触发下载。
  3. 显示错误信息与基本进度状态。
- 后端配合调整：
  1. /api/convert 支持 multipart（Markdown/模板名/参考 docx），在服务端落地后调用
     style_extractor.py 合并样式。
  2. 设置 body/file 大小限制、mime 校验、超时与错误返回。
  3. 使用环境变量配置端口。
- 技术选型：无需框架，直接原生 HTML + fetch 即可；若需一点构建体验，可用 Vite + 轻量
  UI（React/Vue 任意），但保留“静态发布”的打包产物（dist/ 由 Express 静态托管）。
