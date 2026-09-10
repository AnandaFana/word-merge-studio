# 开源方案调研与选型

调研日期：2026-09-10。目标是左右 Word 对比、人工逐项合并并指定一侧格式，而不是简单拼接两份文件。

| 项目 | 官方说明与许可 | 本项目决策 |
| --- | --- | --- |
| [docxjs / docx-preview](https://github.com/VolodymyrBaydalka/docxjs) | Apache-2.0；将 DOCX 渲染到 HTML。官方说明排版受 HTML 能力限制，不具备实时自动分页引擎；公开稳定入口为 `renderAsync` | 复用只读预览，固定使用 0.3.7。输出从原 OOXML 生成，避免 HTML 往返造成格式丢失 |
| [docxcompose](https://github.com/4teamwork/docxcompose) | MIT；主用途是追加/拼接 `.docx`；默认尝试采用第一份文档样式，页眉页脚由第一份文档控制 | 不直接用作交互式 diff 引擎。将来可参考它的资源、编号和样式迁移处理 |
| [Open-Xml-PowerTools](https://github.com/OpenXmlDev/Open-Xml-PowerTools) | .NET/Open XML 生态的文档处理工具，提供文档比较及组合相关功能 | 作为后续复杂 Office 文档后端的候选。本次采用纯浏览器方案，避免首次运行要求 .NET 服务 |
| [Mammoth.js](https://github.com/mwilliamson/mammoth.js) | 以语义 HTML 转换为目标，强调简洁语义而非精确复制源文档格式 | 不用它重建导出文档；这种转换方向不符合保留一侧格式的核心要求 |

选择 React + TypeScript + JSZip + OOXML 定点修改，有两个直接好处：用户文件无需离开本机；底稿样式和非正文资源可以原样保留。代价是必须明确限制不具备关系迁移能力的内容，不应声称支持任意 DOCX 无损合并。

现阶段没有直接引入第三方完整 Word 编辑器。项目提供差异审阅和内容选择，不实现 Word 的自由编辑、三方合并、原生修订记录或完整分页引擎。

以上链接为上游项目，功能与许可应在将来升级依赖或商用分发前按实际版本复核。已安装依赖的实际许可文件位于各自 `node_modules/<package>/LICENSE*`，版本由 `package-lock.json` 固定。
