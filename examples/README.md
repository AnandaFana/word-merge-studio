# Public examples / 公开样例

These fictional documents were created from scratch and are distributed under the project's MIT license. They contain no real project, organisation or personal data.

以下三份文档均为从零制作的虚构样例，可公开分享：

| 文件 / File | 内容 / Scenario | 特点 / Features |
| --- | --- | --- |
| [garden.docx](public/garden.docx) | 社区花园 / Community garden | Georgia、宋体，绿色标题 |
| [library.docx](public/library.docx) | 社区书屋 / Reading room | Arial、微软雅黑，绿色标题 |
| [cafe.docx](public/cafe.docx) | 街角咖啡 / Neighbourhood cafe | Calibri、宋体，棕色标题 |

Each has three corresponding sections (Overview, Goals, Schedule), second-level headings, one embedded PNG and a small table. Import all three into **Merge sections**, then click **Detect and regroup**. Or use **Try three examples** in the app.

三份文档均含概述、目标、计划三个对应章节，包含二级标题、一张嵌入图片与一张表格。导入“按章节融合”后点击“重新识别与配对”，也可直接使用应用里的示例按钮。

Expected output: three groups and nine cards; three navigation levels, three pictures, three tables when all cards are included. Figure and literal heading numbers are not automatically renumbered.

Generate with Node.js 22.13+ after `npm ci`:

```bash
node scripts/generate-examples.mjs
```

Only these exact three filenames are allowed through `.gitignore`. Other local documents under `examples/` remain private and ignored. Put personal experiments in `tmp/` or another ignored directory; never force-add them.
