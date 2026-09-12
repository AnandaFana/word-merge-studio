# 按章节融合 / Section studio

将多份结构相近的 `.docx` 按对应章节组织成一篇文档。适合项目报告、部门计划、课程说明、活动方案等需要“同一章节集中阅读”的材料。

## 操作流程

1. 主页 → **按章节融合**，添加文档或点击 **体验三个示例**。
2. 选择拆分层级（1–3 级标题）。优先使用 Word 原有大纲/标题样式；没有样式时，使用本地编号、字体与加粗规则推断。
3. 点击 **重新识别与配对**。默认按去除编号、标点和空白后的相同标题配对；也可明确选择按章节位置配对。不会猜测语义相近但措辞不同的标题。
4. 修改分组标题与来源小标题；用箭头调整组顺序和组内卡片顺序。通过“移动到”把卡片归入另一个组；可新建组、取消卡片或撤销最近 30 次布局操作。
5. 选择保留来源正文格式，或按正文、标题、题注和表格文字分别统一字体与字号。
6. 预览、导出 `.docx`，在 Word 的“视图 → 导航窗格”检查三级结构。可另存 JSON 布局记录，便于追踪来源与排除项。

增加文件时，会追加独立分组以保护已有手动布局。需要自动匹配时再次点击“重新识别与配对”；此操作会重建布局，但可撤销。拆分选项只在重新识别时应用。示例按钮会替换当前工作区，清空按钮会清除当前来源与布局。

## 输出结构

```text
一级：目标（分组标题，可编辑）
  二级：花园项目（来源小标题，可编辑）
    三级：原有子标题
  二级：书屋项目
    三级：原有子标题
一级：计划
  …
```

原章节根标题由分组标题替代。每张卡片中最高层的剩余标题设为第三级，更深标题降为正文。无标题的前置内容单独保留，不默默丢弃；表格是完整内容块，表格内部标题不会用于拆分。

保留格式指来源正文、表格结构与图片的格式，而不是保留每份文件独立的页面框架。页面尺寸、页边距、页眉页脚和主题采用**来源清单中第一个仍有保留卡片的文档**，与分组重排无关；原有分节移除。新增分组/来源标题采用输出标题设置。统一模式中未勾选的来源内容类型保留外观，导航级别仍按新结构设置。

## 当前边界

- 不自动润色、去重或改写正文；不会重编文字中的章节号、图号和表号。跨引用位置变动后需人工核对。
- 图片复制原始字节并重建包关系；浮动图形、图表主题、特殊表格样式和复杂版面仍需 Word 人工复核。
- 含脚注/尾注、批注、修订、正文域、内容控件、嵌入对象或旧式 VML 的来源会明确拒绝导入。先在副本中处理相关结构，保留原件。
- 最多 20 份文件，单份 25 MB，总压缩大小 100 MB、总解压大小 250 MB；单来源最多 2,500 个正文块、总段落数不超过 10,000。
- 浏览器中的卡片状态只存于当前页面；刷新会清空。语言偏好保存在本机。JSON 是审计记录，包含文件名、标题及块范围，**不能重新载入项目**。

## English quick guide

Open **Merge sections**, load the fictional examples or your `.docx` files, select the heading level, and click **Detect and regroup**. Match exact normalised titles or explicitly choose positional matching. Rename groups and source headings, reorder cards with arrows, move them between groups, and exclude unwanted cards.

Export retains source text, tables and embedded images. Choose source body formatting or type-based formatting. Output navigation has three levels: group, source, original subheading. Deeper headings become body text. Original section roots are replaced by group titles. Front matter is retained separately. Generated titles use the export heading settings.

Page setup, headers and footers follow the first included source; source section breaks are removed. Caption numbers and literal heading numbers are not rewritten. This is deterministic local assembly, not AI rewriting or semantic deduplication. Review the final `.docx` in Word, especially pagination and floating images. The JSON audit cannot restore the working session.
