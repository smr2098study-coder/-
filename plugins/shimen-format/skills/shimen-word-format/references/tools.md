# 工具用法

在技能目录下运行；`python`替换为环境提供的Python路径。依赖为 `python-docx>=1.1,<2`、`lxml>=4.9,<7`。JSON及检查结果使用UTF-8。

```text
python scripts/word_format.py build input.json --out draft.docx
python scripts/word_format.py audit draft.docx --out audit.json
python scripts/word_format.py clean-format original.docx --out cleaned.docx
python scripts/word_format.py audit cleaned.docx --format-only --out cleanup-audit.json
```

`audit` 有硬性错误时退出码为1，否则为0；警告要求人工判断。段落索引从0开始，按主文档XML中所有段落顺序（包含表格）计数。`clean-format`取代旧版 `repair-body-flags`，同时清理标题、正文、样式和默认值，清除三项分页属性及 `numPr`，转换列表样式，保留文字和其他格式。输入输出禁止使用同一路径，不覆盖已有输出。原稿存在未物化的数字自动序号时，先转成准确普通文字；脚本不猜序号。

## 新建JSON

```json
{
  "title": "研究报告",
  "blocks": [
    {"type": "heading", "level": 1, "text": "一、研究背景"},
    {"type": "paragraph", "text": "这里是正文，包含 English 和 2026。"},
    {"type": "heading", "level": 2, "text": "（一）研究目的"},
    {"type": "page_break"},
    {"type": "table", "caption": "表1 数据汇总", "rows": [["指标", "结果"], ["样本数", "20"]]},
    {"type": "image", "path": "figure.png", "caption": "图1 研究流程", "width_cm": 12}
  ]
}
```

块类型支持heading、paragraph、table、image和page_break。page_break插入真实分页符，不设置 `pageBreakBefore`。标题文字应包含要求的层级编号；脚本不推断和自动增号。一级“一、”、二级“（一）”、三级“1、”、四级“（1）”、五级“①”。内嵌图片路径相对JSON位置，绝不联网取图。列数必须一致，图片不超过可用页面宽度；图片保持比例。若表格含长文本，调整列宽并渲染检查，脚本不解决所有复杂表格布局。

新建工具不支持富文本、公式、脚注、目录、引文域、批注和复杂封面；有这些需求时用其他DOCX编辑方式保留或创建，随后审计。新建工具不能用于已有复杂报告的无损重建。

引用模式A/B见[引用模式与验收](citation-modes.md)。本脚本不是引用转换器：不执行脚注转尾注、文献去重、NOTEREF构建、引用簇压缩或Word域更新。其 `--format-only` 只检查段落分页与列表属性；真实脚注/尾注节点及NOTEREF应保留，不能据此宣告引用模式验收通过。模式B必须另在Microsoft Word更新所有域并保存、关闭、重新打开验证。

分页及编号检查扫描全部Word XML部件中的段落、样式与默认值； `--format-only`只检查用户追加的清理规则，不要求顺便改写原有字体、字号及标题体系。完整审计另解析主文档角色、字体、字号、缩进、行距和基本表格对齐。两种检查都不验证页码、目录内容一致性、题注全部配对、标题连续编号、字体是否安装或视觉布局。清理前后内容保真需另做对比。
