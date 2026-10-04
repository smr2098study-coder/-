# 工具用法

在技能目录下运行；`python`替换为环境提供的Python路径。依赖为 `python-docx>=1.1,<2`、`lxml>=4.9,<7`。JSON及检查结果使用UTF-8。

```text
python scripts/word_format.py build input.json --out draft.docx
python scripts/word_format.py audit draft.docx --out audit.json
python scripts/word_format.py repair-body-flags original.docx --out cleaned.docx --paragraphs roles.json
```

`audit` 有硬性错误时退出码为1，否则为0；警告要求人工判断。段落索引从0开始，按主文档XML中所有段落顺序（包含表格）计数。`roles.json` 为已确认普通正文的索引数组，如 `[8,9,12]`。不提供索引时，修复脚本仅选择非空、非表格、非目录、非标题、非题注且没有对象和域的普通正文候选；自动候选不是语义判断，先读审计。输入输出禁止使用同一路径，不覆盖已有输出。

## 新建JSON

```json
{
  "title": "研究报告",
  "blocks": [
    {"type": "heading", "level": 1, "text": "一、研究背景"},
    {"type": "paragraph", "text": "这里是正文，包含 English 和 2026。"},
    {"type": "heading", "level": 2, "text": "（一）研究目的"},
    {"type": "table", "caption": "表1 数据汇总", "rows": [["指标", "结果"], ["样本数", "20"]]},
    {"type": "image", "path": "figure.png", "caption": "图1 研究流程", "width_cm": 12}
  ]
}
```

块类型仅支持heading、paragraph、table、image。标题文字应包含要求的层级编号；脚本不推断和自动增号。一级“一、”、二级“（一）”、三级“1、”、四级“（1）”、五级“①”。内嵌图片路径相对JSON位置，绝不联网取图。列数必须一致，图片不超过可用页面宽度；图片保持比例。若表格含长文本，调整列宽并渲染检查，脚本不解决所有复杂表格布局。

新建工具不支持富文本、公式、脚注、目录、引文域、批注和复杂封面；有这些需求时用其他DOCX编辑方式保留或创建，随后审计。新建工具不能用于已有复杂报告的无损重建。

审计解析主文档段落、段落样式/字符样式继承、文档默认值、主题字体、分页标志、编号及基本表格对齐。它检查可识别的角色，对未知角色给出警告；不验证页码、目录内容一致性、图片题注全部配对、标题连续编号、字体在本机是否存在、视觉布局和脚注等全部部件。自动角色识别需人工复核。
