# 师门 Word 与 PPT 格式

把师门排版规则做成可重复使用的技能，并附带Word生成、检查和正文分页修复脚本。包内包含Word技能和PPT技能；PPT技能仅落实原文的基础视觉规范。

源码仓库：[smr2098study-coder/-](https://github.com/smr2098study-coder/-)。

## 1.2.0 引用模式

引用流程补充在现有Word技能中，保留两种固定模式，PPT技能继续独立。

| 选择 | 输出要求 |
| --- | --- |
| 模式A / A模式 / 脚注版 / 脚注协作版 / 多人合作版 | 真正的Word脚注DOCX，每次引用独立脚注，允许重复，不去重、不建NOTEREF、不转尾注 |
| 模式B / B模式 / 尾注版 / 正式版 / 最终版 / 交叉引用版 / 尾注去重版 | 真正的节尾Endnote、连续阿拉伯编号、安全去重、首次原生尾注引用、后续NOTEREF、上标引用簇和悬挂参考文献；Microsoft Word更新全部域并保存关闭重开验收 |

模式B包含完整16步流程。根据DOI、PMID、NCT、ISBN/其他唯一标识、标准化完整字符串的优先次序核对身份；仅相似的条目输出duplicate candidates等待确认，不能猜测合并或声称已完成。正文显示如`[1]`、`[3,4]`、`[6-9]`，尾注区为baseline`[1] + Tab + 文献正文`。

完整定义及验收标准见[引用模式与验收](plugins/shimen-format/skills/shimen-word-format/references/citation-modes.md)。现有Python格式脚本不实现完整引用转换或Word域更新；本版本新增的是执行规则、验收条件及对原生引用结构的保护测试，不能把更新技能当成某份DOCX已完成模式B验证。没有Word保存重开证据时必须报告待验收。

## 能处理什么

Word正文宋体小四、英文数字Times New Roman、1.5倍行距、首行2字符；五级标题按师门字号、加粗和文字编号体系处理；表格居中、表内黑体五号与1.15倍行距；表题在上、图题在下。正文不用项目符号、隐藏编号或列表样式；标题和正文都取消三项分页属性，章节换页使用真实分页符。

## 1.1.0 新增的严格格式清理

本次把用户追加的十项规则合并进现有Word技能，以“生成排版”和“纯格式清理”两种任务类型处理，不再增加一个竞争同一任务的清理技能。它们与引用模式A/B分别判断。PPT技能继续独立。

这次明确替代1.0.0中保留标题、表图题注分页绑定的建议。现在段落、标题样式、默认值及继承链均清除`keepNext`、`keepLines`、`pageBreakBefore`和`numPr`，包含false/0残留；不靠关闭¶隐藏黑方块。仅清理格式时不改变原文、参考文献、标题层级、表格内容或其余字号字体。可见自动序号先准确转换为普通文字和无编号段落格式，避免丢掉章节或参考文献序号。

详细规则见[Word清理规范](plugins/shimen-format/skills/shimen-word-format/references/format-cleanup.md)。已经安装旧包时，请用更新的安装包更新目标插件，再在新聊天中验证。

这个包的规则来自参考文件的文字要求。原文没有规定的页边距、颜色、边框和总标题格式，单独标为补充默认值。检查脚本和逐页视觉验收共同使用，不能承诺所有字体、设备和复杂文档都绝不会出现格式差异。

## 文件结构

```text
.agents/plugins/marketplace.json
plugins/shimen-format/
  plugin.json
  .codex-plugin/plugin.json
  skills/
    shimen-word-format/
      SKILL.md
      agents/openai.yaml
      references/
      scripts/word_format.py
    shimen-ppt-format/SKILL.md
requirements.txt
tests/
dist/
  shimen-format-plugin.zip
  shimen-word-format-skill.zip
```

`plugin.json`使用Agent Plugins布局；`.codex-plugin/plugin.json`保留兼容布局。marketplace中路径相对仓库根目录。原始参考文件、研究报告、个人路径和报告内容不随仓库上传。

## ChatGPT桌面应用或网页版安装

GitHub是源码和分发来源，上传仓库不会自动安装进ChatGPT。安装入口取决于账户和工作区是否支持插件创建、导入或管理员管理。官方依据：[构建技能](https://learn.chatgpt.com/docs/build-skills)、[构建插件](https://learn.chatgpt.com/docs/build-plugins)、[插件打包](https://developers.openai.com/plugins/build/plugins)。

### 有工作区管理员权限：从GitHub导入

1. 确认管理员的GitHub账号可以访问本仓库。
2. 在ChatGPT的Admin > Plugins > Add > Import marketplace，Source填写`https://github.com/smr2098study-coder/-`，Path留空，Branch填写`main`。
3. 完成GitHub授权，确认导入结果；管理员将插件设为可用。
4. 从Plugins安装“师门 Word 与 PPT 格式”，新建聊天，用`@`选择插件。

这条GitHub导入路径是工作区管理员功能；不能假定所有个人账户都有该入口。参考：[GitHub导入与同步](https://learn.chatgpt.com/docs/enterprise/plugin-management)。

### 个人账户：在支持的界面创建自己的插件

如果你的ChatGPT有Plugin Creator，可以在Chat或Work中用`@`选择它，附上[插件安装包](dist/shimen-format-plugin.zip)（插件目录位于压缩包根部），发送下面的文字：

> 请根据附件创建或更新名为“师门 Word 与 PPT 格式”的个人技能型插件至1.2.0。使用包内Word和PPT的SKILL.md及其references，保留word_format.py辅助脚本。按固定模式A生成独立真实脚注协作版，按固定模式B完成真实尾注、安全去重、NOTEREF交叉引用和最终格式；B须在Word更新所有域并保存重开验收。格式清理不能破坏真实脚注、尾注和交叉引用，也不能改写原文。创建完成后说明此账户支持的安装方式。

文件包是否可以直接导入，以当前界面为准；如果它不接收压缩包，解压后上传SKILL.md、references中的文件和脚本，让Plugin Creator据此创建。不要仅将SKILL.md作为普通聊天附件就认为已经安装。创建、安装后在新聊天中验证；有脚本执行能力时才可运行辅助脚本。若没有插件创建或安装入口，GitHub仓库本身不能开启该权限，可先上传解压后的规则在当前会话使用，但这不等于持久安装。

若希望任何用户都能从公共Plugins目录搜索安装，还需要按官方流程提交插件：[插件提交与发布](https://developers.openai.com/plugins/deploy/submission)。GitHub仓库公开与公共目录发布是两步。

## 日常使用

安装后在新聊天选择插件，输入：

> 根据以下内容生成Word，按师门格式排版并完成交付检查。

纠正旧文档时：

> 只清理这份Word的段落分页属性和隐藏编号，连同标题样式和继承来源一并清理，保留原文、参考文献、标题层级、表格、图表和域。

引用任务可以直接说明：

> 模式A：将这份材料做成真正的Word脚注协作版，每次引用独立脚注，不去重。

> 模式B：将脚注协作版完成为尾注去重交叉引用正式版，按技能完整流程处理，疑似重复先列出让我确认，完成Word保存重开验收后再交付最终版。

默认支持自动匹配相关正式Word任务。若插件没有自动触发，显式用`@`选择；不要假定插件对所有新聊天无条件生效。

## 本地运行与验证

在隔离Python环境中安装`requirements.txt`；文档工具需要python-docx与lxml。回归测试额外需要`tests/requirements.txt`中的Pillow。

```text
python -m pip install -r requirements.txt -r tests/requirements.txt
python -m unittest discover -s tests -v
python plugins/shimen-format/skills/shimen-word-format/scripts/word_format.py build input.json --out draft.docx
python plugins/shimen-format/skills/shimen-word-format/scripts/word_format.py audit draft.docx --out audit.json
python plugins/shimen-format/skills/shimen-word-format/scripts/word_format.py clean-format original.docx --out cleaned.docx
python plugins/shimen-format/skills/shimen-word-format/scripts/word_format.py audit cleaned.docx --format-only --out cleanup-audit.json
```

输入JSON与修复边界见Word技能的`references/tools.md`。检查失败时退出码为1；自动角色识别需人工确认。完整审计核对师门排版，`--format-only`只验收指定清理，不顺便改旧稿的其他格式。回归测试覆盖继承源、扩展样式部件、文字和参考文献保留、真实分页符、图表、域、标题层级及重复清理；由于验证环境缺少LibreOffice，尚未完成Word中开启¶后的界面验收和逐页视觉验收，也尚未在ChatGPT新聊天中验证安装与调用。

## 发布到GitHub

本次使用用户创建的公开仓库`smr2098study-coder/-`。在一个新的目录中克隆后继续维护，已有该仓库的克隆时跳过克隆步骤。不要上传单独提供的0930诊断文件或源报告。

```text
git clone https://github.com/smr2098study-coder/-.git shimen-format-skills
cd shimen-format-skills
git add .
git commit -m "Update Shimen document formatting skills"
git push
```

本包未指定开源许可证；可以按自己的分享意图另行选择许可证。GitHub源码可用不代表插件已被导入、安装或收录到公共Plugins目录；安装后仍需在目标ChatGPT账户的新聊天中验证。
