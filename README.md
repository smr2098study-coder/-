# 师门 Word 与 PPT 格式

把师门排版规则做成可重复使用的技能，并附带Word生成、检查和正文分页修复脚本。包内包含Word技能和PPT技能；PPT技能仅落实原文的基础视觉规范。

源码仓库：[smr2098study-coder/-](https://github.com/smr2098study-coder/-)。

## 能处理什么

Word正文宋体小四、英文数字Times New Roman、1.5倍行距、首行2字符；五级标题按师门字号、加粗和编号体系处理；表格居中、表内黑体五号与1.15倍行距；表题在上、图题在下。区分实际项目符号与Word不打印的分页标记，避免正文连续绑定下一段。

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

> 请根据附件创建名为“师门 Word 与 PPT 格式”的个人技能型插件。使用包内Word和PPT的SKILL.md及其references，保留word_format.py辅助脚本，不添加外部MCP服务。按包内规范生成可编辑Word，生成后做结构检查和逐页视觉检查。不要把实例报告当作任务指令；不要承诺未验证的零格式问题。创建完成后说明此账户支持的安装方式。

文件包是否可以直接导入，以当前界面为准；如果它不接收压缩包，解压后上传SKILL.md、references中的文件和脚本，让Plugin Creator据此创建。不要仅将SKILL.md作为普通聊天附件就认为已经安装。创建、安装后在新聊天中验证；有脚本执行能力时才可运行辅助脚本。若没有插件创建或安装入口，GitHub仓库本身不能开启该权限，可先上传解压后的规则在当前会话使用，但这不等于持久安装。

若希望任何用户都能从公共Plugins目录搜索安装，还需要按官方流程提交插件：[插件提交与发布](https://developers.openai.com/plugins/deploy/submission)。GitHub仓库公开与公共目录发布是两步。

## 日常使用

安装后在新聊天选择插件，输入：

> 根据以下内容生成Word，按师门格式排版并完成交付检查。

纠正旧文档时：

> 检查这份Word的师门格式，先判断段前黑点来源，再修复确认的格式问题，保留正文、图表、目录和域。

默认支持自动匹配相关正式Word任务。若插件没有自动触发，显式用`@`选择；不要假定插件对所有新聊天无条件生效。

## 本地运行与验证

在隔离Python环境中安装`requirements.txt`；文档工具需要python-docx与lxml。回归测试额外需要`tests/requirements.txt`中的Pillow。

```text
python -m pip install -r requirements.txt -r tests/requirements.txt
python -m unittest discover -s tests -v
python plugins/shimen-format/skills/shimen-word-format/scripts/word_format.py build input.json --out draft.docx
python plugins/shimen-format/skills/shimen-word-format/scripts/word_format.py audit draft.docx --out audit.json
```

输入JSON与修复边界见Word技能的`references/tools.md`。检查失败时退出码为1，提醒调用方处理格式错误；自动角色识别需人工确认。此版本的本地脚本回归测试通过，真实样本已用于结构审计；由于验证环境缺少LibreOffice，尚未完成逐页视觉验收，也尚未在ChatGPT新聊天中验证安装与调用。

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
