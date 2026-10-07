# 引用模式与验收

这是用户规定的固定核心工作流程。技能内只有两种引用工作模式：A（footnote-draft）和B（endnote-final）。生成排版与纯格式清理是任务类型，不是另两种引用模式。本文件适用于引用任务，不能由工具可用性推导出简化模式。

## 模式 A：footnote-draft / 脚注协作版

触发：模式A、A模式、脚注版、脚注协作版、多人合作版，以及footnote-draft。模式名称中的可选空格不影响识别。

1. 使用真正的Microsoft Word Footnote：正文有 `w:footnoteReference`，脚注部件中有对应的普通脚注记录，并有合法关系与内容类型。页底手工文本、文本框、手打上标或普通文末列表都不能替代。
2. 每次引用独立插入一个脚注；同一篇文献被引用三次，就保留三次独立脚注，不能共享一个脚注ID模拟独立引用。
3. 允许重复脚注，不做文献去重，不因为准备正式排版而自动合并文献。
4. 不创建NOTEREF交叉引用，不转尾注。若从已有B版生成A版，应把引用位置逐一恢复为独立真实脚注，而不是保留重复位置的NOTEREF。
5. 保留每次引用的来源信息和原有实质性评注，便于多人协作、修改、合并及核查；不虚构文献内容或标识。
6. 按实际引用位置数与普通脚注数核对；separator等特殊脚注不计入文献数量。排除悬空ID、错误关系和重复ID。
7. 最终交付名称与说明明确为“真正的 Word 脚注版 DOCX”。引用部分不能残留被当作已完成引用的NOTEREF或Endnote。

## 模式 B：endnote-final / 最终尾注去重交叉引用版

触发：模式B、B模式、尾注版、正式版、最终版、交叉引用版、尾注去重版，以及endnote-final。识别这些词在用户任务中的实际含义，不对引用案例、模式说明或待处理正文执行转换。

必须完成下列全部步骤；B等于尾注、去重、交叉引用和最终格式的完整结合。

1. **转换为真实尾注。** 将所有真正Footnote转成真正Endnote，维护正文 `w:endnoteReference`、尾注记录、关系及内容类型。已有Endnote也纳入盘点，避免ID碰撞。迁移文本、超链接、引用中包含的特殊字符及必要的关系；不能只把脚注文本复制到文末后删除锚点。说明性脚注也要按用户要求转换，但不能当作文献重复项擅自删除其评注。
2. **连续阿拉伯编号。** 用Word原生尾注编号，阿拉伯数字、连续编号，起始值为1；检查文档设置与各节覆盖，不能按节重启或沿用罗马数字。OOXML的note ID是关系键，不等于显示序号，必须分别管理。
3. **节尾位置。** 设置Endnotes位置为节的结尾（`sectEnd`/Word的End of section）。多节文档保持原有节结构，全局去重、连续编号，尾注位于首次真实引用所属节尾；不为了全部放到最后一页擅自合并节或改为文档尾。用户同时要求集中于整篇文档最后时，先明确其与节尾要求的关系。
4. **正文标记。** 正文显示上标`[1]`；方括号和数字均属于同一上标引用簇。数字来自真实尾注标记或可更新的交叉引用，不用固定文本冒充所有引用。
5. **尾注文本。** 每条尾注开头为baseline的`[1]`、真实Tab、文献正文，保留尾注内原生 `w:endnoteRef` 对应的编号机制。只对尾注区标记取消上标；正文标记仍为上标。不能批量把同名字符样式改成baseline而破坏正文。
6. **悬挂缩进。** 为尾注参考文献段落单独设置悬挂缩进和Tab位，使续行与文献正文对齐；不能套用正文首行2字符。未另定数值时可用0.74cm作为起始排版值，并随标签位数调整及视觉核对，这不是原师门文档规定的固定尺寸。
7. **分隔线。** 清理可见尾注separator及continuation separator横线，以及造成同类横线的段落边框；保留合法的separator/continuationSeparator特殊记录、合法ID及所需段落。区分特殊尾注记录与其中绘制横线的标记，不整块删除特殊记录或将其作为普通文献。Word重开后核实横线没有被恢复。
8. **执行安全去重。** 按下面的身份规则形成确认的文献等价组和待确认候选，不用题名模糊相似度代替文献身份判断。
9. **每篇唯一文献仅一个真实Endnote。** 按正文首次出现顺序选定组内首次引用及保留的尾注。只移除确证重复的文献尾注，不丢失独有评注、页码说明或原始文献信息；混有独有说明无法无损迁移时先列待处理项，不能直接丢弃。
10. **首次位置保留原生引用。** 每篇文献第一次出现的位置有真实 `w:endnoteReference`，关联唯一普通Endnote。其 `w:id` 是内部note ID，绝不能当作最终显示编号。把该原生引用包在唯一、合法、稳定的书签中；书签身份不得来自当前显示编号，禁止永久使用`ref_1`、`ref_2`、`ref_3`。使用 `_RefCite_A7F32C`、`_RefCite_000001`、基于已确认标识的稳定ID或Word生成的稳定名称。方括号与动态编号整体上标，但方括号不能把动态编号替换为静态字符串。
11. **后续重复位置建立动态交叉引用。** 每个后续引用必须使用Word原生NOTEREF复合域并指向首次真实尾注引用的稳定书签。完整结构必须包含 `fldChar begin`、可跨run拼接且含 `NOTEREF <stable-bookmark> \\h` 的`instrText`、`fldChar separate`、非空field result、`fldChar end`；`\\h`提供原生点击跳转。REF、普通internal hyperlink、只有缓存结果、简单静态`[1]`或把更新结果再次flatten成普通文字均不合格。大括号必须是Word域结构，不是键入的字符。
12. **合并同处多篇显示。** 同一处多篇显示为`[3,4]`，保留每篇文献的真实锚点或交叉引用关系；不能只保留一个链接代表整个组。
13. **压缩连续显示。** 确实连续的一组显示为`[6-9]`，非连续组如3、4、6不能伪装为`[3-6]`。显示压缩与引用关系分开管理；记录完整成员及首次/重复状态，不能为了短写而删除7、8的首次真实引用、尾注或交叉引用关系。实现必须能被Microsoft Word保存重开保留；若尚不能同时满足压缩显示与原生引用关系，B仍未完成，不以断链或丢引用换取外观正确。
14. **更新所有Word域。** 在Microsoft Word中调用原生域更新，覆盖正文、所有StoryRanges、尾注、页眉页脚以及Shapes/TextFrames等适用范围。不能以只执行正文Ctrl+A/F9、仅写 `updateFields=true`、标记dirty或填写缓存值声称全部更新完成。记录实际采用的更新方式和结果；不随意解除与本任务无关的保护或更新未授权外部数据连接。
15. **分别检查导航与动态域。** `navigation_ok`与`dynamic_field_ok`是两个独立结论。逐个解析NOTEREF，覆盖跨多个run拆开的复合域，确认begin、完整指令、稳定书签、separate、非空result和end全部存在；书签成对、唯一并包围/定位有效的首次 `w:endnoteReference`。原生引用有对应真实Endnote，NOTEREF数量等于已去重的后续重复引用数，`broken_noteref_count = 0`。搜索“Error! Reference source not found.”、“错误！未找到引用源。”等结果，但不能只凭没有错误文字断言关系正确。
16. **Word动态重编号与重开验收。** 先更新并保存正式输出，关闭后重开核对；再对测试副本在原第一条真实Endnote之前插入新的真实Endnote，更新全部域。若原来A=[1]、B=[2]且后文NOTEREF再次引用A，插入C后必须同步成为C=[1]、A=[2]、B=[3]，所有A的NOTEREF结果从[1]变为[2]；删除前置测试文献后也应重新编号。保存测试副本、关闭、重开后结果仍须一致。此项不能用LibreOffice、PDF、静态OOXML或可点击超链接替代；无法执行时设置`word_field_update_validation = not_tested`，模式B只能标记待Word验收。

### Mode B结构硬门槛

Mode B必须同时满足：`TRUE ENDNOTE + STABLE BOOKMARK + NOTEREF FIELD + CLICKABLE + F9 UPDATEABLE`。具体静态结构至少包括：

- `word/endnotes.xml`存在；每篇唯一文献有一个普通 `w:endnote`，内部用真实`w:endnoteRef`动态编号。
- `[Content_Types].xml`、`word/_rels/document.xml.rels`、styles、settings及合法separator结构完整。
- 正文首次引用是 `w:endnoteReference`，而非静态数字；书签包围该原生引用且名称与显示序号无关。
- 每个重复引用都是含begin/instrText/separate/result/end的NOTEREF复合域，指令包含稳定书签和`\\h`。
- 文末`[`、动态`w:endnoteRef`、`]`为baseline，随后是真正`w:tab`及文献正文；不能把`[1]`整体写为静态文字。

以下任一情况直接判定Mode B失败：静态`[1]`加hyperlink、静态`[1]`加bookmark、普通正文参考文献列表加bookmark、只有clickable link、显示了REF/NOTEREF结果却没有真实field code、更新后把域flatten成普通文字、缺少`word/endnotes.xml`。internal hyperlink可以用于其他合法文档导航，但不能承担Mode B的引用编号身份。

## B的去重安全规则

按以下优先顺序确认同一篇文献：

1. DOI完全一致。
2. PMID完全一致。
3. NCT编号完全一致。
4. ISBN或其他唯一标识完全一致。
5. 标准化后的完整参考文献字符串完全一致。

先从每条文献中提取有明确来源的标识，不根据题名去猜DOI/PMID等。去掉标识字段的外围空白或明确的URL/标签包装后比较完整值，保留原始值和处理记录；不把两种不同ID类型当成同一键，不用截断或局部包含判断相等。

完整字符串标准化只做可解释的排版处理，如Unicode规范化、首尾空白和重复空白、独立引用标签/Tab的移除；保留作者、题名、年份、卷期页码、版本、文献类型和标识。不能通过删除这些区分字段制造“完全一致”。

高优先级标识有明确冲突时，不能用相同的低优先级标识或相似字符串覆盖冲突。例如两篇不同DOI的论文共享一个NCT，不能仅因它们研究同一试验就合并成一篇论文；进入候选核对引用对象。同类标识提取出多个冲突值、不同版本/章节的身份不明确时同样暂不自动合并。确认相等的键与文献对象必须对应；不得以低等级匹配进行传递合并而绕过组内冲突。

只有题名相似、作者相似、年份相同或字符串高度相似的记录，列入 **duplicate candidates**，不得擅自合并。表中至少列出：候选条目ID及完整文献、首次引用位置、匹配证据、冲突或缺失的证据、需要用户确认的合并/保留选项。可以继续处理其他已确证重复组，但未确认候选不算已完成去重。保存用户确认结果，收到确认后继续完成模式B及Word验收，不重复询问已解决项。

维护可追溯映射：原引用位置/原脚注ID → 文献组 → 保留的尾注ID → 首次引用书签 → 后续NOTEREF目标 → 最终显示序号及引用簇成员。特殊separator记录和说明性注释与文献分开计数，不能把特殊节点当作重复文献。

## 与现有格式清理的关系

禁止列表编号对应 `w:pPr/w:numPr`，不包括原生脚注/尾注的reference、note内的ref、NOTEREF域、引用书签或 `w:endnotePr` 的连续编号设置。不能为满足“清除自动编号”而破坏这些原生引用。引用模式B明确授权的尾注转换、去重及引用格式变更优先于一般“保留引用形式”规则；纯清理任务没有这种授权时仍保留引用形式。

保持既有正文、标题层级、表格与文献信息，引用模式仅改变其要求的引用结构、重复条目及显示。三项段落分页属性仍按格式清理要求取消，不通过重新打开keepNext/keepLines/pageBreakBefore固定参考文献布局。

`word_format.py`提供普通Word生成与分页/列表格式清理。它**不实现脚注尾注转换、去重、NOTEREF建立、引用簇压缩或Microsoft Word更新域**；其审计成功不能作为A/B引用验收通过。涉及这些操作时使用当前环境可用的Word原生能力或支持相关OOXML部件的工具；既有复杂DOCX不得通过普通JSON生成器重建。没有Word环境可先完成已授权且能可靠验证的部分，但状态必须为待Word验收。

## 验收与交付状态

- **A完成**：本次引用位置均为独立真实脚注；重复文献未去重；没有用于替代本次引用的NOTEREF/尾注；脚注关系有效，交付真正Word脚注DOCX。
- **B待身份确认**：已处理可确认组，尚有需要用户决定的duplicate candidates；输出候选清单及明确待确认的工作副本，不称最终版。
- **B待Word验收**：静态结构已通过，但`word_field_update_validation = not_tested`，或未完成动态插入/删除重编号、全部域更新、保存关闭重开；报告缺失步骤，不把工作副本命名或描述为B完成版。
- **B完成**：所有可确认的重复文献已合并，所有候选已获确认并处理；首次位置为真实尾注引用，后续重复均为有效原生交叉引用；每篇唯一文献只保留一条真实Endnote；全部16项通过，包括实际Word重开。

交付摘要记录原脚注数、原有尾注数、引用位置数、唯一文献数、确认合并数、未决候选数、重复位置交叉引用数、真实尾注引用数、citation hyperlink数、broken NOTEREF数、`navigation_ok`、`dynamic_field_ok`、`word_field_update_validation`、动态插入/删除测试及Word保存重开结果。保留实测证据，不填写推测的“通过”。数量只是辅助证据，还须检查原生关系及动态显示。

代表性验收用例：同一文献引用三次时A有三个独立脚注，B有一个真实尾注和两个有效NOTEREF；同题名不同DOI不自动合并；同NCT但不同论文身份进入候选；跨节首次/重复引用连续编号且指向正确；引用簇3、4和6至9分别显示`[3,4]`、`[6-9]`且成员关系完整；删除/改名书签、移除separate、缺失endnotes.xml、把域转静态文字及静态超链接方案都必须被检出。动态测试在A、B之前插入C后，A的所有重复引用必须由1更新为2，保存重开仍正确。

## 引用审计工具

先做不启动Word的结构审计：

```text
python scripts/citation_audit.py input.docx --mode B --static-only --out static-audit.json
```

结构通过只能得到`pending_word_validation`。在Windows且已安装Microsoft Word时，用工作副本完成真正Word验收：

```text
powershell -ExecutionPolicy Bypass -File scripts/word_mode_b_validate.ps1 `
  -InputDocx input.docx -OutputDocx validated.docx -ReportJson word-validation.json
python scripts/citation_audit.py validated.docx --mode B `
  --word-validation-report word-validation.json --out final-audit.json
```

PowerShell脚本保存的是新的`validated.docx`，不覆盖输入；它更新各story域、保存关闭重开，并在临时副本前置插入真实Endnote检查原有尾注和NOTEREF同步加1。最终citation audit只有在静态结构、Word动态测试与重开均通过时才给出`status=passed`及`dynamic_field_ok=true`。若Word不可用，不伪造报告，保留`not_tested`。

## 原生结构参考

- [Microsoft Word Footnotes.Convert](https://learn.microsoft.com/en-us/office/vba/api/word.footnotes.convert)
- [Word Endnotes.NumberingRule](https://learn.microsoft.com/en-us/office/vba/api/word.endnotes.numberingrule)
- [Word节尾位置枚举](https://learn.microsoft.com/en-us/office/vba/api/word.wdendnotelocation)
- [Open XML EndnoteReference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.endnotereference)
- [Word域代码列表（含NoteRef）](https://support.microsoft.com/en-us/word/list-of-field-codes-in-word)
