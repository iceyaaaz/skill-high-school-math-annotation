# high-school-math-annotation

按任务要求和标注示例批改高中数学解答工作簿，填写分步批改结果、块错因、块基础错因、得分和思路总结，
合并连续同类批改块，并回写成**保留原题图片**的 Excel。

## 目录结构

```text
high-school-math-annotation/
├── SKILL.md                              入口：工作流程与规则
├── references/
│   ├── annotation-rules.md               状态判定、字段口径、措辞句式
│   └── excel-image-preservation.md       两阶段回写与 DISPIMG 图片保留
├── scripts/
│   └── xlsx_preserve_tool.py             仅用标准库的 XLSX 检查/回写工具
└── agents/openai.yaml                    Codex 端调用界面（WorkBuddy 不读取）
```

## 在 WorkBuddy 中使用

安装在 `~/.workbuddy/skills/high-school-math-annotation/`，对话里直接描述任务即可触发，例如
“按任务要求和标注示例批改这份高中数学解答工作簿，保留题目图片并导出 Excel”。

## 在 Codex / 其他 agent 中使用

把 `SKILL.md` 与 `references/` 提供给 agent，或按 `agents/openai.yaml` 的界面显式调用
`$high-school-math-annotation`。若只希望显式调用，把 `allow_implicit_invocation` 改为 `false`。

`scripts/xlsx_preserve_tool.py` 只依赖 Python 标准库，任何环境都能直接跑：

```bash
python xlsx_preserve_tool.py inspect 原卷.xlsx
python xlsx_preserve_tool.py extract-media 原卷.xlsx --out ./media
python xlsx_preserve_tool.py write-back --original 原卷.xlsx --spec spec.json --out 批改后.xlsx
```

## 为什么需要一个专门的回写脚本

题目图片往往以 `DISPIMG` 公式 + `xl/cellimages.xml` + `xl/media/*` 的形式存在，
常规表格库读写往返会把它们全部丢掉。本 skill 的回写脚本直接改 XLSX 的 ZIP/XML，
只动单元格值和 `<mergeCells>`，其余部件字节级原样保留，并默认拒绝覆盖图片公式单元格。
