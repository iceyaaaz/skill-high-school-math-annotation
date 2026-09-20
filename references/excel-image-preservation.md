# Excel 回写与题目图片保留

标注工作簿通常用 WPS/Excel 的单元格图片机制存放题目图与手写解答：`xl/media/*` 存图片字节，
`xl/cellimages.xml` 描述图片与单元格的关系，单元格里则是 `=DISPIMG("ID_xxx",1)` 公式。

绝大多数表格库（包括 openpyxl 重新保存）会在读写往返中丢掉这些资源，导致交付文件变成“题干全没了”的空壳。
因此采用**两阶段流程**：内容阶段任意工具，落盘阶段只用本 skill 的脚本。

## 1. 基线探查

```bash
python scripts/xlsx_preserve_tool.py inspect 原卷.xlsx
```

关注三项：

- `media_files`：`xl/media/` 下的图片数量。
- `cellimages_xml`：是否存在 `xl/cellimages.xml`。
- 每张表的 `dispimg_formulas`：`DISPIMG` 公式数量。

只要三者任一不为零，最终回写必须走第 4 步的 `write-back`。

## 2. 看图再批改

```bash
python scripts/xlsx_preserve_tool.py extract-media 原卷.xlsx --out ./media
```

抽出后逐张读取图片。批改结论必须来自实际看到的题干与解答；图片缺失或无法辨认时据实说明，不要凭空定罪。

## 3. 组装 spec.json

顶层用 `sheets` 按工作表名组织（也可以直接把表名作为顶层键）：

```json
{
  "sheets": {
    "题目1": {
      "cells": {
        "B7": "错误",
        "C7": "内分点坐标计算错误；相似对应边混淆",
        "D7": "内分比例代入时坐标分配有误，点的位置偏离题设比例，后续向量表示随之偏移。"
      },
      "clear": ["B8", "B9", "C8", "C9", "D8", "D9"],
      "add_merges": ["B7:B9", "C7:C9", "D7:D9"],
      "remove_merges": ["B7:B7", "C7:C7", "D7:D7"]
    }
  }
}
```

字段说明：

- `cells`：`地址 -> 值`。字符串写成内联字符串，数字直接写 JSON 数字（得分用数字）。
- `clear`：置空的单元格地址，用于清掉被合并区域里非左上角的隐藏值。
- `add_merges`：新增的合并区，与现有合并区取并集。
- `remove_merges`：要移除的合并区（通常是模板里原本的单块合并，需要扩成多块合并）。
- `merges`：仅在加 `--replace-merges` 时生效，整体替换该表的合并区。**慎用**——会丢掉标题、表头等无关合并区。

默认行为：合并区增删，不清空其他合并；加进来的合并区会自动清空其覆盖范围内非左上角单元格的值（可用 `--no-clear-covered` 关闭）。

## 4. 回写

```bash
python scripts/xlsx_preserve_tool.py write-back \
  --original 原卷.xlsx --spec spec.json --out 批改后.xlsx
```

脚本只做三件事：改目标单元格值、改 `<mergeCells>`、把其余 ZIP 部件字节级原样复制。
常用开关：

- `--force`：允许覆盖含 `DISPIMG` 的单元格（默认拒绝，保护题目图）。
- `--replace-merges`：用 `merges` 整体替换该表合并区。
- `--no-clear-covered`：保留被合并区域里非左上角单元格的旧值。

## 5. 复核

```bash
python scripts/xlsx_preserve_tool.py inspect 批改后.xlsx --json
```

- `media_files` 与 `cellimages_xml` 必须与基线一致。
- 用 `inspect --sheet <表名> --range B5:G30` 抽查每题首个、中间、末尾块的值与合并范围。

## 常见坑

- **不要用表格库直接保存最终文件。** 用它们读、算、校验都没问题，落盘交给 `write-back`。
- **不要动 `DISPIMG` 单元格。** 脚本默认会拒绝并提示，确属必要才 `--force`。
- **不要整体替换合并区。** 模板里标题行、表头通常是合并的，`--replace-merges` 会把它们一起抹掉。
- **`<mergeCells>` 有 schema 位置要求**，必须在 `sheetData` 之后、`pageMargins` 之前；脚本会自动放到合法位置。
- **字符串写成内联字符串**，不依赖 `sharedStrings.xml`，避免索引错位。
- **`DISPIMG` 预览显示 `#NAME?`** 是 Excel/预览工具未注册该 WPS 函数的兼容性现象，不是数据损坏，不要改写原公式。
- 图片放在 `xl/drawings/` 的浮动图片同样会被字节复制保留，无需额外处理。
