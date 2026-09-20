#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx_preserve_tool.py — 只读/最小侵入地操作 .xlsx，全程保留题目图片（DISPIMG/cellimages.xml/xl/media）。

用途：标注类任务中，绝大多数 Excel 工具（含 openpyxl 重新保存）会丢掉工作簿内嵌的
题目图片。本工具直接操作 XLSX 的 ZIP/XML，只改单元格值和 <mergeCells>，其余部件
（xl/media、cellimages.xml、drawings、rels、样式、数据验证）原样字节复制。

子命令：
  inspect       查看结构：工作表、合并区、DISPIMG 公式数、媒体文件数，可选导出单元格值
  extract-media 抽出 xl/media 下的图片到目录，便于逐张查看题目图
  write-back    按 JSON 规格写入单元格值 + 合并区，保留其他所有部件

仅依赖 Python 标准库。

示例：
  python xlsx_preserve_tool.py inspect 原卷.xlsx
  python xlsx_preserve_tool.py inspect 原卷.xlsx --sheet "题目1" --range B5:G20
  python xlsx_preserve_tool.py extract-media 原卷.xlsx --out ./media
  python xlsx_preserve_tool.py write-back --original 原卷.xlsx --spec spec.json --out 批改后.xlsx
"""

from __future__ import annotations

import argparse
import json
import os
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XMLNS = "http://www.w3.org/XML/1998/namespace"

for _p, _u in (("", MAIN), ("r", RELS), ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")):
    ET.register_namespace(_p, _u)

REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")
RANGE_RE = re.compile(r"^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$")

# CT_Worksheet 中必须排在 <mergeCells> 之前的元素
BEFORE_MERGE = {
    "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
    "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews",
}


def _tag(el) -> str:
    t = el.tag
    if not isinstance(t, str):
        return ""
    return t.split("}")[-1]


def col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_ref(ref: str):
    m = REF_RE.match(ref.strip())
    if not m:
        raise ValueError(f"非法单元格地址: {ref}")
    return int(m.group(2)), col_to_num(m.group(1))


def iter_range(a1: str):
    m = RANGE_RE.match(a1.strip())
    if not m:
        raise ValueError(f"非法区域地址: {a1}")
    c1, r1, c2, r2 = col_to_num(m.group(1)), int(m.group(2)), col_to_num(m.group(3)), int(m.group(4))
    for r in range(min(r1, r2), max(r1, r2) + 1):
        for c in range(min(c1, c2), max(c1, c2) + 1):
            yield f"{num_to_col(c)}{r}"


def top_left(a1: str) -> str:
    return next(iter_range(a1))


# --------------------------------------------------------------------------
# ZIP 读写：其余部件字节级原样保留
# --------------------------------------------------------------------------

def read_zip(path: str):
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}
    return infos, data


def write_zip(path: str, infos, data: dict):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            ni = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            ni.compress_type = info.compress_type
            ni.external_attr = info.external_attr
            ni.create_system = info.create_system
            z.writestr(ni, data[info.filename])


def sheet_index(data: dict):
    """返回 [(sheet_name, zip_path), ...]"""
    wb = ET.fromstring(data["xl/workbook.xml"])
    rels = ET.fromstring(data["xl/_rels/workbook.xml.rels"])
    rid2target = {rel.get("Id"): rel.get("Target") for rel in rels}
    out = []
    for sh in wb.find(f"{{{MAIN}}}sheets"):
        rid = sh.get(f"{{{RELS}}}id")
        tgt = rid2target.get(rid)
        if not tgt:
            continue
        if tgt.startswith("/"):
            p = tgt.lstrip("/")
        else:
            p = posixpath.normpath(posixpath.join("xl", tgt)).replace("\\", "/")
        out.append((sh.get("name"), p))
    return out


def shared_strings(data: dict):
    if "xl/sharedStrings.xml" not in data:
        return []
    root = ET.fromstring(data["xl/sharedStrings.xml"])
    out = []
    for si in root:
        out.append("".join(t.text or "" for t in si.iter(f"{{{MAIN}}}t")))
    return out


# --------------------------------------------------------------------------
# 单元格操作
# --------------------------------------------------------------------------

def ensure_row(sheet_data, r: int):
    rows = [el for el in sheet_data if _tag(el) == "row"]
    for el in rows:
        if int(el.get("r", 0)) == r:
            return el
    el = ET.SubElement(sheet_data, f"{{{MAIN}}}row")
    el.set("r", str(r))
    ordered = sorted(list(sheet_data), key=lambda e: int(e.get("r", 0)) if _tag(e) == "row" else -1)
    for c in list(sheet_data):
        sheet_data.remove(c)
    for c in ordered:
        sheet_data.append(c)
    return el


def ensure_cell(row, ref: str):
    for c in row:
        if c.get("r") == ref:
            return c
    c = ET.SubElement(row, f"{{{MAIN}}}c")
    c.set("r", ref)
    ordered = sorted(list(row), key=lambda e: col_to_num(re.sub(r"\d", "", e.get("r", "A1"))))
    for e in list(row):
        row.remove(e)
    for e in ordered:
        row.append(e)
    return c


def find_cell(sheet_data, ref: str):
    r, _ = parse_ref(ref)
    for row in sheet_data:
        if int(row.get("r", 0)) == r:
            for c in row:
                if c.get("r") == ref:
                    return c
    return None


def has_dispimg(cell) -> bool:
    for f in cell.findall(f"{{{MAIN}}}f"):
        if "DISPIMG" in (f.text or "").upper():
            return True
    return False


def set_value(cell, value):
    for child in list(cell):
        if _tag(child) in ("v", "is", "f"):
            cell.remove(child)
    if value is None or value == "":
        cell.attrib.pop("t", None)
        return
    if isinstance(value, bool):
        cell.set("t", "b")
        ET.SubElement(cell, f"{{{MAIN}}}v").text = "1" if value else "0"
    elif isinstance(value, (int, float)):
        cell.attrib.pop("t", None)
        ET.SubElement(cell, f"{{{MAIN}}}v").text = str(value)
    else:
        cell.set("t", "inlineStr")
        is_el = ET.SubElement(cell, f"{{{MAIN}}}is")
        t_el = ET.SubElement(is_el, f"{{{MAIN}}}t")
        t_el.set(f"{{{XMLNS}}}space", "preserve")
        t_el.text = str(value)


def cell_text(cell, shared):
    if cell is None:
        return ""
    t = cell.get("t")
    if t == "inlineStr":
        return "".join(x.text or "" for x in cell.iter(f"{{{MAIN}}}t"))
    v = cell.find(f"{{{MAIN}}}v")
    if v is None:
        f = cell.find(f"{{{MAIN}}}f")
        return ("=" + (f.text or "")) if f is not None else ""
    if t == "s":
        try:
            return shared[int(v.text)]
        except (ValueError, IndexError, TypeError):
            return v.text or ""
    return v.text or ""


# --------------------------------------------------------------------------
# 合并区
# --------------------------------------------------------------------------

def merge_refs(sheet) -> list:
    mc = sheet.find(f"{{{MAIN}}}mergeCells")
    if mc is None:
        return []
    return [m.get("ref") for m in mc.findall(f"{{{MAIN}}}mergeCell")]


def set_merges(sheet, refs: list):
    refs = sorted(set(refs), key=lambda a: (parse_ref(top_left(a))[0], parse_ref(top_left(a))[1]))
    old = sheet.find(f"{{{MAIN}}}mergeCells")
    if old is not None:
        idx = list(sheet).index(old)
        sheet.remove(old)
    else:
        idx = None
    mc = ET.Element(f"{{{MAIN}}}mergeCells")
    mc.set("count", str(len(refs)))
    for a in refs:
        ET.SubElement(mc, f"{{{MAIN}}}mergeCell").set("ref", a)
    if idx is not None:
        sheet.insert(idx, mc)
    else:
        pos = len(list(sheet))
        children = list(sheet)
        for i, el in enumerate(children):
            if _tag(el) in BEFORE_MERGE:
                pos = i + 1
        sheet.insert(min(pos, len(children)), mc)


# --------------------------------------------------------------------------
# 子命令
# --------------------------------------------------------------------------

def cmd_inspect(args):
    infos, data = read_zip(args.file)
    media = [i.filename for i in infos if i.filename.startswith("xl/media/")]
    drawings = [i.filename for i in infos if i.filename.startswith("xl/drawings/")]
    report = {
        "file": args.file,
        "zip_entries": len(infos),
        "cellimages_xml": "xl/cellimages.xml" in data,
        "media_files": len(media),
        "drawing_files": len(drawings),
        "sheets": [],
    }
    shared = shared_strings(data)
    for name, path in sheet_index(data):
        if path not in data:
            report["sheets"].append({"name": name, "path": path, "error": "missing"})
            continue
        root = ET.fromstring(data[path])
        sd = root.find(f"{{{MAIN}}}sheetData")
        formulas = [f for f in root.iter(f"{{{MAIN}}}f") if "DISPIMG" in (f.text or "").upper()]
        info = {
            "name": name,
            "path": path,
            "rows": len(sd) if sd is not None else 0,
            "merges": len(merge_refs(root)),
            "dispimg_formulas": len(formulas),
        }
        if args.sheet in (None, name) and args.range:
            info["values"] = {
                ref: cell_text(find_cell(sd, ref), shared) for ref in iter_range(args.range)
            }
        report["sheets"].append(info)

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return
    print(f"文件: {report['file']}")
    print(f"ZIP 条目: {report['zip_entries']}  媒体图片: {report['media_files']}  "
          f"drawings: {report['drawing_files']}  cellimages.xml: {report['cellimages_xml']}")
    for s in report["sheets"]:
        print(f"  - 工作表 {s['name']!r}  行:{s.get('rows')}  合并区:{s.get('merges')}  "
              f"DISPIMG公式:{s.get('dispimg_formulas')}  ({s.get('path')})")
        if args.range and "values" in s:
            refs = list(iter_range(args.range))
            rows = sorted({parse_ref(r)[0] for r in refs})
            for r in rows:
                line = [f"{c}={s['values'].get(f'{c}{r}', '')!r}" for c in
                        sorted({re.sub(r'\\d', '', x) for x in refs}, key=col_to_num)]
                print("      " + " | ".join(line))


def cmd_extract_media(args):
    infos, data = read_zip(args.file)
    out = args.out or os.path.splitext(os.path.basename(args.file))[0] + "_media"
    os.makedirs(out, exist_ok=True)
    n = 0
    for name in data:
        if name.startswith("xl/media/"):
            p = os.path.join(out, os.path.basename(name))
            with open(p, "wb") as fh:
                fh.write(data[name])
            print(p)
            n += 1
    print(f"共导出 {n} 个媒体文件 -> {os.path.abspath(out)}")
    if n == 0:
        print("提示: 该工作簿无 xl/media，题目图片可能存放在 drawings 或为浮动图片。")


def cmd_write_back(args):
    with open(args.spec, encoding="utf-8") as fh:
        spec = json.load(fh)
    sheets_spec = spec.get("sheets", spec)

    infos, data = read_zip(args.original)
    index = dict(sheet_index(data))
    shared = shared_strings(data)

    stats = {"cells": 0, "cleared": 0, "merges_added": 0, "merges_removed": 0, "skipped_dispimg": 0}

    for sheet_name, cfg in sheets_spec.items():
        if sheet_name not in index:
            raise SystemExit(f"找不到工作表: {sheet_name}（现有: {list(index)}）")
        path = index[sheet_name]
        root = ET.fromstring(data[path])
        sd = root.find(f"{{{MAIN}}}sheetData")
        if sd is None:
            raise SystemExit(f"{sheet_name} 缺少 sheetData")

        for ref, val in (cfg.get("cells") or {}).items():
            cell = find_cell(sd, ref)
            if cell is None:
                r, c = parse_ref(ref)
                cell = ensure_cell(ensure_row(sd, r), ref)
            if has_dispimg(cell) and not args.force:
                stats["skipped_dispimg"] += 1
                print(f"[跳过] {sheet_name}!{ref} 含 DISPIMG 图片公式，拒绝覆盖（--force 可强改）")
                continue
            set_value(cell, val)
            stats["cells"] += 1

        for ref in (cfg.get("clear") or []):
            cell = find_cell(sd, ref)
            if cell is not None and not (has_dispimg(cell) and not args.force):
                set_value(cell, None)
                stats["cleared"] += 1

        existing = merge_refs(root)
        if args.replace_merges and "merges" in cfg:
            final = list(cfg["merges"])
        else:
            remove = set(cfg.get("remove_merges") or [])
            final = [m for m in existing if m not in remove]
            for m in cfg.get("add_merges") or []:
                if m not in final:
                    final.append(m)
                    stats["merges_added"] += 1
            stats["merges_removed"] += len([m for m in existing if m in remove])

        if args.clear_covered:
            for rng in (cfg.get("add_merges") or cfg.get("merges") or []):
                tl = top_left(rng)
                for ref in iter_range(rng):
                    if ref == tl:
                        continue
                    cell = find_cell(sd, ref)
                    if cell is None or has_dispimg(cell):
                        continue
                    if cell_text(cell, shared) != "":
                        set_value(cell, None)
                        stats["cleared"] += 1

        set_merges(root, final)
        data[path] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + \
                     ET.tostring(root, encoding="utf-8")

    write_zip(args.out, infos, data)
    media_before = len([i.filename for i in infos if i.filename.startswith("xl/media/")])
    print(json.dumps({"out": args.out, **stats, "media_preserved": media_before},
                     ensure_ascii=False, indent=2))


def main():
    ap = argparse.ArgumentParser(description="保留图片的 XLSX 只读检查与最小侵入回写工具")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("inspect", help="查看工作簿结构")
    p1.add_argument("file")
    p1.add_argument("--sheet")
    p1.add_argument("--range", dest="range")
    p1.add_argument("--json", action="store_true")
    p1.set_defaults(func=cmd_inspect)

    p2 = sub.add_parser("extract-media", help="导出内嵌图片")
    p2.add_argument("file")
    p2.add_argument("--out")
    p2.set_defaults(func=cmd_extract_media)

    p3 = sub.add_parser("write-back", help="写回单元格与合并区，保留媒体")
    p3.add_argument("--original", required=True)
    p3.add_argument("--spec", required=True)
    p3.add_argument("--out", required=True)
    p3.add_argument("--force", action="store_true", help="允许覆盖含 DISPIMG 的单元格")
    p3.add_argument("--replace-merges", action="store_true", help="用 spec 的 merges 整体替换该表合并区")
    p3.add_argument("--no-clear-covered", dest="clear_covered", action="store_false",
                    help="不清空被合并区域覆盖的非左上角单元格")
    p3.set_defaults(func=cmd_write_back, clear_covered=True)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()
