#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inspect_metadata.py — 查看 docx 的：
  1. 核心属性 / 扩展属性 / 自定义属性（元数据）
  2. 内部结构硬证据：段落数、run 数、run/段比、pStyle 使用、固定模板时间戳、
     批注、修订、trackChanges
零第三方依赖，Python 3.7+。

用法：
    python3 inspect_metadata.py 文档.docx
"""
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

for _p, _uri in NS.items():
    ET.register_namespace(_p, _uri)


def qn(prefix, tag):
    return "{%s}%s" % (NS[prefix], tag)


def pretty_tag(tag):
    if tag.startswith("{"):
        uri, local = tag[1:].split("}", 1)
        for p, u in NS.items():
            if u == uri:
                return "%s:%s" % (p, local)
    return tag


def dump_part(title, data):
    print("=" * 70)
    print(title)
    print("=" * 70)
    if data is None:
        print("（不存在）\n")
        return
    root = ET.fromstring(data)
    for elem in root.iter():
        if elem is root:
            continue
        text = (elem.text or "").strip()
        if elem.tag.endswith("}property"):
            # custom.xml 的属性节点，值在子节点
            child_texts = [c.text for c in elem if c.text]
            print("  [property] %-20s %s" % (elem.get("name"), " | ".join(child_texts)))
            continue
        if text or elem.attrib:
            xtype = elem.get("{http://www.w3.org/2001/XMLSchema-instance}type")
            extra = "  (xsi:type=%s)" % xtype if xtype else ""
            print("  %-28s %s%s" % (pretty_tag(elem.tag), text or "(空)", extra))
    print()


def analyze_structure(entries):
    print("=" * 70)
    print("内部结构分析 word/（硬证据）")
    print("=" * 70)
    names = set(entries)

    if "word/document.xml" not in names:
        print("（缺少 word/document.xml，不是正常 docx）\n")
        return

    doc_root = ET.fromstring(entries["word/document.xml"])
    paragraphs = list(doc_root.iter(qn("w", "p")))
    runs = list(doc_root.iter(qn("w", "r")))
    pstyles = list(doc_root.iter(qn("w", "pStyle")))
    ins = list(doc_root.iter(qn("w", "ins")))
    dels = list(doc_root.iter(qn("w", "del")))

    n_p, n_r = len(paragraphs), len(runs)
    ratio = n_r / n_p if n_p else 0

    print("  段落数 w:p                 %d" % n_p)
    print("  run 数 w:r                %d" % n_r)
    flag = "" if ratio >= 3 else "  ← 明显偏低，疑似一次性程序写入（真人文档通常 3–10）"
    print("  平均每段 run 数            %.2f%s" % (ratio, flag))
    print("  使用 pStyle 的段落数       %d%s" %
          (len(pstyles), "" if pstyles else "  ← 未使用任何标题样式，脚本直接设字体的特征"))

    # 段落长度变异系数 CV
    lengths = [sum(len(t.text or "") for t in p.iter(qn("w", "t"))) for p in paragraphs]
    if lengths and sum(lengths):
        mean = sum(lengths) / len(lengths)
        var = sum((x - mean) ** 2 for x in lengths) / len(lengths)
        cv = (var ** 0.5) / mean if mean else 0
        flag = "" if cv >= 0.3 else "  ← 过低，段落被均匀切块（真人评述经验上多在 0.3 以上，经验值）"
        print("  段落长度变异系数 CV       %.2f%s" % (cv, flag))
    print("  修订 w:ins / w:del         %d / %d" % (len(ins), len(dels)))

    # 批注
    has_comments = any(n.startswith("word/comments") for n in names)
    refs = list(doc_root.iter(qn("w", "commentReference")))
    print("  批注 comments.xml          %s（commentReference %d 处）" %
          ("存在" if has_comments else "不存在", len(refs)))

    # settings：trackChanges
    if "word/settings.xml" in names:
        sroot = ET.fromstring(entries["word/settings.xml"])
        tc = sroot.find(qn("w", "trackChanges"))
        print("  settings:trackChanges     %s" % ("存在（修订模式开启）" if tc is not None else "不存在"))

    # 固定模板时间戳（python-docx 默认模板 2013-12-24T07:15:00Z）
    fixed_hits = []
    if "docProps/core.xml" in entries:
        core_text = entries["docProps/core.xml"].decode("utf-8", "ignore")
        if "2013-12-24T07:15:00Z" in core_text:
            fixed_hits.append("core.xml 含 python-docx 固定时间戳 2013-12-24T07:15:00Z")
    if fixed_hits:
        for h in fixed_hits:
            print("  固定模板痕迹              %s" % h)

    # 自定义属性
    print("  docProps/custom.xml       %s" %
          ("存在（建议检查/删除）" if "docProps/custom.xml" in names else "不存在"))
    print()


def main():
    if len(sys.argv) != 2:
        print("用法: python3 inspect_metadata.py 文档.docx")
        sys.exit(1)
    path = sys.argv[1]
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            entries = {n: z.read(n) for n in names}
            core = entries.get("docProps/core.xml")
            app = entries.get("docProps/app.xml")
    except zipfile.BadZipFile as e:
        print("无法读取 docx：%s" % e)
        sys.exit(1)

    print("\n文件：%s\n" % path)
    analyze_structure(entries)
    dump_part("核心属性 docProps/core.xml", core)
    dump_part("扩展属性 docProps/app.xml", app)
    dump_part("自定义属性 docProps/custom.xml", entries.get("docProps/custom.xml"))


if __name__ == "__main__":
    main()
