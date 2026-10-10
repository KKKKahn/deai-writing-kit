#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inspect_metadata.py — 查看 docx 的全部元数据（核心属性 / 扩展属性 / 自定义属性）
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
        # 自定义属性的名称在 name 属性上
        name = elem.get(qn("vt", "lpwstr"))
        if elem.tag.endswith("}property"):
            print("  [property] %s" % elem.get("name"))
            continue
        if text or elem.attrib:
            extra = ""
            xtype = elem.get("{http://www.w3.org/2001/XMLSchema-instance}type")
            if xtype:
                extra = "  (xsi:type=%s)" % xtype
            print("  %-28s %s%s" % (pretty_tag(elem.tag), text or "(空)", extra))
    print()


def main():
    if len(sys.argv) != 2:
        print("用法: python3 inspect_metadata.py 文档.docx")
        sys.exit(1)
    path = sys.argv[1]
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            core = z.read("docProps/core.xml") if "docProps/core.xml" in names else None
            app = z.read("docProps/app.xml") if "docProps/app.xml" in names else None
            custom = z.read("docProps/custom.xml") if "docProps/custom.xml" in names else None
    except (zipfile.BadZipFile, KeyError) as e:
        print("无法读取 docx：%s" % e)
        sys.exit(1)

    print("\n文件：%s\n" % path)
    dump_part("核心属性 docProps/core.xml", core)
    dump_part("扩展属性 docProps/app.xml", app)
    dump_part("自定义属性 docProps/custom.xml", custom)


if __name__ == "__main__":
    main()
