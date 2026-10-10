#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
clean_metadata.py — docx 元数据清洗 / 伪装脚本（零第三方依赖，Python 3.7+）

功能：
  1. 核心属性 core.xml：作者/最后修改者改为指定姓名，标题改正常标题，
     删除 description、subject、keywords、category、revision、identifier；
  2. 扩展属性 app.xml：Application/Template 改为正常桌面编辑器，
     TotalTime 按字数估算（不再是 0），Words/Characters 从正文真实重算；
  3. 删除 docProps/custom.xml 自定义属性；
  4. 可选：删除批注（--remove-comments）、接受并清除修订（--remove-revisions）；
  5. 打印「字段变更表」用于验收。

用法：
  python3 clean_metadata.py 原稿.docx --name "杨力杰" --title "DeMamba论文评述"
  python3 clean_metadata.py 原稿.docx --name "杨力杰" --editor word \
      --editing-minutes 95 --remove-comments --remove-revisions --in-place
"""
import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

# ---------------------------------------------------------------- 命名空间
NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
for _p, _u in NS.items():
    ET.register_namespace(_p, _u)
ET.register_namespace("r", NS["pr"])  # word/_rels/document.xml.rels 常用 r 前缀

CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml"
CT_APP = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
CT_CUSTOM = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
REL_APP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
REL_CUSTOM = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"

CORE_ORDER = ["dc:creator", "cp:category", "cp:contentStatus", "dc:title",
              "dc:subject", "cp:keywords", "dc:description", "cp:lastModifiedBy",
              "cp:revision", "dcterms:created", "dcterms:modified"]
REMOVE_CORE_TAGS = ["dc:description", "dc:subject", "cp:keywords",
                    "cp:category", "cp:revision", "dc:identifier"]

EDITORS = {
    "word": {"application": "Microsoft Office Word", "appversion": "16.0000",
             "template": "Normal.dotm"},
    "wps": {"application": "WPS Office", "appversion": None,
            "template": "Normal.dotm"},
}


def qn(prefix, tag):
    return "{%s}%s" % (NS[prefix], tag)


def cqn(prefixed):
    p, t = prefixed.split(":")
    return qn(p, t)


# ---------------------------------------------------------------- 统计
CJK_RE = re.compile(
    r"[㐀-䶿一-鿿豈-﫿]"  # CJK 扩展A + 基本 + 兼容
)
LATIN_WORD_RE = re.compile(r"[A-Za-z0-9]+(?:[.\-'][A-Za-z0-9]+)*")


def extract_paragraphs(document_xml):
    """从 word/document.xml 提取段落纯文本；不含 delText / 域代码。"""
    root = ET.fromstring(document_xml)
    paragraphs = []
    for p in root.iter(qn("w", "p")):
        buf = []
        for node in p.iter():
            if node.tag == qn("w", "t"):
                buf.append(node.text or "")
            elif node.tag == qn("w", "tab"):
                buf.append("\t")
            elif node.tag in (qn("w", "br"), qn("w", "cr")):
                buf.append("\n")
        paragraphs.append("".join(buf))
    return paragraphs


def calc_stats(paragraphs):
    all_text = "\n".join(paragraphs)
    cjk = len(CJK_RE.findall(all_text))
    latin_words = len(LATIN_WORD_RE.findall(all_text))
    words = cjk + latin_words
    chars_no_spaces = sum(1 for ch in all_text if ch not in " \t\n\r")
    chars_with_spaces = sum(1 for ch in all_text if ch not in "\t\n\r")
    paragraphs_n = len(paragraphs)
    pages = max(1, round(chars_no_spaces / 700))
    lines = max(1, pages * 44)
    return {
        "Words": str(words),
        "Characters": str(chars_no_spaces),
        "CharactersWithSpaces": str(chars_with_spaces),
        "Paragraphs": str(paragraphs_n),
        "Pages": str(pages),
        "Lines": str(lines),
    }


def estimate_minutes(words):
    return max(10, min(180, round(words / 40)))


# ---------------------------------------------------------------- XML 辅助
def find_child(parent, prefixed):
    return parent.find(cqn(prefixed))


def upsert_core(root, prefixed, text):
    """按 CT_CoreProperties 的顺序插入/赋值。"""
    elem = find_child(root, prefixed)
    if elem is None:
        elem = ET.Element(cqn(prefixed))
        idx = CORE_ORDER.index(prefixed)
        insert_at = len(root)
        for i, child in enumerate(root):
            cpref = None
            for p, u in NS.items():
                if child.tag.startswith("{%s}" % u):
                    cpref = "%s:%s" % (p, child.tag[len(u) + 2:])
                    break
            if cpref in CORE_ORDER and CORE_ORDER.index(cpref) > idx:
                insert_at = i
                break
        root.insert(insert_at, elem)
    elem.text = text
    return elem


def remove_children(root, prefixed_list, changes, part):
    for prefixed in prefixed_list:
        for elem in root.findall(cqn(prefixed)):
            changes.append((part, prefixed, (elem.text or "").strip() or "(空)", "已删除"))
            root.remove(elem)


def app_get(root, tag):
    elem = root.find(cqn("ep:" + tag))
    return elem


def app_set(root, tag, value, changes):
    elem = root.find(cqn("ep:" + tag))
    old = elem.text if elem is not None and elem.text else "(缺失)"
    if elem is None:
        elem = ET.SubElement(root, cqn("ep:" + tag))
    elem.text = value
    if old != value:
        changes.append(("docProps/app.xml", tag, old, value))


def app_remove(root, tag, changes):
    for elem in root.findall(cqn("ep:" + tag)):
        changes.append(("docProps/app.xml", tag, elem.text or "(空)", "已删除"))
        root.remove(elem)


# ---------------------------------------------------------------- 包级辅助
def get_parts(contents_xml, rels_xml, partname=None, reltype=None, content_type=None):
    """根据 ContentType Override 或 Relationship 返回部件路径（去掉开头 /）。"""
    ct_root = ET.fromstring(contents_xml)
    targets = set()
    if content_type:
        for o in ct_root.findall(qn("ct", "Override")):
            if o.get("ContentType") == content_type:
                targets.add(o.get("PartName").lstrip("/"))
    if reltype:
        rel_root = ET.fromstring(rels_xml)
        for r in rel_root.findall(qn("pr", "Relationship")):
            if r.get("Type") == reltype:
                targets.add(r.get("Target").lstrip("/"))
    return targets


def remove_ct_override(ct_root, partname):
    partname = partname if partname.startswith("/") else "/" + partname
    for o in ct_root.findall(qn("ct", "Override")):
        if o.get("PartName") == partname:
            ct_root.remove(o)
            return True
    return False


def remove_rel(rel_root, reltype=None, target_suffix=None):
    for r in rel_root.findall(qn("pr", "Relationship")):
        if (reltype and r.get("Type") == reltype) or \
           (target_suffix and r.get("Target", "").endswith(target_suffix)):
            rel_root.remove(r)
            return True
    return False


def next_rid(rel_root):
    used = {r.get("Id") for r in rel_root.findall(qn("pr", "Relationship"))}
    i = 1
    while "rId%d" % i in used:
        i += 1
    return "rId%d" % i


def add_ct_override(ct_root, partname, content_type):
    partname = partname if partname.startswith("/") else "/" + partname
    for o in ct_root.findall(qn("ct", "Override")):
        if o.get("PartName") == partname:
            return
    ET.SubElement(ct_root, qn("ct", "Override"),
                  {"PartName": partname, "ContentType": content_type})


def add_rel(rel_root, reltype, target):
    for r in rel_root.findall(qn("pr", "Relationship")):
        if r.get("Type") == reltype and r.get("Target") == target:
            return
    ET.SubElement(rel_root, qn("pr", "Relationship"),
                  {"Id": next_rid(rel_root), "Type": reltype, "Target": target})


# ---------------------------------------------------------------- 批注 / 修订
def strip_comments_in_document(document_xml):
    root = ET.fromstring(document_xml)
    parents = {c: p for p in root.iter() for c in p}
    kill_tags = {qn("w", "commentRangeStart"), qn("w", "commentRangeEnd")}
    for elem in list(root.iter()):
        if elem.tag in kill_tags:
            parents[elem].remove(elem)
    # 含 commentReference 的 run 整体删除
    for r in list(root.iter(qn("w", "r"))):
        if r.find(qn("w", "commentReference")) is not None:
            parents[r].remove(r)
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")


def accept_revisions(document_xml):
    root = ET.fromstring(document_xml)
    parents = {c: p for p in root.iter() for c in p}

    def unwrap(elem):
        parent = parents[elem]
        idx = list(parent).index(elem)
        for child in list(elem):
            parent.insert(idx, child)
            parents[child] = parent
            idx += 1
        parent.remove(elem)

    # w:ins / w:moveTo：接受（解开包裹）；w:del / w:moveFrom：连内容删除
    for tag_unwrap in (qn("w", "ins"), qn("w", "moveTo")):
        for elem in [e for e in root.iter() if e.tag == tag_unwrap]:
            if elem in parents:
                unwrap(elem)
    for tag_kill in (qn("w", "del"), qn("w", "moveFrom")):
        for elem in [e for e in root.iter() if e.tag == tag_kill]:
            if elem in parents:
                parents[elem].remove(elem)
    # 各类 *Change 格式修订记录直接删除
    for elem in [e for e in root.iter() if e.tag.split("}")[-1].endswith("Change")]:
        if elem in parents:
            parents[elem].remove(elem)
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser(description="docx 元数据清洗脚本")
    ap.add_argument("docx", help="原始 docx 路径")
    ap.add_argument("--name", required=True, help="作者真实姓名（写入 creator / lastModifiedBy）")
    ap.add_argument("--title", default=None, help="文档标题")
    ap.add_argument("--editor", choices=list(EDITORS), default="word")
    ap.add_argument("--editing-minutes", type=int, default=None,
                    help="编辑时长（分钟），默认按字数估算")
    ap.add_argument("--in-place", action="store_true", help="原地覆盖文件")
    ap.add_argument("-o", "--output", default=None, help="输出路径")
    ap.add_argument("--keep-custom", action="store_true", help="保留 custom.xml")
    ap.add_argument("--remove-comments", action="store_true", help="删除批注")
    ap.add_argument("--remove-revisions", action="store_true", help="接受并清除修订痕迹")
    ap.add_argument("--touch", action="store_true", help="最后修改时间刷新为当前时间")
    args = ap.parse_args()

    if not os.path.isfile(args.docx):
        print("找不到文件：%s" % args.docx)
        sys.exit(1)

    with zipfile.ZipFile(args.docx) as z:
        entries = {n: z.read(n) for n in z.namelist()}

    contents = "[Content_Types].xml"
    rels = "_rels/.rels"
    ct_root = ET.fromstring(entries[contents])
    rel_root = ET.fromstring(entries[rels])
    changes = []

    # ---- 定位 core / app / custom 部件
    core_parts = get_parts(entries[contents], entries[rels],
                           content_type=CT_CORE, reltype=REL_CORE)
    app_parts = get_parts(entries[contents], entries[rels],
                          content_type=CT_APP, reltype=REL_APP)
    custom_parts = get_parts(entries[contents], entries[rels],
                             content_type=CT_CUSTOM, reltype=REL_CUSTOM)
    core_part = next(iter(core_parts), "docProps/core.xml")
    app_part = next(iter(app_parts), "docProps/app.xml")

    # ---- 正文统计（先做，后面 app 要用）
    doc_part = "word/document.xml"
    paragraphs = extract_paragraphs(entries[doc_part])
    stats = calc_stats(paragraphs)
    words_n = int(stats["Words"])

    # ================= core.xml =================
    if core_part in entries:
        core_root = ET.fromstring(entries[core_part])
    else:
        core_root = ET.Element(qn("cp", "coreProperties"))
        add_ct_override(ct_root, core_part, CT_CORE)
        add_rel(rel_root, REL_CORE, core_part)

    for prefixed, val in (("dc:creator", args.name),
                          ("cp:lastModifiedBy", args.name)):
        elem = find_child(core_root, prefixed)
        old = elem.text if elem is not None and elem.text else "(缺失)"
        upsert_core(core_root, prefixed, val)
        if old != val:
            changes.append(("docProps/core.xml", prefixed, old, val))

    if args.title:
        elem = find_child(core_root, "dc:title")
        old = elem.text if elem is not None and elem.text else "(缺失)"
        upsert_core(core_root, "dc:title", args.title)
        if old != args.title:
            changes.append(("docProps/core.xml", "dc:title", old, args.title))

    remove_children(core_root, REMOVE_CORE_TAGS, changes, "docProps/core.xml")

    if args.touch:
        from datetime import datetime, timezone
        now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        for prefixed in ("dcterms:created", "dcterms:modified"):
            elem = find_child(core_root, prefixed)
            old = elem.text if elem is not None and elem.text else "(缺失)"
            upsert_core(core_root, prefixed, now)
            elem.set(qn("xsi", "type"), "dcterms:W3CDTF")
            if old != now:
                changes.append(("docProps/core.xml", prefixed, old, now))

    entries[core_part] = ET.tostring(core_root, xml_declaration=True,
                                     encoding="UTF-8")

    # ================= app.xml =================
    if app_part in entries:
        app_root = ET.fromstring(entries[app_part])
    else:
        app_root = ET.Element(qn("ep", "Properties"))
        add_ct_override(ct_root, app_part, CT_APP)
        add_rel(rel_root, REL_APP, app_part)

    ed = EDITORS[args.editor]
    app_set(app_root, "Template", ed["template"], changes)
    app_set(app_root, "Application", ed["application"], changes)
    if ed["appversion"]:
        app_set(app_root, "AppVersion", ed["appversion"], changes)
    else:
        app_remove(app_root, "AppVersion", changes)
    app_remove(app_root, "Company", changes)
    app_remove(app_root, "Manager", changes)

    minutes = args.editing_minutes
    if minutes is None:
        minutes = estimate_minutes(words_n)
    app_set(app_root, "TotalTime", str(minutes), changes)

    for tag in ("Pages", "Words", "Characters", "CharactersWithSpaces",
                "Paragraphs", "Lines"):
        app_set(app_root, tag, stats[tag], changes)

    entries[app_part] = ET.tostring(app_root, xml_declaration=True,
                                    encoding="UTF-8")

    # ================= custom.xml =================
    if custom_parts and not args.keep_custom:
        for cp in custom_parts:
            if cp in entries:
                del entries[cp]
            remove_ct_override(ct_root, cp)
        remove_rel(rel_root, reltype=REL_CUSTOM)
        changes.append(("docProps/custom.xml", "自定义属性", "存在", "已整体删除"))

    # ================= 批注 =================
    if args.remove_comments:
        drels_part = "word/_rels/document.xml.rels"
        drels_root = ET.fromstring(entries[drels_part])
        kill_parts = []
        for name in list(entries):
            base = name.split("/")[-1]
            if name.startswith("word/") and (base.startswith("comments") or base == "people.xml"):
                kill_parts.append(name)
        for kp in kill_parts:
            del entries[kp]
            remove_ct_override(ct_root, kp)
            target = kp[len("word/"):]
            remove_rel(drels_root, target_suffix=target)
            changes.append((kp, "批注部件", "存在", "已删除"))
        entries[drels_part] = ET.tostring(drels_root, xml_declaration=True,
                                          encoding="UTF-8")
        entries[doc_part] = strip_comments_in_document(entries[doc_part])
        changes.append(("word/document.xml", "批注引用", "存在", "已清除"))

    # ================= 修订 =================
    if args.remove_revisions:
        entries[doc_part] = accept_revisions(entries[doc_part])
        settings = "word/settings.xml"
        if settings in entries:
            sroot = ET.fromstring(entries[settings])
            if sroot.find(qn("w", "trackChanges")) is not None:
                sroot.remove(sroot.find(qn("w", "trackChanges")))
                entries[settings] = ET.tostring(sroot, xml_declaration=True,
                                                encoding="UTF-8")
        changes.append(("word/document.xml", "修订痕迹", "存在", "已接受并清除"))

    # ---- 回写 ContentTypes / Rels
    entries[contents] = ET.tostring(ct_root, xml_declaration=True,
                                    encoding="UTF-8")
    entries[rels] = ET.tostring(rel_root, xml_declaration=True,
                                encoding="UTF-8")

    # ---- 输出
    if args.output:
        out_path = args.output
    elif args.in_place:
        out_path = args.docx
    else:
        stem, ext = os.path.splitext(args.docx)
        out_path = stem + "_deai" + (ext or ".docx")

    fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in entries.items():
                z.writestr(name, data)
        # 校验：新 zip 可打开且 XML 可解析
        with zipfile.ZipFile(tmp_path) as z:
            bad = z.testzip()
            if bad:
                raise zipfile.BadZipFile("损坏条目: %s" % bad)
            for n in z.namelist():
                if n.endswith(".xml") or n.endswith(".rels"):
                    ET.fromstring(z.read(n))
        shutil.move(tmp_path, out_path)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    # ---- 字段变更表
    print("\n元数据清洗完成：%s\n" % out_path)
    print("%-22s %-18s %-30s %s" % ("所在部件", "字段", "原值", "现值"))
    print("-" * 92)
    seen = set()
    for part, field, old, new in changes:
        key = (part, field, old, new)
        if key in seen:
            continue
        seen.add(key)
        old_s = old if len(old) <= 28 else old[:27] + "…"
        new_s = new if len(new) <= 30 else new[:29] + "…"
        print("%-22s %-18s %-30s %s" % (part, field, old_s, new_s))
    if not changes:
        print("（未发现需要修改的字段）")
    print("\n下一步：用 Word/WPS 打开「%s」，通读并随手改几句后保存。\n" % out_path)


if __name__ == "__main__":
    main()
