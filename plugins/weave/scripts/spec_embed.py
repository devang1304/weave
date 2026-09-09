#!/usr/bin/env python3
"""
spec_embed.py - store / read the Weave content spec inside a .docx or .pptx.

Library:
  embed_spec(path, spec)   write or replace the spec part (idempotent)
  read_spec(path)          the embedded spec dict, or None
  mark_block(body, first_el, last_el, block_id)   body-level bookmark pair
  mark_slide(slide, slide_id)                     p:cSld/@name = "weave:<id>"
  bookmark_name / id_from_bookmark / slide_name / id_from_slide_name

CLI:
  spec_embed.py FILE --show            print the embedded spec as JSON (exit 1 if none)
  spec_embed.py FILE --spec spec.json  embed (or replace) the spec from a file

What embed_spec writes (SPEC.md "Embedding inside .docx / .pptx"):
  * customXml/itemN.xml       <weave:spec xmlns:weave="http://netwoven.com/weave/spec/2.0">
                              with the JSON in a CDATA section
  * customXml/itemPropsN.xml  ds:datastoreItem with a GUID generated once per
                              file (kept on re-embed) and a schemaRef to the
                              weave namespace
  * customXml/_rels/itemN.xml.rels, a customXml relationship from
    word/_rels/document.xml.rels or ppt/_rels/presentation.xml.rels, and the
    [Content_Types].xml override for itemPropsN.xml
  * docProps/custom.xml: Weave_Recipe, Weave_SpecSha256, Weave_Version (the
    part, its root relationship and content-type override are created when
    the package has none; existing MSIP_Label_* properties are never touched:
    the part is edited at string level so their bytes stay as they were)
Everything else in the zip, docMetadata/LabelInfo.xml included, is copied
byte for byte.  Block bookmarks and slide names are written by the builder
(it is the only party that knows which element belongs to which block) with
the helpers below; the naming lives here so the reader and writer agree.

Requires lxml only (python-docx / python-pptx objects are accepted by the
mark_* helpers but not imported here).  Python 3.9+.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import uuid
import zipfile

from lxml import etree

WEAVE_NS = "http://netwoven.com/weave/spec/2.0"
DS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
RT_CUSTOMXML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
RT_CUSTOMXML_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
RT_CUSTOM_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
CT_CUSTOMXML_PROPS = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
CT_CUSTOM_PROPS = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS

BOOKMARK_PREFIX = "_weave_"
SLIDE_PREFIX = "weave:"
PROP_RECIPE = "Weave_Recipe"
PROP_SHA = "Weave_SpecSha256"
PROP_VERSION = "Weave_Version"


# --------------------------------------------------------------------------
# naming helpers shared by builder and reader
# --------------------------------------------------------------------------
def safe_id(block_id):
    """Bookmark-safe form of a block id: letters, digits, underscore, at most
    32 chars (Word caps bookmark names at 40 incl. the prefix)."""
    s = re.sub(r"[^A-Za-z0-9_]", "_", str(block_id))
    return s[:32] or "x"


def bookmark_name(block_id):
    return BOOKMARK_PREFIX + safe_id(block_id)


def id_from_bookmark(name):
    """The (bookmark-safe) block id of a `_weave_*` bookmark name, else None."""
    if name and name.startswith(BOOKMARK_PREFIX):
        return name[len(BOOKMARK_PREFIX):]
    return None


def slide_name(slide_id):
    return SLIDE_PREFIX + str(slide_id)


def id_from_slide_name(name):
    if name and name.startswith(SLIDE_PREFIX):
        return name[len(SLIDE_PREFIX):]
    return None


def canonical_json(spec):
    """The exact text stored in the part (and hashed for Weave_SpecSha256)."""
    return json.dumps(spec, indent=2, ensure_ascii=False)


def spec_sha256(spec):
    return hashlib.sha256(canonical_json(spec).encode("utf-8")).hexdigest()


# --------------------------------------------------------------------------
# Word bookmarks / PowerPoint slide names
# --------------------------------------------------------------------------
def next_bookmark_id(root):
    ids = [0]
    for b in root.iter(W + "bookmarkStart"):
        try:
            ids.append(int(b.get(W + "id") or 0))
        except ValueError:
            pass
    return max(ids) + 1


def mark_block(body, first_el, last_el, block_id):
    """Wrap body children [first_el .. last_el] in a body-level bookmark pair
    `_weave_<id>`.  Body-level markers (siblings of w:p / w:tbl, as the 2026
    bases themselves use) survive python-docx's paragraph.text setter, which
    would delete an in-paragraph marker on a hand edit.  A pre-existing pair
    with the same name is removed first, so re-marking is idempotent."""
    name = bookmark_name(block_id)
    for b in list(body.iter(W + "bookmarkStart")):
        if b.get(W + "name") == name:
            bid = b.get(W + "id")
            for e in list(body.iter(W + "bookmarkEnd")):
                if e.get(W + "id") == bid:
                    e.getparent().remove(e)
            b.getparent().remove(b)
    bid = str(next_bookmark_id(body.getroottree().getroot()))
    start = etree.Element(W + "bookmarkStart")
    start.set(W + "id", bid)
    start.set(W + "name", name)
    end = etree.Element(W + "bookmarkEnd")
    end.set(W + "id", bid)
    first_el.addprevious(start)
    last_el.addnext(end)
    return start, end


def mark_slide(slide, slide_id):
    slide.name = slide_name(slide_id)
    return slide.name


# --------------------------------------------------------------------------
# zip plumbing
# --------------------------------------------------------------------------
def _rewrite(path, replacements, additions):
    """Copy the zip, swapping entries named in `replacements` and appending
    `additions` (both name -> bytes).  Untouched entries keep their bytes and
    ZipInfo (LabelInfo.xml stays byte-identical)."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        seen = set()
        for item in zin.infolist():
            seen.add(item.filename)
            data = zin.read(item.filename)
            if item.filename in replacements:
                data = replacements[item.filename]
            zout.writestr(item, data)
        for name, data in additions.items():
            if name not in seen:
                zout.writestr(name, data)
    shutil.move(tmp, path)


def _main_parts(path):
    low = path.lower()
    if low.endswith(".docx") or low.endswith(".dotx"):
        return "docx", "word/document.xml", "word/_rels/document.xml.rels"
    if low.endswith(".pptx") or low.endswith(".potx"):
        return "pptx", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"
    raise ValueError("unsupported file type: %s" % path)


def _find_weave_item(z):
    """(itemN.xml name, N, existing GUID or None) of the weave part, else
    (None, None, None)."""
    for n in z.namelist():
        m = re.match(r"customXml/item(\d+)\.xml$", n)
        if not m:
            continue
        head = z.read(n)[:600]
        if WEAVE_NS.encode("utf-8") in head:
            num = m.group(1)
            guid = None
            props = "customXml/itemProps%s.xml" % num
            if props in z.namelist():
                g = re.search(r'ds:itemID="([^"]+)"', z.read(props).decode("utf-8", "replace"))
                guid = g.group(1) if g else None
            return n, num, guid
    return None, None, None


def _next_item_number(names):
    nums = [0]
    for n in names:
        m = re.match(r"customXml/item(?:Props)?(\d+)\.xml$", n)
        if m:
            nums.append(int(m.group(1)))
    return str(max(nums) + 1)


def _next_rid(rels_xml):
    ids = [0] + [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return "rId%d" % (max(ids) + 1)


def _add_rel(rels_xml, rtype, target, mode=None):
    """Append a Relationship unless one with this target exists.  String-level
    so the other relationships keep their bytes."""
    if re.search(r'<Relationship\b[^>]*Target="%s"' % re.escape(target), rels_xml):
        return rels_xml
    rid = _next_rid(rels_xml)
    rel = '<Relationship Id="%s" Type="%s" Target="%s"%s/>' % (
        rid, rtype, target, ' TargetMode="%s"' % mode if mode else "")
    if "</Relationships>" in rels_xml:
        return rels_xml.replace("</Relationships>", rel + "</Relationships>")
    return rels_xml.replace("<Relationships/>", "<Relationships>%s</Relationships>" % rel)


def _add_override(ct_xml, partname, content_type):
    if re.search(r'<Override\b[^>]*PartName="%s"' % re.escape(partname), ct_xml):
        return ct_xml
    return ct_xml.replace("</Types>", '<Override PartName="%s" ContentType="%s"/></Types>' % (partname, content_type))


def _spec_part_xml(spec_text):
    root = etree.Element("{%s}spec" % WEAVE_NS, nsmap={"weave": WEAVE_NS})
    # lxml's own serializer already splits any embedded "]]>" when it writes
    # out a CDATA section, so do not pre-split it here too: doing both means
    # the manually-inserted "]]]]><![CDATA[>" escape (which itself contains
    # "]]>") gets re-split by lxml, corrupting the text. Let CDATA() hold the
    # raw spec text verbatim and leave the escaping to tostring().
    root.text = etree.CDATA(spec_text)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _item_props_xml(guid):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n'
            '<ds:datastoreItem ds:itemID="%s" xmlns:ds="%s"><ds:schemaRefs>'
            '<ds:schemaRef ds:uri="%s"/></ds:schemaRefs></ds:datastoreItem>' % (guid, DS_NS, WEAVE_NS)).encode("utf-8")


def _item_rels_xml(num):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s" Target="itemProps%s.xml"/>'
            '</Relationships>' % (PKG_REL_NS, RT_CUSTOMXML_PROPS, num)).encode("utf-8")


def _xml_escape(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))


def _custom_props_xml(existing, props):
    """Return docProps/custom.xml text with the Weave_* properties replaced /
    appended.  Existing (MSIP_Label_*) properties are left untouched as text."""
    if existing:
        text = existing
        # a custom.xml with zero prior properties often ships as a
        # self-closing root, <Properties .../>; expand it to the paired
        # empty form (preserving whatever attributes/namespaces it had)
        # before the "</Properties>" insertion below, which would otherwise
        # match nothing against a self-closing tag and silently return the
        # input unchanged, dropping the Weave_* properties.
        text = re.sub(r'<Properties\b([^>]*)/>', r'<Properties\1></Properties>', text)
        text = re.sub(r'<property\b[^>]*\bname="Weave_[^"]*"[^>]*/>', "", text)
        text = re.sub(r'<property\b[^>]*\bname="Weave_[^"]*"[^>]*>.*?</property>', "", text, flags=re.S)
    else:
        text = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Properties xmlns="%s" xmlns:vt="%s"></Properties>' % (CP_NS, VT_NS))
    pids = [1] + [int(x) for x in re.findall(r'\bpid="(\d+)"', text)]
    pid = max(pids)
    new = []
    for name, value in props:
        pid += 1
        new.append('<property fmtid="%s" pid="%d" name="%s"><vt:lpwstr>%s</vt:lpwstr></property>'
                   % (FMTID, pid, name, _xml_escape(value)))
    return text.replace("</Properties>", "".join(new) + "</Properties>")


# --------------------------------------------------------------------------
# public API
# --------------------------------------------------------------------------
def embed_spec(path, spec):
    """Write or replace the embedded spec in the .docx/.pptx at `path`.
    Returns {"part": "customXml/itemN.xml", "sha256": ..., "created_custom_xml": bool}."""
    kind, main_part, main_rels = _main_parts(path)
    spec_text = canonical_json(spec)
    sha = hashlib.sha256(spec_text.encode("utf-8")).hexdigest()
    weave = spec.get("weave") or {}
    props = [(PROP_RECIPE, weave.get("recipe") or weave.get("kind") or ""),
             (PROP_SHA, sha),
             (PROP_VERSION, weave.get("generator") or ("spec %s" % weave.get("spec_version", "2.0")))]

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        item, num, guid = _find_weave_item(z)
        if item is None:
            num = _next_item_number(names)
            item = "customXml/item%s.xml" % num
            guid = "{%s}" % str(uuid.uuid4()).upper()
        rels_xml = z.read(main_rels).decode("utf-8")
        ct_xml = z.read("[Content_Types].xml").decode("utf-8")
        root_rels = z.read("_rels/.rels").decode("utf-8")
        custom = z.read("docProps/custom.xml").decode("utf-8") if "docProps/custom.xml" in names else None
    if main_part not in names:
        raise ValueError("%s has no %s" % (path, main_part))

    replacements, additions = {}, {}
    spec_bytes = _spec_part_xml(spec_text)
    props_name = "customXml/itemProps%s.xml" % num
    item_rels = "customXml/_rels/item%s.xml.rels" % num
    (replacements if item in names else additions)[item] = spec_bytes
    (replacements if props_name in names else additions)[props_name] = _item_props_xml(guid)
    (replacements if item_rels in names else additions)[item_rels] = _item_rels_xml(num)
    replacements[main_rels] = _add_rel(rels_xml, RT_CUSTOMXML, "../" + item).encode("utf-8")
    ct_xml = _add_override(ct_xml, "/" + props_name, CT_CUSTOMXML_PROPS)
    created_custom = custom is None
    if created_custom:
        ct_xml = _add_override(ct_xml, "/docProps/custom.xml", CT_CUSTOM_PROPS)
        replacements["_rels/.rels"] = _add_rel(root_rels, RT_CUSTOM_PROPS, "docProps/custom.xml").encode("utf-8")
        additions["docProps/custom.xml"] = _custom_props_xml(None, props).encode("utf-8")
    else:
        replacements["docProps/custom.xml"] = _custom_props_xml(custom, props).encode("utf-8")
    replacements["[Content_Types].xml"] = ct_xml.encode("utf-8")
    _rewrite(path, replacements, additions)
    return {"part": item, "sha256": sha, "created_custom_xml": created_custom, "kind": kind}


def read_spec(path):
    """The embedded spec dict, or None when the file carries none."""
    with zipfile.ZipFile(path) as z:
        item, _, _ = _find_weave_item(z)
        if item is None:
            return None
        root = etree.fromstring(z.read(item))
    if root.tag != "{%s}spec" % WEAVE_NS:
        return None
    text = root.text or ""
    try:
        return json.loads(text)
    except ValueError:
        return None


def read_properties(path):
    """The Weave_* custom properties as a dict (empty when absent)."""
    out = {}
    with zipfile.ZipFile(path) as z:
        if "docProps/custom.xml" not in z.namelist():
            return out
        root = etree.fromstring(z.read("docProps/custom.xml"))
    for p in root.iter("{%s}property" % CP_NS):
        name = p.get("name") or ""
        if name.startswith("Weave_"):
            out[name] = "".join(p.itertext())
    return out


# --------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--show", action="store_true", help="print the embedded spec")
    ap.add_argument("--spec", default=None, help="embed this JSON file")
    args = ap.parse_args(argv)
    if not os.path.exists(args.file):
        sys.exit("error: %s not found" % args.file)
    if args.spec:
        with open(args.spec, "r", encoding="utf-8") as fh:
            spec = json.load(fh)
        info = embed_spec(args.file, spec)
        print(json.dumps({"ok": True, "file": os.path.abspath(args.file), "part": info["part"],
                          "sha256": info["sha256"]}))
        return 0
    spec = read_spec(args.file)
    if spec is None:
        print(json.dumps({"ok": False, "error": "no embedded weave spec in %s" % args.file,
                          "hint": "build the file with build_deliverable.py, or run spec_from_file.py for a best-effort spec"}))
        return 1
    print(canonical_json(spec))
    return 0


if __name__ == "__main__":
    sys.exit(main())
