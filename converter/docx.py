"""Validate an OOXML package and resolve fonts inherited by text runs."""
from io import BytesIO
from pathlib import Path
from lxml import etree as ET
from zipfile import BadZipFile, ZipFile, ZIP_DEFLATED

from .errors import ConversionError
from .fonts import family_key, FAMILIES

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
MAX_UPLOAD = 30 * 1024 * 1024
MAX_EXPANDED = 160 * 1024 * 1024
MAX_XML = 12 * 1024 * 1024
ET.register_namespace("w", W[1:-1])


def parse_xml(data: bytes):
    if len(data) > MAX_XML or b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise ConversionError("XML dokumen terlalu besar atau mengandung deklarasi yang tidak didukung.")
    try:
        return ET.fromstring(data, parser=ET.XMLParser(resolve_entities=False, no_network=True))
    except ET.ParseError as exc:
        raise ConversionError("Struktur XML DOCX tidak valid.") from exc


def _font_attrs(parent) -> dict:
    if parent is None:
        return {}
    font = parent.find(W + "rFonts")
    return dict(font.attrib) if font is not None else {}


def _merge_fonts(base: dict, override: dict) -> dict:
    result = base.copy()
    for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
        if W + slot in override or W + slot + "Theme" in override:
            result.pop(W + slot, None)
            result.pop(W + slot + "Theme", None)
    result.update(override)
    return result


def prepare_docx(data: bytes, output: Path, selected_font: str = "original") -> set[str]:
    if selected_font not in ("original", *FAMILIES):
        raise ConversionError("Pilihan font tidak valid.", 400)
    if not data or len(data) > MAX_UPLOAD:
        raise ConversionError("Ukuran DOCX harus antara 1 byte dan 30 MB.", 413)
    try:
        archive = ZipFile(BytesIO(data))
        infos = archive.infolist()
        names = [item.filename for item in infos]
        if len(infos) > 4000 or len(names) != len(set(names)) or sum(i.file_size for i in infos) > MAX_EXPANDED:
            raise ConversionError("Paket DOCX terlalu besar atau memiliki entri duplikat.")
        if not {"[Content_Types].xml", "word/document.xml", "_rels/.rels"}.issubset(names):
            raise ConversionError("File ini bukan dokumen DOCX yang valid.")
        if any(i.flag_bits & 1 for i in infos):
            raise ConversionError("DOCX terenkripsi belum didukung. Hapus password sebelum konversi.")
        parts = {}
        for item in infos:
            name = item.filename
            if "vbaproject" in name.lower() or name.startswith("word/embeddings/"):
                raise ConversionError("Dokumen dengan macro atau objek tersemat belum didukung.")
            content = archive.read(item)
            if name.endswith((".xml", ".rels")):
                root = parse_xml(content)
                parts[name] = root
                if name.endswith(".rels"):
                    for rel in root:
                        if rel.get("TargetMode") == "External" and not rel.get("Type", "").endswith("/hyperlink"):
                            raise ConversionError("Dokumen memiliki gambar atau template eksternal. Sematkan kontennya terlebih dahulu.")
        # Reject a renamed macro-enabled package, even without vbaProject.bin.
        if any("macroEnabled" in node.get("ContentType", "") for node in parts["[Content_Types].xml"]):
            raise ConversionError("Gunakan DOCX tanpa macro.")
        style_root = parts.get("word/styles.xml")
        styles = {} if style_root is None else {s.get(W + "styleId"): s for s in style_root.findall(W + "style")}
        defaults = {} if style_root is None else _font_attrs(style_root.find(W + "docDefaults/" + W + "rPrDefault/" + W + "rPr"))
        default_style = next((key for key, style in styles.items() if style.get(W + "type") == "paragraph" and style.get(W + "default") == "1"), "Normal")

        def style_fonts(style_id, seen=None):
            seen = set() if seen is None else seen
            if style_id not in styles or style_id in seen:
                return {}
            seen.add(style_id)
            style = styles[style_id]
            based = style.find(W + "basedOn")
            parent = style_fonts(based.get(W + "val"), seen) if based is not None else {}
            return _merge_fonts(parent, _font_attrs(style.find(W + "rPr")))

        theme = parts.get("word/theme/theme1.xml")
        theme_fonts = {}
        if theme is not None:
            for group in ("major", "minor"):
                latin = theme.find(".//" + A + group + "Font/" + A + "latin")
                if latin is not None:
                    theme_fonts[group + "HAnsi"] = latin.get("typeface", "")
                    theme_fonts[group + "Ascii"] = latin.get("typeface", "")

        needed = set()
        content_parts = [root for name, root in parts.items() if name in {"word/document.xml", "word/footnotes.xml", "word/endnotes.xml"} or (name.startswith(("word/header", "word/footer")) and name.endswith(".xml"))]
        for root in content_parts:
            for paragraph in root.iter(W + "p"):
                p_style = paragraph.find(W + "pPr/" + W + "pStyle")
                p_fonts = _merge_fonts(defaults, style_fonts(p_style.get(W + "val") if p_style is not None else default_style))
                for run in paragraph.iter(W + "r"):
                    if not any((text.text or "").strip() for text in run.iter(W + "t")):
                        continue
                    rpr = run.find(W + "rPr")
                    r_style = rpr.find(W + "rStyle") if rpr is not None else None
                    attrs = _merge_fonts(p_fonts, style_fonts(r_style.get(W + "val"))) if r_style is not None else p_fonts
                    attrs = _merge_fonts(attrs, _font_attrs(rpr))
                    if selected_font != "original":
                        needed.add(selected_font)
                    else:
                        for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
                            name = attrs.get(W + slot) or theme_fonts.get(attrs.get(W + slot + "Theme", ""), "")
                            family = family_key(name)
                            if family:
                                needed.add(family)

        if selected_font != "original":
            # Rewrite only text font properties; preserve sizes, styles, layout,
            # images, and numbering glyphs such as Symbol/Wingdings.
            for root in [*content_parts, *([style_root] if style_root is not None else [])]:
                for font in root.iter(W + "rFonts"):
                    font.attrib.clear()
                    font.attrib.update({W + slot: selected_font for slot in ("ascii", "hAnsi", "eastAsia", "cs")})
                for run in root.iter(W + "r"):
                    if run.find(W + "t") is None:
                        continue
                    rpr = run.find(W + "rPr")
                    if rpr is None:
                        rpr = ET.Element(W + "rPr")
                        run.insert(0, rpr)
                    if rpr.find(W + "rFonts") is None:
                        ET.SubElement(rpr, W + "rFonts", {W + slot: selected_font for slot in ("ascii", "hAnsi", "eastAsia", "cs")})
        # Normalize aliases to the family name that Word actually recognizes.
        changed = selected_font != "original"
        for root in content_parts + ([style_root] if style_root is not None else []):
            for font in root.iter(W + "rFonts"):
                for attr, name in list(font.attrib.items()):
                    canonical = family_key(name)
                    if canonical and canonical != name:
                        font.set(attr, canonical)
                        changed = True
        if not changed:
            output.write_bytes(data)
        else:
            with ZipFile(output, "w", ZIP_DEFLATED) as result:
                for item in infos:
                    content = ET.tostring(parts[item.filename], encoding="utf-8", xml_declaration=True) if item.filename in parts and (item.filename == "word/styles.xml" or parts[item.filename] in content_parts) else archive.read(item)
                    result.writestr(item, content)
        archive.close()
        return needed
    except (BadZipFile, RuntimeError, NotImplementedError, EOFError) as exc:
        raise ConversionError("DOCX rusak, terenkripsi, atau memakai kompresi yang tidak didukung.") from exc
