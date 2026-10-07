from pathlib import Path
from io import BytesIO
from fontTools.cffLib import CFFFontSet
from fontTools.ttLib import TTFont
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject

from .errors import ConversionError
from .fonts import family_key, inventory, LM_VERSION


def iter_pdf_fonts(pages):
    visited = set()

    def scan(resources):
        if not resources:
            return
        resources = resources.get_object()
        if id(resources) in visited:
            return
        visited.add(id(resources))
        fonts = resources.get('/Font', {})
        fonts = fonts.get_object() if hasattr(fonts, 'get_object') else fonts
        yield from (ref.get_object() for ref in fonts.values())
        xobjects = resources.get('/XObject', {})
        xobjects = xobjects.get_object() if hasattr(xobjects, 'get_object') else xobjects
        for ref in xobjects.values():
            yield from scan(ref.get_object().get('/Resources'))

    for page in pages:
        yield from scan(page.get('/Resources'))


def embed_supported_fonts(path: Path) -> None:
    """Word references CFF OpenType fonts without embedding them on Windows.

    Embed the exact, licensed original font program for full, simple Type1
    fonts only. Do not graft full fonts onto subsetted or CID glyph mappings.
    """
    writer = PdfWriter(clone_from=path)
    changed = False
    programs = {}
    for font in iter_pdf_fonts(writer.pages):
        name = str(font.get('/BaseFont', '')).lstrip('/')
        family = family_key(name)
        if not family or '+' in name or font.get('/Subtype') != '/Type1':
            continue
        descriptor = font.get('/FontDescriptor')
        if not descriptor:
            continue
        descriptor = descriptor.get_object()
        if any(key in descriptor for key in ('/FontFile', '/FontFile2', '/FontFile3')):
            continue
        match = next((v for v in inventory()[family]['variants'].values() if v['postscript'] == name and v['embeddable']), None)
        if not match:
            continue
        if name not in programs:
            with TTFont(match['path']) as original:
                if 'CFF ' not in original:
                    continue
                stream = DecodedStreamObject()
                stream.set_data(original.getTableData('CFF '))
                stream[NameObject('/Subtype')] = NameObject('/Type1C')
                programs[name] = writer._add_object(stream.flate_encode())
        descriptor[NameObject('/FontFile3')] = programs[name]
        changed = True
    if changed:
        embedded = path.with_name('embedded.pdf')
        writer.write(embedded)
        embedded.replace(path)


def inspect_pdf(path: Path, expected: set[str]) -> dict:
    """Inspect both page fonts and fonts nested in Form XObjects."""
    reader = PdfReader(path, strict=False)
    if reader.is_encrypted or not reader.pages:
        raise ConversionError("Mesin konversi menghasilkan PDF yang tidak valid.", 500)
    fonts = {}
    for font in iter_pdf_fonts(reader.pages):
        children = font.get("/DescendantFonts") or [font]
        name = str(font.get("/BaseFont", "Unknown")).lstrip("/").split("+")[-1]
        embedded = True
        version = None
        for child in children:
            child = child.get_object()
            descriptor = child.get("/FontDescriptor")
            descriptor = descriptor.get_object() if descriptor else {}
            embedded = embedded and any(key in descriptor for key in ("/FontFile", "/FontFile2", "/FontFile3"))
            if family_key(name) == 'LM Roman 10' and '/FontFile3' in descriptor:
                program = descriptor['/FontFile3'].get_object()
                if program.get('/Subtype') in ('/Type1C', '/CIDFontType0C'):
                    cff = CFFFontSet()
                    cff.decompile(BytesIO(program.get_data()), None)
                    version = str(getattr(cff.topDictIndex[0], 'version', ''))
        fonts[name] = {"name": name, "embedded": embedded and fonts.get(name, {"embedded": True})["embedded"], "version": version}
        if family_key(name) == 'LM Roman 10' and embedded and version != LM_VERSION:
            raise ConversionError(f"LM Roman 10 di PDF bukan versi {LM_VERSION}. Periksa versi font yang terpasang di Windows.")
    for family in expected:
        matches = [font for name, font in fonts.items() if family_key(name) == family]
        if not matches or not all(font["embedded"] for font in matches):
            raise ConversionError(
                f"Font {family} tidak terdeteksi sebagai font tertanam di PDF. "
                "Konversi dihentikan agar hasil tidak memakai font pengganti. "
                "Periksa font dan lisensinya, lalu coba lagi."
            )
    return {"pages": len(reader.pages), "fonts": list(fonts.values()), "verified_families": sorted(expected)}
