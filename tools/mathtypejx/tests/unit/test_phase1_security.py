"""Phase 1 security, OPC traversal and transactional writer gates."""

from __future__ import annotations

import hashlib
import struct
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from mathtypejx.extractor import HEADER_SIZE, _parse_header, extract_mtef
from mathtypejx.models import FormulaInfo, FormulaStatus
from mathtypejx.package import PackageLimits, PackageValidationError
from mathtypejx.replacer import replace_formulas
from mathtypejx.scanner import read_ole_binary, scan_docx

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
O = "urn:schemas-microsoft-com:office:office"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL = f"{R}/officeDocument"
OLE_REL = f"{R}/oleObject"
HEADER_REL = f"{R}/header"


def _content_types() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''


def _root_rels() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="{PR}"><Relationship Id="rDoc" Type="{OFFICE_REL}" Target="word/document.xml"/></Relationships>'''


def _document(body: str) -> str:
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:o="{O}"><w:body>{body}</w:body></w:document>'''


def _relationships(*items: tuple[str, str, str]) -> str:
    rows = "".join(
        f'<Relationship Id="{rid}" Type="{kind}" Target="{target}"/>'
        for rid, kind, target in items
    )
    return f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PR}">{rows}</Relationships>'


def _write_docx(path: Path, entries: dict[str, bytes | str]) -> None:
    base = {"[Content_Types].xml": _content_types(), "_rels/.rels": _root_rels()}
    base.update(entries)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, value in base.items():
            archive.writestr(name, value)


def _ole(rid: str, prog_id: str = "Equation.DSMT4") -> str:
    return f'<w:object><o:OLEObject r:id="{rid}" ProgID="{prog_id}"/></w:object>'


def _formula(fid: str, rid: str, ole_part: str = "word/embeddings/shared.bin") -> FormulaInfo:
    return FormulaInfo(
        formula_id=fid,
        ole_name=Path(ole_part).name,
        ole_part_name=ole_part,
        part_name="word/document.xml",
        rels_path="word/_rels/document.xml.rels",
        relationship_id=rid,
        prog_id="Equation.DSMT4",
        status=FormulaStatus.CONVERTED,
        omml=f'<m:oMath xmlns:m="{M}"><m:r><m:t>{fid}</m:t></m:r></m:oMath>',
    )


def test_zip_traversal_is_rejected(tmp_path):
    source = tmp_path / "traversal.docx"
    _write_docx(source, {"word/document.xml": _document("<w:p/>"), "../escape": "bad"})
    with pytest.raises(PackageValidationError, match="traversal"):
        scan_docx(str(source))
    assert not (tmp_path / "escape").exists()


def test_zip_bomb_metadata_is_rejected_before_extraction(tmp_path):
    source = tmp_path / "oversized.docx"
    _write_docx(source, {"word/document.xml": _document("<w:p/>"), "large.bin": b"x" * 128})
    limits = PackageLimits(uncompressed_package_max=64)
    with pytest.raises(PackageValidationError, match="Uncompressed DOCX"):
        scan_docx(str(source), limits=limits)


def test_digitally_signed_docx_is_blocked(tmp_path):
    source = tmp_path / "signed.docx"
    _write_docx(source, {
        "word/document.xml": _document("<w:p/>"),
        "_xmlsignatures/sig1.xml": "<Signature/>",
    })
    with pytest.raises(PackageValidationError, match="Digitally signed"):
        scan_docx(str(source))


@pytest.mark.parametrize("payload", [
    "<w:document",
    f'<!DOCTYPE x [<!ENTITY leak SYSTEM "file:///etc/passwd">]><w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>&leak;</w:t></w:r></w:p></w:body></w:document>',
])
def test_malformed_or_xxe_xml_fails_closed(tmp_path, payload):
    source = tmp_path / "bad-xml.docx"
    _write_docx(source, {"word/document.xml": payload})
    with pytest.raises(PackageValidationError):
        scan_docx(str(source))


def test_broken_ole_relationship_is_detected_but_not_read(tmp_path):
    source = tmp_path / "broken-rel.docx"
    _write_docx(source, {"word/document.xml": _document(f"<w:p><w:r>{_ole('missing')}</w:r></w:p>")})
    formulas, workdir = scan_docx(str(source))
    assert len(formulas) == 1
    assert formulas[0].ole_part_name == ""
    assert read_ole_binary(workdir, formulas[0].ole_name, formulas[0].ole_part_name) is None


def test_dynamic_header_relationship_discovers_header_27(tmp_path):
    source = tmp_path / "header.docx"
    _write_docx(source, {
        "word/document.xml": _document("<w:p/>") ,
        "word/_rels/document.xml.rels": _relationships(("rHeader", HEADER_REL, "header27.xml")),
        "word/header27.xml": f'<w:hdr xmlns:w="{W}" xmlns:r="{R}" xmlns:o="{O}"><w:p><w:r>{_ole("rOle")}</w:r></w:p></w:hdr>',
        "word/_rels/header27.xml.rels": _relationships(("rOle", OLE_REL, "embeddings/header-equation.bin")),
        "word/embeddings/header-equation.bin": b"ole",
    })
    formulas, workdir = scan_docx(str(source))
    assert [(f.part_name, f.ole_part_name) for f in formulas] == [
        ("word/header27.xml", "word/embeddings/header-equation.bin")
    ]
    assert read_ole_binary(workdir, formulas[0].ole_name, formulas[0].ole_part_name) == b"ole"


def test_relationship_target_cannot_escape_package(tmp_path):
    source = tmp_path / "bad-target.docx"
    _write_docx(source, {
        "word/document.xml": _document("<w:p/>") ,
        "word/_rels/document.xml.rels": _relationships(("rHeader", HEADER_REL, "../../outside.xml")),
    })
    with pytest.raises(PackageValidationError, match="escapes package root"):
        scan_docx(str(source))


def test_multiple_formulas_in_one_run_preserve_intervening_text(tmp_path):
    source = tmp_path / "two-in-run.docx"
    body = f'<w:p><w:r>{_ole("r1")}<w:t>between</w:t>{_ole("r2")}</w:r></w:p>'
    _write_docx(source, {
        "word/document.xml": _document(body),
        "word/_rels/document.xml.rels": _relationships(
            ("r1", OLE_REL, "embeddings/one.bin"),
            ("r2", OLE_REL, "embeddings/two.bin"),
        ),
        "word/embeddings/one.bin": b"one",
        "word/embeddings/two.bin": b"two",
    })
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    formulas, workdir = scan_docx(str(source))
    for formula in formulas:
        formula.status = FormulaStatus.CONVERTED
        formula.omml = f'<m:oMath xmlns:m="{M}"><m:r><m:t>{formula.formula_id}</m:t></m:r></m:oMath>'
    output = tmp_path / "converted.docx"
    replace_formulas(workdir, formulas, str(output))
    with zipfile.ZipFile(output) as archive:
        xml = archive.read("word/document.xml")
    root = etree.fromstring(xml)
    assert len(root.findall(f".//{{{M}}}oMath")) == 2
    assert "between" in "".join(root.itertext())
    assert all(formula.status == FormulaStatus.REPLACED for formula in formulas)
    assert hashlib.sha256(source.read_bytes()).hexdigest() == source_hash


def test_shared_embedding_is_not_deleted_while_still_referenced(tmp_path):
    source = tmp_path / "shared.docx"
    body = f'<w:p><w:r>{_ole("r1")}</w:r><w:r>{_ole("r2")}</w:r></w:p>'
    _write_docx(source, {
        "word/document.xml": _document(body),
        "word/_rels/document.xml.rels": _relationships(
            ("r1", OLE_REL, "embeddings/shared.bin"),
            ("r2", OLE_REL, "embeddings/shared.bin"),
        ),
        "word/embeddings/shared.bin": b"shared",
        "custom/unknown.dat": b"preserve-me",
    })
    formulas, workdir = scan_docx(str(source))
    output = tmp_path / "shared-output.docx"
    replace_formulas(workdir, [_formula("F0001", "r1")], str(output))
    with zipfile.ZipFile(output) as archive:
        names = set(archive.namelist())
        rels = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
        assert archive.read("custom/unknown.dat") == b"preserve-me"
    assert "word/embeddings/shared.bin" in names
    assert {rel.get("Id") for rel in rels} == {"r2"}


def test_unreferenced_embedding_and_override_are_removed(tmp_path):
    source = tmp_path / "single.docx"
    content_types = _content_types().replace(
        "</Types>",
        '<Override PartName="/word/embeddings/one.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>',
    )
    _write_docx(source, {
        "[Content_Types].xml": content_types,
        "word/document.xml": _document(f'<w:p><w:r>{_ole("r1")}</w:r></w:p>'),
        "word/_rels/document.xml.rels": _relationships(("r1", OLE_REL, "embeddings/one.bin")),
        "word/embeddings/one.bin": b"one",
    })
    formulas, workdir = scan_docx(str(source))
    formulas[0].status = FormulaStatus.CONVERTED
    formulas[0].omml = f'<m:oMath xmlns:m="{M}"><m:r><m:t>x</m:t></m:r></m:oMath>'
    output = tmp_path / "single-output.docx"
    replace_formulas(workdir, formulas, str(output))
    with zipfile.ZipFile(output) as archive:
        assert "word/embeddings/one.bin" not in archive.namelist()
        assert b"/word/embeddings/one.bin" not in archive.read("[Content_Types].xml")


def test_output_collision_is_rejected_without_modifying_existing_file(tmp_path):
    source = tmp_path / "input.docx"
    _write_docx(source, {"word/document.xml": _document("<w:p/>")})
    _, workdir = scan_docx(str(source))
    output = tmp_path / "exists.docx"
    output.write_bytes(b"sentinel")
    with pytest.raises(FileExistsError):
        replace_formulas(workdir, [], str(output))
    assert output.read_bytes() == b"sentinel"


def test_extractor_rejects_oversized_ole_and_truncated_mtef_header(mock_ole_binary):
    formula = FormulaInfo("F1", "one.bin", "word/document.xml", "rels", "r1", "Equation.DSMT4")
    assert not extract_mtef(formula, mock_ole_binary, PackageLimits(ole_part_max=1))
    assert "size limit" in formula.error_message

    header = struct.pack("<H I H I I I I I", HEADER_SIZE, 0, 0, 50, 0, 0, 0, 0)
    assert _parse_header(header + b"short") is None
