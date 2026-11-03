# Copyright (c) 2010-2024 openpyxl

"""Round-trip tests for sensitivity labels"""

import uuid
import zipfile
from io import BytesIO

import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.packaging.custom import StringProperty
from openpyxl.packaging.manifest import Manifest
from openpyxl.packaging.relationship import RelationshipList
from openpyxl.packaging.sensitivity import LabelRegistry
from openpyxl.xml.constants import (
    ARC_CONTENT_TYPES,
    ARC_CUSTOM,
    ARC_LABEL_INFO,
    ARC_ROOT_RELS,
    LABELINFO_REL,
    LABELINFO_TYPE,
)
from openpyxl.xml.functions import fromstring

SITE = "72f988bf-86f1-41af-91ab-2d7cd011db47"
PUBLIC = "11111111-aaaa-4aaa-8aaa-111111111111"
CONFIDENTIAL = "22222222-bbbb-4bbb-8bbb-222222222222"


@pytest.fixture
def registry():
    return LabelRegistry.from_dict({
        "site_id": SITE,
        "labels": [
            {"name": "Public", "id": PUBLIC},
            {"name": "Confidential", "id": CONFIDENTIAL},
        ],
    })


def save(wb):
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def package_parts(buf):
    with zipfile.ZipFile(buf) as z:
        names = z.namelist()
        rels = RelationshipList.from_tree(fromstring(z.read(ARC_ROOT_RELS)))
        manifest = Manifest.from_tree(fromstring(z.read(ARC_CONTENT_TYPES)))
        custom = z.read(ARC_CUSTOM) if ARC_CUSTOM in names else b""
    buf.seek(0)
    return names, rels, manifest, custom


def test_write_label(registry):
    wb = Workbook()
    wb.set_sensitivity_label("Confidential", registry=registry)
    names, rels, manifest, custom = package_parts(save(wb))

    assert ARC_LABEL_INFO in names
    assert [r.Target for r in rels.find(LABELINFO_REL)] == [ARC_LABEL_INFO]
    override = manifest.find(LABELINFO_TYPE)
    assert override.PartName == "/" + ARC_LABEL_INFO
    assert f"MSIP_Label_{CONFIDENTIAL}_Enabled".encode() in custom


def test_no_label_by_default():
    names, rels, manifest, custom = package_parts(save(Workbook()))
    assert ARC_LABEL_INFO not in names
    assert list(rels.find(LABELINFO_REL)) == []
    assert LABELINFO_TYPE not in manifest


def test_without_legacy_properties(registry):
    wb = Workbook()
    wb.set_sensitivity_label(registry["Public"], legacy_properties=False)
    names, rels, manifest, custom = package_parts(save(wb))
    assert ARC_LABEL_INFO in names
    assert ARC_CUSTOM not in names


def test_round_trip(registry):
    wb = Workbook()
    wb.set_sensitivity_label("Confidential", registry=registry, method="Privileged")
    wb2 = load_workbook(save(wb))

    label = wb2.sensitivity_label
    assert label.guid == uuid.UUID(CONFIDENTIAL)
    assert label.method == "Privileged"
    assert registry.lookup(label).name == "Confidential"
    # survives a second save without any changes
    wb3 = load_workbook(save(wb2))
    assert registry.lookup(wb3.sensitivity_label).name == "Confidential"


def test_change_label(registry):
    wb = Workbook()
    wb.custom_doc_props.append(StringProperty(name="Department", value="Finance"))
    wb.set_sensitivity_label("Confidential", registry=registry)
    wb = load_workbook(save(wb))

    wb.set_sensitivity_label("Public", registry=registry)
    wb = load_workbook(save(wb))

    assert registry.lookup(wb.sensitivity_label).name == "Public"
    assert len(wb.sensitivity_labels) == 1
    names = wb.custom_doc_props.names
    assert "Department" in names
    assert not [n for n in names if CONFIDENTIAL in n]
    assert f"MSIP_Label_{PUBLIC}_Name" in names


def test_remove_label(registry):
    wb = Workbook()
    wb.set_sensitivity_label("Public", registry=registry)
    wb = load_workbook(save(wb))
    wb.remove_sensitivity_label()
    assert wb.sensitivity_label is None

    names, rels, manifest, custom = package_parts(save(wb))
    assert ARC_LABEL_INFO not in names
    assert list(rels.find(LABELINFO_REL)) == []
    assert LABELINFO_TYPE not in manifest
    assert ARC_CUSTOM not in names


def test_write_only(registry):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()
    ws.append([1, 2, 3])
    wb.set_sensitivity_label("Public", registry=registry)
    wb2 = load_workbook(save(wb))
    assert registry.lookup(wb2.sensitivity_label).name == "Public"


def test_read_only_load(registry):
    wb = Workbook()
    wb.set_sensitivity_label("Public", registry=registry)
    wb2 = load_workbook(save(wb), read_only=True)
    assert registry.lookup(wb2.sensitivity_label).name == "Public"
    wb2.close()


def test_name_without_registry():
    with pytest.raises(TypeError):
        Workbook().set_sensitivity_label("Public")


def test_invalid_method(registry):
    with pytest.raises(ValueError):
        Workbook().set_sensitivity_label("Public", registry=registry, method="Auto")


def _office_style_file(target, labelinfo_path):
    """
    Build a workbook whose label part is laid out the way Office writes it:
    lower-case XML declaration, an extension element, and a custom target path.
    """
    buf = save(Workbook())
    src = zipfile.ZipFile(buf)
    out = BytesIO()
    labelinfo = f"""<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata"><clbl:label id="{{{CONFIDENTIAL}}}" enabled="1" method="Standard" siteId="{{{SITE}}}" removed="0" /><clbl:ext uri="{{00000000-0000-0000-0000-000000000000}}"/></clbl:labelList>"""
    with zipfile.ZipFile(out, "w") as dst:
        for name in src.namelist():
            data = src.read(name)
            if name == ARC_ROOT_RELS:
                data = data.replace(
                    b"</Relationships>",
                    f'<Relationship Id="rId9" Type="{LABELINFO_REL}" Target="{target}"/></Relationships>'.encode())
            elif name == ARC_CONTENT_TYPES:
                data = data.replace(
                    b"</Types>",
                    f'<Override PartName="/{labelinfo_path}" ContentType="{LABELINFO_TYPE}"/></Types>'.encode())
            dst.writestr(name, data)
        dst.writestr(labelinfo_path, labelinfo)
    out.seek(0)
    return out


@pytest.mark.parametrize("target, path", [
    ("docMetadata/LabelInfo.xml", "docMetadata/LabelInfo.xml"),
    ("/docMetadata/LabelInfo.xml", "docMetadata/LabelInfo.xml"),
    ("docMetadata/Labels.xml", "docMetadata/Labels.xml"),
])
def test_preserve_office_label(registry, target, path):
    wb = load_workbook(_office_style_file(target, path))
    assert registry.lookup(wb.sensitivity_label).name == "Confidential"

    # openpyxl previously dropped the label on save
    names, rels, manifest, custom = package_parts(save(wb))
    assert ARC_LABEL_INFO in names
    wb2 = load_workbook(save(wb))
    assert registry.lookup(wb2.sensitivity_label).name == "Confidential"


def test_invalid_label_part_warns():
    buf = save(Workbook())
    src = zipfile.ZipFile(buf)
    out = BytesIO()
    with zipfile.ZipFile(out, "w") as dst:
        for name in src.namelist():
            dst.writestr(name, src.read(name))
        dst.writestr(ARC_LABEL_INFO, "<labelList><label enabled='maybe'/></labelList>")
    out.seek(0)
    with pytest.warns(UserWarning, match="sensitivity label"):
        wb = load_workbook(out)
    assert wb.sensitivity_label is None
