# Copyright (c) 2010-2024 openpyxl
import json
import uuid
from datetime import datetime, timezone

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml
from openpyxl.packaging.custom import CustomPropertyList, StringProperty

from ..sensitivity import (
    ClassificationLabel,
    ClassificationLabelList,
    LabelRegistry,
    SensitivityLabel,
    _legacy_props,
    _strip_legacy_props,
)

SITE = "72f988bf-86f1-41af-91ab-2d7cd011db47"
PUBLIC = "11111111-aaaa-4aaa-8aaa-111111111111"
CONFIDENTIAL = "22222222-bbbb-4bbb-8bbb-222222222222"


@pytest.fixture
def config():
    return {
        "site_id": SITE,
        "labels": [
            {"name": "Public", "id": PUBLIC},
            {"name": "Confidential", "id": CONFIDENTIAL.upper(), "content_bits": 2},
        ],
    }


class TestClassificationLabel:

    def test_to_tree(self):
        label = ClassificationLabel(id="{%s}" % PUBLIC, siteId="{%s}" % SITE)
        expected = f"""
        <clbl:label xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata"
          id="{{{PUBLIC}}}" enabled="1" method="Standard" siteId="{{{SITE}}}"
          contentBits="0" removed="0" />
        """
        diff = compare_xml(tostring(label.to_tree()), expected)
        assert diff is None, diff


    def test_from_tree(self):
        src = f"""
        <clbl:label xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata"
          id="{{{PUBLIC}}}" enabled="1" method="Privileged" siteId="{{{SITE}}}"
          contentBits="3" removed="0" />
        """
        label = ClassificationLabel.from_tree(fromstring(src))
        assert label.guid == uuid.UUID(PUBLIC)
        assert label.site_guid == uuid.UUID(SITE)
        assert label.method == "Privileged"
        assert label.contentBits == 3
        assert label.active


    @pytest.mark.parametrize("enabled, removed, active", [
        (True, False, True),
        (False, False, False),
        (True, True, False),
    ])
    def test_active(self, enabled, removed, active):
        label = ClassificationLabel(id=PUBLIC, siteId=SITE,
                                    enabled=enabled, removed=removed)
        assert label.active is active


class TestClassificationLabelList:

    def test_write(self):
        labels = ClassificationLabelList(label=[
            ClassificationLabel(id="{%s}" % PUBLIC, siteId="{%s}" % SITE)
        ])
        xml = labels._write()
        assert xml.startswith(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        assert b"<clbl:labelList" in xml
        expected = f"""
        <clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata">
          <clbl:label id="{{{PUBLIC}}}" enabled="1" method="Standard"
            siteId="{{{SITE}}}" contentBits="0" removed="0" />
        </clbl:labelList>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_read_office_file(self):
        # As written by Office, including extension elements we ignore
        src = f"""<?xml version="1.0" encoding="utf-8" standalone="yes"?>
        <clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata">
          <clbl:label id="{{{PUBLIC}}}" enabled="0" method="Standard"
            siteId="{{{SITE}}}" removed="1" />
          <clbl:label id="{{{CONFIDENTIAL}}}" enabled="1" method="Privileged"
            siteId="{{{SITE}}}" contentBits="0" removed="0" />
          <clbl:ext uri="{{00000000-0000-0000-0000-000000000000}}" />
        </clbl:labelList>
        """.encode()
        labels = ClassificationLabelList.from_tree(fromstring(src))
        assert len(labels) == 2
        assert labels.active.guid == uuid.UUID(CONFIDENTIAL)


    def test_no_active(self):
        assert ClassificationLabelList().active is None


class TestSensitivityLabel:

    def test_normalises_guids(self):
        label = SensitivityLabel("Public", "{%s}" % PUBLIC.upper(), SITE)
        assert label.id == uuid.UUID(PUBLIC)
        assert label.site_id == uuid.UUID(SITE)


    @pytest.mark.parametrize("kw", [
        dict(name="", id=PUBLIC, site_id=SITE),
        dict(name="Public", id="not-a-guid", site_id=SITE),
        dict(name="Public", id=PUBLIC, site_id="nope"),
        dict(name="Public", id=PUBLIC, site_id=SITE, content_bits=-1),
    ])
    def test_invalid(self, kw):
        with pytest.raises(ValueError):
            SensitivityLabel(**kw)


class TestLabelRegistry:

    def test_from_dict(self, config):
        reg = LabelRegistry.from_dict(config)
        assert len(reg) == 2
        assert reg.names == ["Public", "Confidential"]
        assert reg["confidential"].id == uuid.UUID(CONFIDENTIAL)
        assert reg["Confidential"].content_bits == 2
        assert "PUBLIC" in reg
        assert "Secret" not in reg


    def test_from_json(self, config, tmpdir):
        path = tmpdir.join("labels.json")
        path.write(json.dumps(config))
        reg = LabelRegistry.from_json(str(path))
        assert reg["Public"].site_id == uuid.UUID(SITE)


    def test_per_label_site(self, config):
        other = "33333333-cccc-4ccc-8ccc-333333333333"
        del config["site_id"]
        for entry in config["labels"]:
            entry["site_id"] = other
        reg = LabelRegistry.from_dict(config)
        assert reg["Public"].site_id == uuid.UUID(other)


    def test_unknown_name(self, config):
        reg = LabelRegistry.from_dict(config)
        with pytest.raises(KeyError, match="Known labels: Public, Confidential"):
            reg["Secret"]


    def test_lookup(self, config):
        reg = LabelRegistry.from_dict(config)
        assert reg.lookup("{%s}" % PUBLIC.upper()).name == "Public"
        label = ClassificationLabel(id="{%s}" % CONFIDENTIAL, siteId=SITE)
        assert reg.lookup(label).name == "Confidential"
        assert reg.lookup("44444444-dddd-4ddd-8ddd-444444444444") is None
        assert reg.lookup(None) is None


    @pytest.mark.parametrize("change, message", [
        (lambda c: c["labels"].append({"name": "public", "id": str(uuid.uuid4())}),
         "Duplicate label name"),
        (lambda c: c["labels"].append({"name": "Other", "id": PUBLIC}),
         "share id"),
        (lambda c: c["labels"].append({"name": "Other"}),
         "missing id"),
        (lambda c: c.pop("site_id"), "No site_id"),
        (lambda c: c.pop("labels"), "'labels' list"),
    ])
    def test_invalid_config(self, config, change, message):
        change(config)
        with pytest.raises(ValueError, match=message):
            LabelRegistry.from_dict(config)


    def test_not_a_dict(self):
        with pytest.raises(ValueError):
            LabelRegistry.from_dict([])


class TestLegacyProperties:

    def test_props(self):
        label = SensitivityLabel("Confidential", CONFIDENTIAL, SITE, content_bits=1)
        action = uuid.UUID("55555555-eeee-4eee-8eee-555555555555")
        when = datetime(2024, 5, 1, 12, 30, 15, tzinfo=timezone.utc)
        props = _legacy_props(label, "Privileged", set_date=when, action_id=action)
        prefix = f"MSIP_Label_{CONFIDENTIAL}_"
        assert {p.name: p.value for p in props} == {
            prefix + "Enabled": "true",
            prefix + "SetDate": "2024-05-01T12:30:15Z",
            prefix + "Method": "Privileged",
            prefix + "Name": "Confidential",
            prefix + "SiteId": SITE,
            prefix + "ActionId": str(action),
            prefix + "ContentBits": "1",
        }


    def test_strip(self):
        props = CustomPropertyList()
        props.append(StringProperty(name="Department", value="Finance"))
        for p in _legacy_props(SensitivityLabel("Public", PUBLIC, SITE), "Standard"):
            props.append(p)
        _strip_legacy_props(props)
        assert props.names == ["Department"]
