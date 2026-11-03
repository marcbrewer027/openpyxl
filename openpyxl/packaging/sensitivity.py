# Copyright (c) 2010-2024 openpyxl

"""
Microsoft Purview (MIP) sensitivity labels.

Office stores a workbook's sensitivity label in two places:

* ``docMetadata/LabelInfo.xml`` - the current format, read by Microsoft 365
  and Office 2021 or later.
* ``MSIP_Label_<guid>_*`` custom document properties - the legacy format,
  still read by older Office versions and some scanning / DLP tools.

Only label metadata is written. Content markings (headers, footers,
watermarks) and encryption defined by a label are *not* applied.
"""

import json
import uuid
from datetime import datetime, timezone

from openpyxl.descriptors import (
    Bool,
    Integer,
    Sequence,
    String,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.xml.constants import (
    ARC_LABEL_INFO,
    LABELINFO_NS,
    LABELINFO_TYPE,
)
from openpyxl.xml.functions import register_namespace, tostring

from .custom import StringProperty

register_namespace("clbl", LABELINFO_NS)

LABEL_METHODS = ("Standard", "Privileged")
LEGACY_PREFIX = "MSIP_Label_"


def _parse_guid(value, what="GUID"):
    """
    Return a ``uuid.UUID`` for value, which may be a UUID or a string with or
    without braces.
    """
    if isinstance(value, uuid.UUID):
        return value
    try:
        return uuid.UUID(str(value).strip())
    except (ValueError, AttributeError, TypeError):
        raise ValueError(f"{value!r} is not a valid {what}")


def _braced(value):
    """Format a GUID the way Office writes it in LabelInfo.xml"""
    return "{%s}" % _parse_guid(value)


class ClassificationLabel(Serialisable):
    """
    A ``<clbl:label>`` element of LabelInfo.xml
    """

    tagname = "label"
    namespace = LABELINFO_NS

    id = String()
    enabled = Bool()
    method = String(allow_none=True)
    siteId = String()
    contentBits = Integer(allow_none=True)
    removed = Bool()

    __elements__ = ()

    def __init__(self,
                 id=None,
                 enabled=True,
                 method="Standard",
                 siteId=None,
                 contentBits=0,
                 removed=False,
                ):
        self.id = id
        self.enabled = enabled
        self.method = method
        self.siteId = siteId
        self.contentBits = contentBits
        self.removed = removed


    @property
    def active(self):
        return bool(self.enabled) and not self.removed


    @property
    def guid(self):
        """Label id as a ``uuid.UUID``"""
        return _parse_guid(self.id)


    @property
    def site_guid(self):
        """Tenant (site) id as a ``uuid.UUID``"""
        return _parse_guid(self.siteId)


class ClassificationLabelList(Serialisable):
    """
    The ``<clbl:labelList>`` root element of LabelInfo.xml
    """

    tagname = "labelList"
    namespace = LABELINFO_NS
    path = "/" + ARC_LABEL_INFO
    mime_type = LABELINFO_TYPE

    label = Sequence(expected_type=ClassificationLabel)

    __elements__ = ("label",)

    def __init__(self, label=()):
        self.label = label


    def __len__(self):
        return len(self.label)


    @property
    def active(self):
        """The first label that is enabled and not removed, or None"""
        for label in self.label:
            if label.active:
                return label


    def _write(self):
        xml = tostring(self.to_tree())
        if not xml.startswith(b"<?xml"):
            xml = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml
        return xml


class SensitivityLabel:
    """
    Definition of one of the organisation's labels, as configured in Purview.

    :param name: display name, e.g. "Confidential"
    :param id: label GUID
    :param site_id: Entra ID tenant GUID
    :param content_bits: bitmask of content markings the label defines
        (1 header, 2 footer, 4 watermark). openpyxl does not add these
        markings, so leave this at 0 unless they are added another way.
    """

    def __init__(self, name, id, site_id, content_bits=0):
        if not isinstance(name, str) or not name.strip():
            raise ValueError("Label name must be a non-empty string")
        content_bits = int(content_bits)
        if content_bits < 0:
            raise ValueError("content_bits must not be negative")
        self.name = name.strip()
        self.id = _parse_guid(id, f"label id for {name!r}")
        self.site_id = _parse_guid(site_id, f"site id for {name!r}")
        self.content_bits = content_bits


    def __eq__(self, other):
        if not isinstance(other, SensitivityLabel):
            return NotImplemented
        return (self.name, self.id, self.site_id, self.content_bits) == (
            other.name, other.id, other.site_id, other.content_bits)


    def __hash__(self):
        return hash((self.name, self.id, self.site_id, self.content_bits))


    def __repr__(self):
        return f"SensitivityLabel(name={self.name!r}, id='{self.id}')"


class LabelRegistry:
    """
    The set of sensitivity labels available to the organisation.

    Usually loaded from a JSON file::

        {
          "site_id": "<tenant GUID>",
          "labels": [
            {"name": "Public", "id": "<label GUID>"},
            {"name": "Confidential", "id": "<label GUID>"}
          ]
        }

    Each label may override ``site_id`` and set ``content_bits``.
    Names are matched case-insensitively.
    """

    def __init__(self, labels=()):
        self._labels = []
        for label in labels:
            self.add(label)


    def add(self, label):
        if not isinstance(label, SensitivityLabel):
            raise TypeError("Expected a SensitivityLabel")
        key = label.name.casefold()
        for existing in self._labels:
            if existing.name.casefold() == key:
                raise ValueError(f"Duplicate label name {label.name!r}")
            if existing.id == label.id:
                raise ValueError(
                    f"Labels {existing.name!r} and {label.name!r} share id {label.id}")
        self._labels.append(label)


    @classmethod
    def from_dict(cls, data):
        if not isinstance(data, dict):
            raise ValueError("Label configuration must be a JSON object")
        labels = data.get("labels")
        if not isinstance(labels, list):
            raise ValueError("Label configuration needs a 'labels' list")
        default_site = data.get("site_id")
        registry = cls()
        for idx, entry in enumerate(labels):
            if not isinstance(entry, dict):
                raise ValueError(f"Label entry {idx} must be an object")
            missing = {"name", "id"} - set(entry)
            if missing:
                raise ValueError(
                    f"Label entry {idx} is missing {', '.join(sorted(missing))}")
            site_id = entry.get("site_id", default_site)
            if site_id is None:
                raise ValueError(
                    f"No site_id for label {entry['name']!r}: set a top-level "
                    "'site_id' or one on the label")
            registry.add(SensitivityLabel(
                name=entry["name"],
                id=entry["id"],
                site_id=site_id,
                content_bits=entry.get("content_bits", 0),
            ))
        return registry


    @classmethod
    def from_json(cls, path):
        with open(path, encoding="utf-8") as src:
            return cls.from_dict(json.load(src))


    def __getitem__(self, name):
        key = str(name).casefold()
        for label in self._labels:
            if label.name.casefold() == key:
                return label
        known = ", ".join(l.name for l in self._labels)
        raise KeyError(f"Unknown sensitivity label {name!r}. Known labels: {known}")


    def __contains__(self, name):
        try:
            self[name]
        except KeyError:
            return False
        return True


    def __iter__(self):
        return iter(self._labels)


    def __len__(self):
        return len(self._labels)


    @property
    def names(self):
        return [label.name for label in self._labels]


    def lookup(self, label):
        """
        Find the definition for a label GUID, a ``ClassificationLabel``
        (e.g. ``wb.sensitivity_label``) or ``None``.
        Returns ``None`` if the label is not in the registry.
        """
        if label is None:
            return None
        if isinstance(label, ClassificationLabel):
            label = label.id
        guid = _parse_guid(label)
        for defined in self._labels:
            if defined.id == guid:
                return defined


def _strip_legacy_props(custom_doc_props):
    """Remove every MSIP_Label_* custom document property"""
    for name in list(custom_doc_props.names):
        if name.startswith(LEGACY_PREFIX):
            del custom_doc_props[name]


def _legacy_props(label, method, set_date=None, action_id=None):
    """
    Build the legacy ``MSIP_Label_<guid>_*`` properties for a label
    """
    if set_date is None:
        set_date = datetime.now(timezone.utc)
    if action_id is None:
        action_id = uuid.uuid4()
    prefix = f"{LEGACY_PREFIX}{label.id}_"
    values = (
        ("Enabled", "true"),
        ("SetDate", set_date.strftime("%Y-%m-%dT%H:%M:%SZ")),
        ("Method", method),
        ("Name", label.name),
        ("SiteId", str(label.site_id)),
        ("ActionId", str(action_id)),
        ("ContentBits", str(label.content_bits)),
    )
    return [StringProperty(name=prefix + key, value=value) for key, value in values]


def apply_label(workbook, label, method="Standard", legacy_properties=True):
    """
    Replace any existing sensitivity label on the workbook with label
    """
    if method not in LABEL_METHODS:
        raise ValueError(f"method must be one of {LABEL_METHODS}")
    if not isinstance(label, SensitivityLabel):
        raise TypeError("Expected a SensitivityLabel")

    workbook.sensitivity_labels = ClassificationLabelList(label=[
        ClassificationLabel(
            id=_braced(label.id),
            enabled=True,
            method=method,
            siteId=_braced(label.site_id),
            contentBits=label.content_bits,
            removed=False,
        )
    ])

    _strip_legacy_props(workbook.custom_doc_props)
    if legacy_properties:
        for prop in _legacy_props(label, method):
            workbook.custom_doc_props.append(prop)


def remove_label(workbook):
    """Remove all sensitivity label metadata from the workbook"""
    workbook.sensitivity_labels = ClassificationLabelList()
    _strip_legacy_props(workbook.custom_doc_props)
