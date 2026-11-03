Sensitivity Labels
==================


A workbook can carry a Microsoft Purview (MIP) sensitivity label such as
"Public" or "Confidential". openpyxl can apply, read, change and remove the
label. Labels are preserved when a labelled workbook is loaded and saved.

openpyxl writes the label in both formats Office understands:

 * ``docMetadata/LabelInfo.xml``, read by Microsoft 365 and Office 2021 or later
 * ``MSIP_Label_<guid>_*`` custom document properties, read by older Office
   versions and some scanning tools (pass ``legacy_properties=False`` to skip)

.. warning::

    Only label *metadata* is written. openpyxl does not add the headers,
    footers or watermarks a label may define, and it cannot encrypt a
    workbook. Do not use it to apply labels that encrypt or restrict access.
    Policy checks, such as asking for a reason when a label is downgraded,
    are not run.


Describing your labels
----------------------

Each label is identified by a GUID within your organisation's Entra ID
tenant. Keep the definitions in a JSON file rather than in code:

.. code:: json

    {
      "site_id": "00000000-0000-0000-0000-000000000000",
      "labels": [
        {"name": "Public",       "id": "11111111-1111-1111-1111-111111111111"},
        {"name": "General",      "id": "22222222-2222-2222-2222-222222222222"},
        {"name": "Confidential", "id": "33333333-3333-3333-3333-333333333333"}
      ]
    }

``site_id`` is your tenant ID (Entra admin center, *Overview*). The label
GUIDs can be listed with Security & Compliance PowerShell::

    Connect-IPPSSession
    Get-Label | Select-Object DisplayName, Guid

A label can set its own ``site_id``. A label can also set ``content_bits``,
the bitmask of markings it defines (1 header, 2 footer, 4 watermark). Leave
``content_bits`` at 0 unless those markings are added another way. Label
names are matched case-insensitively.


Sample use
----------

Applying a label:

.. code::

    from openpyxl import load_workbook
    from openpyxl.packaging.sensitivity import LabelRegistry

    labels = LabelRegistry.from_json("labels.json")

    wb = load_workbook("report.xlsx")
    wb.set_sensitivity_label("Confidential", registry=labels)
    # or: wb.set_sensitivity_label(labels["Confidential"])
    wb.save("report.xlsx")

Setting a label replaces any label already on the workbook. ``method`` may
be ``"Standard"`` (the default) or ``"Privileged"``. Privileged marks the label
as a deliberate choice that automatic labelling should not override::

    wb.set_sensitivity_label("Public", registry=labels, method="Privileged")

Reading the label:

.. code::

    >>> label = wb.sensitivity_label      # None if the workbook has no label
    >>> label.id, label.siteId, label.method
    ('{33333333-3333-3333-3333-333333333333}', '{00000000-...}', 'Standard')
    >>> labels.lookup(label).name
    'Confidential'

``lookup`` returns ``None`` for a label that isn't in your registry, for
example one from another organisation.

Removing the label:

.. code::

    wb.remove_sensitivity_label()
    wb.save("report.xlsx")
