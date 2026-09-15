import base64
import unittest
import xml.etree.ElementTree as ET

from tocgen.onenote.client import ONE_NS, Page, Section
from tocgen.onenote.copy import (
    CopyValidationError,
    assert_images_equal,
    copy_printout_page,
    image_signatures,
    prepare_printout_page_xml,
    verify_resume_prefix,
)


def page_xml(page_id: str, title: str = "第一页", data: bytes = b"image-one",
             *, x: str = "10", y: str = "20", width: str = "300",
             height: str = "400", extra: str = "") -> str:
    encoded = base64.b64encode(data).decode("ascii")
    return f'''<?xml version="1.0"?>
<one:Page xmlns:one="{ONE_NS}" ID="{page_id}" objectID="root-object"
 lastModifiedTime="2026-01-01T00:00:00.000Z">
  <one:PageSettings objectID="settings-object" pageSize="2" />
  <one:QuickStyleDef index="1" name="PageTitle" fontColor="automatic" />
  <one:Title><one:OE objectID="title-object" quickStyleIndex="1">
    <one:T>{title}</one:T>
  </one:OE></one:Title>
  <one:XPSFile objectID="local-xps" />
  <one:CallbackID callbackID="local-callback" />
  <one:Image objectID="image-object" format="png" isPrintOut="true"
    xpsFileIndex="0" originalPageNumber="1">
    <one:Position x="{x}" y="{y}" />
    <one:Size width="{width}" height="{height}" />
    <one:Data>{encoded}</one:Data>
  </one:Image>
  {extra}
</one:Page>'''


class PreparePageXmlTests(unittest.TestCase):
    def test_keeps_self_contained_printout_and_strips_local_references(self):
        output = prepare_printout_page_xml(page_xml("source"), "destination")
        root = ET.fromstring(output)

        self.assertEqual(root.get("ID"), "destination")
        self.assertEqual(
            [child.tag.rsplit("}", 1)[-1] for child in root],
            ["PageSettings", "Title", "Image"],
        )
        attributes = [name.rsplit("}", 1)[-1]
                      for node in root.iter() for name in node.attrib]
        for forbidden in (
                "objectID", "lastModifiedTime", "quickStyleIndex", "isPrintOut",
                "xpsFileIndex", "originalPageNumber"):
            self.assertNotIn(forbidden, attributes)
        self.assertEqual(image_signatures(output), image_signatures(page_xml("source")))

    def test_rejects_content_that_would_be_silently_lost(self):
        extra = '<one:Outline><one:OEChildren /></one:Outline>'
        with self.assertRaisesRegex(CopyValidationError, "Outline"):
            prepare_printout_page_xml(page_xml("source", extra=extra), "destination")

    def test_image_comparison_allows_only_small_float_normalization(self):
        expected = image_signatures(page_xml("a", x="10.00", width="300.00"))
        normalized = image_signatures(page_xml("b", x="10.04", width="300.04"))
        assert_images_equal(expected, normalized)

        changed = image_signatures(page_xml("c", data=b"different"))
        with self.assertRaises(CopyValidationError):
            assert_images_equal(expected, changed)


class FakeClient:
    def __init__(self, *, corrupt_destination: bool = False):
        self.source_xml = page_xml("source")
        self.destination_xml = page_xml("destination", title="无标题页", data=b"blank")
        self.destination_name = "无标题页"
        self.corrupt_destination = corrupt_destination
        self.calls: list[tuple] = []
        self.deleted: list[str] = []

    def get_page_content(self, page_id, page_info=0):
        self.calls.append(("get", page_id, page_info))
        return self.source_xml if page_id == "source" else self.destination_xml

    def create_new_page(self, section_id, style=0):
        self.calls.append(("create", section_id, style))
        return "destination"

    def list_section_pages(self, section_id):
        self.calls.append(("list", section_id))
        return [Page("destination", self.destination_name)]

    def update_page_content(self, xml):
        self.calls.append(("update",))
        self.destination_xml = xml
        self.destination_name = "第一页"
        if self.corrupt_destination:
            self.destination_xml = self.destination_xml.replace(
                base64.b64encode(b"image-one").decode("ascii"),
                base64.b64encode(b"corrupted").decode("ascii"),
            )

    def sync_hierarchy(self, object_id):
        self.calls.append(("sync", object_id))

    def delete_page(self, page_id):
        self.calls.append(("delete", page_id))
        self.deleted.append(page_id)


class LinearCopyTests(unittest.TestCase):
    def test_copy_syncs_and_verifies_before_returning(self):
        client = FakeClient()
        source = Page("source", "第一页")
        destination = Section("section", "01")

        result = copy_printout_page(
            client, source, destination, "notebook",
            sync_settle=0, ready_timeout=0.1,
        )

        self.assertEqual(result, "destination")
        syncs = [call for call in client.calls if call[0] == "sync"]
        self.assertEqual(syncs, [
            ("sync", "destination"),
            ("sync", "notebook"),
            ("sync", "notebook"),
        ])
        self.assertEqual(client.deleted, [])

    def test_failed_verification_recycles_only_new_page(self):
        client = FakeClient(corrupt_destination=True)
        with self.assertRaisesRegex(CopyValidationError, "已中止"):
            copy_printout_page(
                client, Page("source", "第一页"), Section("section", "01"),
                "notebook", sync_settle=0, ready_timeout=0.01,
            )
        self.assertEqual(client.deleted, ["destination"])
        self.assertEqual(client.calls[-1], ("sync", "notebook"))

    def test_resume_requires_matching_title_and_image_prefix(self):
        client = FakeClient()
        client.destination_xml = page_xml("destination")
        client.destination_name = "第一页"
        source = Section("source-section", "01", [Page("source", "第一页")])
        destination = Section("destination-section", "01")
        signatures = {"source": image_signatures(client.source_xml)}

        self.assertEqual(
            verify_resume_prefix(client, source, destination, signatures), 1)

        client.destination_name = "错位页"
        with self.assertRaisesRegex(CopyValidationError, "不构成续跑前缀"):
            verify_resume_prefix(client, source, destination, signatures)


if __name__ == "__main__":
    unittest.main()
