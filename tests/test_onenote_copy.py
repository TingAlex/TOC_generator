import base64
import unittest

from tocgen.onenote.client import ONE_NS, Page, Section
from tocgen.onenote.copy import (
    CopyValidationError,
    assert_pages_equal,
    page_signature,
    verify_resume_prefix,
)
from tocgen.onenote.native_copy import copy_native_page


def page_xml(page_id: str, title: str = "第一页", data: bytes = b"image-one",
             *, include_xps: bool = True, x: str = "10", y: str = "20",
             width: str = "300", height: str = "400") -> str:
    encoded = base64.b64encode(data).decode("ascii")
    xps = """
  <one:XPSFile xpsFileIndex="0" idDocument="{TEST-DOCUMENT}">
    <one:CallbackID callbackID="{TEST-CALLBACK}" />
  </one:XPSFile>""" if include_xps else ""
    return f'''<?xml version="1.0"?>
<one:Page xmlns:one="{ONE_NS}" ID="{page_id}">
  <one:QuickStyleDef index="1" name="PageTitle" />{xps}
  <one:PageSettings pageSize="2" />
  <one:Title><one:OE quickStyleIndex="1"><one:T>{title}</one:T></one:OE></one:Title>
  <one:Image format="png" isPrintOut="true" xpsFileIndex="0">
    <one:Position x="{x}" y="{y}" />
    <one:Size width="{width}" height="{height}" />
    <one:Data>{encoded}</one:Data>
  </one:Image>
</one:Page>'''


class PageSignatureTests(unittest.TestCase):
    def test_native_page_signature_keeps_xps_and_image_identity(self):
        source = page_signature(page_xml("source"))
        native_copy = page_signature(page_xml("destination"))

        self.assertEqual(source.xps_files, 1)
        self.assertEqual(source.callbacks, 1)
        assert_pages_equal(source, native_copy)

    def test_raster_only_page_is_rejected(self):
        with self.assertRaisesRegex(CopyValidationError, "固定栅格预览图"):
            page_signature(page_xml("raster", include_xps=False))

    def test_image_change_is_rejected_even_when_xps_exists(self):
        source = page_signature(page_xml("source"))
        changed = page_signature(page_xml("changed", data=b"different"))
        with self.assertRaises(CopyValidationError):
            assert_pages_equal(source, changed)


class FakeClient:
    def __init__(self):
        self.source_xml = page_xml("source")
        self.destination_xml: str | None = None
        self.destination_pages: list[Page] = []
        self.calls: list[tuple] = []
        self.deleted: list[str] = []

    def get_page_content(self, page_id, page_info=0):
        self.calls.append(("get", page_id, page_info))
        if page_id == "source":
            return self.source_xml
        if self.destination_xml is None:
            raise RuntimeError("destination is not ready")
        return self.destination_xml

    def list_section_pages(self, section_id):
        self.calls.append(("list", section_id))
        return list(self.destination_pages)

    def navigate_to(self, page_id):
        self.calls.append(("navigate", page_id))

    def sync_hierarchy(self, object_id):
        self.calls.append(("sync", object_id))

    def delete_page(self, page_id):
        self.calls.append(("delete", page_id))
        self.deleted.append(page_id)


class NativeLinearCopyTests(unittest.TestCase):
    def test_native_action_is_used_then_xps_is_verified_and_synced(self):
        client = FakeClient()
        action_calls = []

        def native_action(notebook, section, title, timeout):
            action_calls.append((notebook, section, title, timeout))
            client.destination_xml = page_xml("destination")
            client.destination_pages = [Page("destination", "第一页")]

        result = copy_native_page(
            client, Page("source", "第一页"), Section("section", "01"),
            "notebook-id", "在线笔记本", sync_settle=0, ready_timeout=0.1,
            copy_action=native_action,
        )

        self.assertEqual(result, "destination")
        self.assertEqual(action_calls, [("在线笔记本", "01", "第一页", 0.1)])
        self.assertIn(("navigate", "source"), client.calls)
        self.assertEqual(
            [call for call in client.calls if call[0] == "sync"],
            [("sync", "destination"),
             ("sync", "notebook-id"),
             ("sync", "notebook-id")],
        )
        self.assertEqual(client.deleted, [])

    def test_raster_result_is_recycled(self):
        client = FakeClient()

        def raster_action(notebook, section, title, timeout):
            client.destination_xml = page_xml("destination", include_xps=False)
            client.destination_pages = [Page("destination", "第一页")]

        with self.assertRaisesRegex(CopyValidationError, "原生整页复制"):
            copy_native_page(
                client, Page("source", "第一页"), Section("section", "01"),
                "notebook-id", "在线笔记本", sync_settle=0,
                ready_timeout=0.01, copy_action=raster_action,
            )
        self.assertEqual(client.deleted, ["destination"])
        self.assertEqual(client.calls[-1], ("sync", "notebook-id"))

    def test_resume_rejects_legacy_raster_prefix(self):
        client = FakeClient()
        client.destination_xml = page_xml("destination", include_xps=False)
        client.destination_pages = [Page("destination", "第一页")]
        source = Section("source-section", "01", [Page("source", "第一页")])
        destination = Section("destination-section", "01")
        signatures = {"source": page_signature(client.source_xml)}

        with self.assertRaisesRegex(CopyValidationError, "不是完整的 OneNote 原生副本"):
            verify_resume_prefix(client, source, destination, signatures)


if __name__ == "__main__":
    unittest.main()
