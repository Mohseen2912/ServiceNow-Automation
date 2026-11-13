import tempfile
import unittest
import zipfile
from pathlib import Path

from scripts.servicenow_job_agent import insert_servicenow_job, read_excel_row


class ServicenowJobAgentTests(unittest.TestCase):
    def test_inserts_job_before_existing_jobs_when_cherwell_task_is_present(self):
        original = """stages:
  - stage: Build
    jobs:
      - job: unit-tests
        steps:
          - script: echo tests
          - task: CherwellChange@1
            inputs:
              changeId: CHG123
      - job: deploy
        steps:
          - script: echo deploy
"""

        updated = insert_servicenow_job(original)

        self.assertIn("- job: servicenow", updated)
        self.assertLess(updated.index("- job: servicenow"), updated.index("- job: unit-tests"))
        self.assertLess(updated.index("- job: servicenow"), updated.index("- task: CherwellChange@1"))

    def test_leaves_content_unchanged_when_no_cherwell_job_is_present(self):
        original = """stages:
  - stage: Build
    jobs:
      - job: unit-tests
        steps:
          - script: echo tests
"""

        self.assertEqual(insert_servicenow_job(original), original)

    def test_does_not_insert_duplicate_job_when_already_present(self):
        original = """stages:
  - stage: Build
    jobs:
      - job: servicenow
        steps:
          - script: echo done
      - job: deploy
        steps:
          - task: cherwell
"""

        updated = insert_servicenow_job(original)
        self.assertEqual(updated.count("- job: servicenow"), 1)

    def test_inserts_job_in_each_stage_with_cherwell_task(self):
        original = """stages:
  - stage: Build
    jobs:
      - job: validate
        steps:
          - task: cherwell
  - stage: Deploy
    jobs:
      - job: unit-tests
        steps:
          - script: echo tests
      - job: deploy
        steps:
          - task: CherwellChange@1
"""

        updated = insert_servicenow_job(original)
        self.assertEqual(updated.count("- job: servicenow"), 2)

    def test_detects_task_name_containing_cherwell(self):
        original = """stages:
  - stage: Build
    jobs:
      - job: validate
        steps:
          - task: Company.Cherwell.Change@1
"""

        updated = insert_servicenow_job(original)

        self.assertIn("- job: servicenow", updated)

    def test_detects_cherwell_task_inside_conditional_steps(self):
        original = """stages:
  - stage: Build
    jobs:
      - job: validate
        steps:
          - ${{ if eq(variables['CreateChange'], 'true') }}:
            - task: Company.Cherwell.Change@1
"""

        updated = insert_servicenow_job(original)

        self.assertIn("- job: servicenow", updated)
        self.assertLess(updated.index("- job: servicenow"), updated.index("- job: validate"))

    def test_inserts_before_conditional_job_wrapper(self):
        original = """stages:
  - stage: Prod
    jobs:
      - ${{ if eq(variables['DeployProd'], 'true') }}:
        - job: deploy_prod
          steps:
            - task: Company.Cherwell.Change@1
      - ${{ else }}:
        - job: skip_prod
          steps:
            - script: echo Skipping prod
"""

        updated = insert_servicenow_job(original)

        self.assertIn("      - job: servicenow\n", updated)
        self.assertLess(
            updated.index("      - job: servicenow"),
            updated.index("      - ${{ if eq(variables['DeployProd'], 'true') }}:"),
        )
        self.assertNotIn("        - job: servicenow", updated)

    def test_inserts_inside_conditional_stage_with_cherwell_task(self):
        original = """stages:
  - ${{ if eq(variables['DeployProd'], 'true') }}:
    - stage: prod
      jobs:
        - job: deploy_prod
          steps:
            - task: Company.Cherwell.Change@1
  - ${{ else }}:
    - stage: prod_skipped
      jobs:
        - job: skip_prod
          steps:
            - script: echo Skipping prod
"""

        updated = insert_servicenow_job(original)

        self.assertIn("        - job: servicenow\n", updated)
        self.assertLess(
            updated.index("        - job: servicenow"),
            updated.index("        - job: deploy_prod"),
        )
        self.assertEqual(updated.count("- job: servicenow"), 1)

    def test_preserves_crlf_line_endings(self):
        original = (
            "stages:\r\n"
            "  - stage: Build\r\n"
            "    jobs:\r\n"
            "      - job: validate\r\n"
            "        steps:\r\n"
            "          - task: cherwell\r\n"
        )

        updated = insert_servicenow_job(original)

        self.assertIn("      - job: servicenow\r\n", updated)
        self.assertNotIn("\n      - job: servicenow\n", updated)

    def test_adds_servicenow_fields_from_excel_row(self):
        original = """stages:
  - stage: Prod
    jobs:
      - job: deploy
        steps:
          - task: CherwellChange@1
"""
        excel_row = {
            "Authored By": "Jane O'Brien",
            "Service Offering": "Payments",
            "CMDB CI": "payments-api-prod",
            "Owner By": "Platform Operations",
            "Unrelated Column": "must not be copied",
        }

        updated = insert_servicenow_job(original, excel_row)

        self.assertIn("        authored_by: 'Jane O''Brien'", updated)
        self.assertIn("        service_offering: 'Payments'", updated)
        self.assertIn("        cmdb_ci: 'payments-api-prod'", updated)
        self.assertIn("        owned_by: 'Platform Operations'", updated)
        self.assertNotIn("Unrelated Column", updated)

    def test_enriches_existing_servicenow_job_without_duplication(self):
        original = """stages:
  - stage: Prod
    jobs:
      - job: servicenow
        variables:
          keep_me: existing
          cmdb_ci: old-ci
        steps:
          - script: echo ServiceNow job
      - job: deploy
        steps:
          - task: cherwell
"""
        excel_row = {"CMDB_CI": "new-ci", "Owned By": "Cloud Team"}

        updated = insert_servicenow_job(original, excel_row)
        updated_again = insert_servicenow_job(updated, excel_row)

        self.assertEqual(updated.count("- job: servicenow"), 1)
        self.assertEqual(updated.count("cmdb_ci: 'new-ci'"), 1)
        self.assertIn("keep_me: existing", updated)
        self.assertEqual(updated_again, updated)

    def test_supports_additional_field_mapping(self):
        original = """stages:
  - stage: Prod
    jobs:
      - job: deploy
        steps:
          - task: cherwell
"""

        updated = insert_servicenow_job(
            original,
            {"Assignment Group": "Change Management"},
            field_aliases={"assignment_group": ("Assignment Group",)},
        )

        self.assertIn("assignment_group: 'Change Management'", updated)

    def test_reads_selected_csv_row_using_headers(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            csv_path = Path(temporary_directory) / "changes.csv"
            csv_path.write_text(
                "Authored By,Service Offering,CMDB CI,Owned By\n"
                "Jane,Payments,ci-one,Team One\n"
                "John,Orders,ci-two,Team Two\n",
                encoding="utf-8",
            )

            row = read_excel_row(csv_path, row_number=3)

        self.assertEqual(row["Authored By"], "John")
        self.assertEqual(row["CMDB CI"], "ci-two")

    def test_reads_selected_xlsx_row_using_headers(self):
        workbook_xml = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Changes" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""
        relationships_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
        shared_strings_xml = """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Authored By</t></si><si><t>CMDB CI</t></si>
  <si><t>Jane</t></si><si><t>ci-one</t></si>
  <si><t>John</t></si><si><t>ci-two</t></si>
</sst>"""
        worksheet_xml = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
    <row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" t="s"><v>3</v></c></row>
    <row r="3"><c r="A3" t="s"><v>4</v></c><c r="B3" t="s"><v>5</v></c></row>
  </sheetData>
</worksheet>"""

        with tempfile.TemporaryDirectory() as temporary_directory:
            xlsx_path = Path(temporary_directory) / "changes.xlsx"
            with zipfile.ZipFile(xlsx_path, "w") as archive:
                archive.writestr("xl/workbook.xml", workbook_xml)
                archive.writestr("xl/_rels/workbook.xml.rels", relationships_xml)
                archive.writestr("xl/sharedStrings.xml", shared_strings_xml)
                archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml)

            row = read_excel_row(
                xlsx_path,
                row_number=3,
                sheet_name="Changes",
            )

        self.assertEqual(row, {"Authored By": "John", "CMDB CI": "ci-two"})


if __name__ == "__main__":
    unittest.main()
