# ServiceNow Pipeline Generator

A Windows desktop and command-line utility that updates Azure DevOps pipeline YAML files. When a `jobs:` block contains a Cherwell task, the utility inserts a ServiceNow job or enriches an existing ServiceNow job with values from an Excel or CSV row.

The GUI creates a new output file and leaves the selected source pipeline unchanged. The command-line interface edits the YAML file passed to it, so copy the source first when you want to preserve it.

## Requirements

- Windows 10 or Windows 11 for the WinForms GUI
- Windows PowerShell 5.1
- Python 3.11 or later available as `python.exe` or through `py.exe -3`

The Python converter uses only the standard library. There is currently no `pip install` step and no web server, database, port, or environment variable to configure.

## Get the project

After cloning or downloading the repository, open PowerShell in the project directory:

```powershell
git clone https://github.com/Mohseen2912/ServiceNow-Automation.git
cd ServiceNow-Automation
```

If the project is already available locally:

```powershell
cd "E:\Automation Folder"
```

Confirm that Python works:

```powershell
python --version
```

If Windows resolves `python` to a non-working Microsoft Store alias, install Python with the **Add Python to PATH** option or prepend the installed Python directory for the current PowerShell session. For example:

```powershell
$env:Path = "C:\Path\To\Python311;$env:Path"
python --version
```

## Start the GUI

From the project directory:

```powershell
& ".\ServiceNow Pipeline Generator.cmd"
```

If you need to see PowerShell startup errors, launch the GUI script directly:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass `
  -File ".\scripts\servicenow_job_gui.ps1"
```

In the application:

1. Select an Azure pipeline `.yml` or `.yaml` file.
2. Select an `.xlsx`, `.xlsm`, or `.csv` input file.
3. Enter the worksheet name. The bundled sample uses `Changes`.
4. Select the data row. Row 1 contains headers, so the default data row is 2.
5. Choose a different output path from the source pipeline.
6. Select **Generate Pipeline**.

The bundled sample workbook is:

```text
outputs\servicenow_excel_sample\servicenow_changes_sample.xlsx
```

## Run from the command line

The CLI changes its YAML argument in place. This example first creates a working copy:

```powershell
Copy-Item `
  ".\sample_conditional_stages_pipeline.yml" `
  ".\tmp\sample_pipeline_run.yml"

python ".\scripts\servicenow_job_agent.py" `
  ".\tmp\sample_pipeline_run.yml" `
  --excel ".\outputs\servicenow_excel_sample\servicenow_changes_sample.xlsx" `
  --sheet Changes `
  --row 2
```

Run without Excel data when only the ServiceNow job skeleton is required:

```powershell
python ".\scripts\servicenow_job_agent.py" ".\path\to\pipeline.yml"
```

Preview whether a file would change without writing it:

```powershell
python ".\scripts\servicenow_job_agent.py" `
  ".\path\to\pipeline.yml" `
  --excel ".\path\to\changes.xlsx" `
  --sheet Changes `
  --row 2 `
  --dry-run
```

`--dry-run` returns exit code `1` when changes would be made and `0` when no changes are required. A `Would update` message therefore represents a successful preview, not a conversion error.

Show every CLI option:

```powershell
python ".\scripts\servicenow_job_agent.py" --help
```

## Excel and CSV input

The built-in ServiceNow fields recognize these header variations without case or punctuation sensitivity:

| ServiceNow field | Recognized headers |
|---|---|
| `authored_by` | Authored By, AuthoredBy, Author |
| `service_offering` | Service Offering, ServiceOffering |
| `cmdb_ci` | CMDB CI, CMDBCI, Configuration Item |
| `owned_by` | Owned By, Owner By, Owner |

Empty cells are omitted. Add another mapping from the CLI by repeating `--field`:

```powershell
python ".\scripts\servicenow_job_agent.py" `
  ".\path\to\pipeline.yml" `
  --excel ".\path\to\changes.xlsx" `
  --field "assignment_group=Assignment Group"
```

To regenerate the bundled sample workbook:

```powershell
powershell.exe -NoProfile `
  -File ".\scripts\create_sample_servicenow_excel.ps1"
```

## Run the tests

```powershell
python -m unittest discover -s ".\tests" -v
```

The current suite contains 14 tests covering job insertion, conditional stages and jobs, idempotency, CRLF preservation, field mappings, CSV input, and XLSX input.

## Project structure

```text
ServiceNow Pipeline Generator.cmd       GUI launcher
scripts/servicenow_job_gui.ps1          Windows desktop interface
scripts/servicenow_job_agent.py         YAML conversion and spreadsheet reader
scripts/create_sample_servicenow_excel.ps1
tests/test_servicenow_job_agent.py       Unit tests
outputs/servicenow_excel_sample/         Bundled sample workbook
sample_*.yml                             Example Azure pipelines
AZURE_DEVOPS_SERVICE_DEPLOYMENT.md       Proposed Azure DevOps deployment design
```

## Current YAML limitations

The current converter is line- and regular-expression-based rather than YAML-parser-based. Use conventional Azure DevOps YAML with spaces and two-space indentation.

Known limitations to address in a future parser-based update:

- Non-standard indentation beneath an existing `variables:` block can be rewritten incorrectly.
- Indentless sequences such as `jobs:` followed by an unindented `- job:` are not detected.
- Inline comments such as `jobs: # deployment jobs` are not recognized.
- Cherwell-looking task text inside a multiline script can be mistaken for a real task.
- Tabs are not supported for indentation.

Always review generated YAML before using it in an Azure DevOps pipeline.

## Azure DevOps deployment

The current project is a local Windows GUI/CLI utility. `AZURE_DEVOPS_SERVICE_DEPLOYMENT.md` describes a proposed browser-driven Azure DevOps service and future batch converter; those deployment components are not yet implemented in this repository.

## License

This project is licensed under the [MIT License](./LICENSE).