# ServiceNow Pipeline Converter on Azure DevOps Services

## 1. Purpose

This document describes how to operate the ServiceNow pipeline converter entirely through Azure DevOps Services when users work on a locked-down VDI and cannot bypass PowerShell execution policy.

The intended workflow is:

1. A user receives an offline Excel workbook.
2. Each data row contains ServiceNow values and a `pipeline_url` column. (We can change the column name later as per the excel)
3. The user uploads the workbook through the Azure DevOps web interface.
4. An Azure Pipeline reads the workbook.
5. For each row, the pipeline securely downloads the pipeline YAML identified by `pipeline_url` column of excel.
6. The existing Python conversion logic inserts or enriches the ServiceNow job.
7. The pipeline publishes the converted YAML files and a processing report as an Azure Pipeline artifact.

The first release does **not** update target repositories. It only reads source YAML and publishes converted copies for review. Automatic branches and pull requests are described as a later phase.

## 2. Important terminology

Azure DevOps Services is not hosting the existing Windows desktop GUI. Instead, the Azure Pipeline is the application entry point:

```text
VDI browser
    |
    | Upload workbook and select Run pipeline
    v
Azure Repos intake repository
    |
    v
Microsoft-hosted Azure Pipeline agent
    |-- Read workbook
    |-- Validate pipeline_url
    |-- Download pipeline YAML with System.AccessToken (IT IS NOT RECOMMENDED TO USE PAT token)
    |-- Run Python converter
    |-- Write results and report
    v
Azure Pipeline artifact
    |
    v
VDI browser download
```

No PowerShell script runs on the VDI. No Python installation is required on the VDI.

## 3. Decisions for the first release

Use these defaults unless the security or platform teams require different settings:

| Decision | First-release choice |
|---|---|
| Azure DevOps location | One Azure DevOps organization |
| Converter repository | Dedicated Azure Repo |
| Workbook intake | Browser upload to `requests/incoming/` |
| Run method | Manual **Run pipeline** with the workbook path |
| Build agent | Microsoft-hosted Ubuntu agent |
| Authentication | `System.AccessToken` |
| PAT | Not used |
| Source access | Read-only |
| Output | Pipeline artifact |
| Repository changes | None |
| Pull requests | Deferred to phase 2 |

Manual pipeline execution is recommended initially because it makes the selected workbook explicit and avoids ambiguity when several people upload workbooks at nearly the same time.

## 4. Prerequisites

### 4.1 Azure DevOps prerequisites

An Azure DevOps administrator or project administrator must provide:

- An Azure DevOps Services organization and project.
- Azure Repos enabled for the project.
- Azure Pipelines enabled for the project.
- Permission to create the converter repository and pipeline.
- A Microsoft-hosted agent entitlement, or an approved self-hosted agent with Python installed.
- Read permission for the converter pipeline's Build Service identity on every repository that may appear in `pipeline_url`.

### 4.2 User prerequisites

Workbook submitters need:

- Browser access to Azure DevOps from the VDI.
- Read and Contribute permission on the converter intake repository.
- Permission to run the converter pipeline.
- Permission to download artifacts from pipeline runs.

They do not need PowerShell, Python, a PAT, or administrator rights on the VDI.

### 4.3 Workbook prerequisites

Supported file formats are:

- `.xlsx`
- `.xlsm` (macros are not executed)
- `.csv`

Legacy `.xls` files are not supported by the current converter.

The workbook should contain literal cell values. It must not depend on macros or formulas being recalculated by the pipeline agent.

## 5. Workbook contract

### 5.1 Required worksheet and headers

Use a worksheet named `Changes`. Row 1 contains headers and data begins on row 2.

Required header:

- `pipeline_url`

ServiceNow headers currently recognized by the converter:

- `authored_by` or `Authored By`
- `service_offering` or `Service Offering`
- `cmdb_ci` or `CMDB CI`
- `owned_by` or `Owned By`

Header matching for ServiceNow fields is case-insensitive and ignores spaces, underscores, and punctuation.

Example:

| pipeline_url | authored_by | service_offering | cmdb_ci | owned_by |
|---|---|---|---|---|
| `https://dev.azure.com/contoso/Payments/_git/Payments.Api?path=/azure-pipelines.yml&version=GBmain` | Jane Smith | Payments | payments-api-prod | Cloud Operations |
| `https://dev.azure.com/contoso/Orders/_git/Orders.Api?path=/pipelines/prod.yml&version=GBrelease/2026.07` | John Smith | Orders | orders-api-prod | Platform Operations |

### 5.2 Accepted `pipeline_url` format

For the first release, accept only Azure Repos file URLs in this form:

```text
https://dev.azure.com/{organization}/{project}/_git/{repository}?path=/{yaml-path}&version=GB{branch}
```

Example:

```text
https://dev.azure.com/contoso/Payments/_git/Payments.Api?path=/pipelines/deploy.yml&version=GBmain
```

The implementation must URL-decode the project, repository, path, and version values. A copied Azure Repos browser URL may contain encoded characters such as `%20` and `%2F`.

### 5.3 URL validation rules

The orchestration code must validate every URL before making a network request:

1. Scheme must be `https`.
2. Host must be exactly `dev.azure.com`.
3. Organization must match an configured allowlist.
4. Project may optionally be restricted by a configured allowlist.
5. URL path must match `/{organization}/{project}/_git/{repository}`.
6. The `path` query value must be present and end in `.yml` or `.yaml`.
7. The decoded YAML path must be repository-relative and must not contain traversal segments such as `..`.
8. The URL must not contain credentials.
9. The URL must not redirect to an unapproved host.
10. Duplicate URLs should be processed once per distinct row because ServiceNow values may differ between rows.

Do not allow arbitrary HTTP or HTTPS URLs. Blindly requesting URLs supplied through Excel creates a server-side request forgery risk.

### 5.4 Workbook processing rules

- Skip completely blank rows.
- Fail a row when `pipeline_url` is blank or invalid.
- Continue processing the remaining rows after a row-level failure.
- Set a configurable maximum number of data rows, for example 500.
- Generate one output file per successful row.
- Never modify the uploaded workbook.
- Never modify the downloaded source YAML in place; work on an output copy.
- Preserve the original YAML filename in the output name where practical.

Suggested output name:

```text
{excel-row}_{project}_{repository}_{source-name}_servicenow.yml
```

## 6. Repository structure

Create or update the converter repository to use the following structure:

```text
/
|-- azure-pipelines.yml
|-- scripts/
|   |-- __init__.py
|   |-- servicenow_job_agent.py
|   `-- azure_devops_batch_converter.py   # new orchestration script
|-- tests/
|   |-- test_servicenow_job_agent.py
|   `-- test_azure_devops_batch_converter.py
|-- requests/
|   `-- incoming/
|       `-- .gitkeep
|-- samples/
|   `-- service_now_request_sample.xlsx
`-- docs/
    `-- operating-guide.md
```

The existing `servicenow_job_agent.py` remains responsible for transforming YAML. The new `azure_devops_batch_converter.py` is responsible for Azure DevOps integration and batch processing.

## 7. Required code changes

The current converter supports one local YAML and one selected workbook row. It does not yet read `pipeline_url` or process an entire workbook. The following work is required.

### 7.1 Add a workbook-row iterator

Add a function that returns every populated row rather than only one selected row. It should:

- Support `.xlsx`, `.xlsm`, and `.csv` consistently with the existing reader.
- Return the physical Excel row number with each row.
- Preserve headers as strings.
- Detect a missing or duplicate `pipeline_url` header.
- Enforce the configured maximum row count.
- Produce clear errors for unreadable or malformed workbooks.

Suggested interface:

```python
def iter_excel_rows(
    path: Path,
    *,
    header_row: int = 1,
    sheet_name: str = "Changes",
) -> Iterable[tuple[int, dict[str, str]]]:
    ...
```

### 7.2 Add an Azure Repos URL parser

Parse `pipeline_url` with `urllib.parse`; do not use a regular expression alone for the complete URL.

Suggested result type:

```python
@dataclass(frozen=True)
class AzureRepoFile:
    organization: str
    project: str
    repository: str
    path: str
    version: str | None
    version_type: str
```

For Azure Repos browser URLs:

- `version=GBmain` means branch `main`.
- A missing version can use the repository's default branch, but the resolved commit must be written to the report.
- If support for tags or commits is added, handle them explicitly rather than treating every version as a branch.

### 7.3 Download YAML through the Git Items REST API

Convert the browser URL into a REST request like:

```text
GET https://dev.azure.com/{organization}/{project}/_apis/git/repositories/{repository}/items
    ?path={yaml-path}
    &versionDescriptor.version={branch}
    &versionDescriptor.versionType=branch
    &download=true
    &api-version=7.1
```

Send the job access token in the HTTP header:

```text
Authorization: Bearer {System.AccessToken}
```

The token must come from the `SYSTEM_ACCESSTOKEN` environment variable. It must never come from the workbook, a command-line argument, or a committed configuration file.

The downloader should:

- Use a reasonable connection and response timeout.
- Limit the maximum downloaded file size.
- Reject non-YAML paths before download.
- Disable or validate redirects.
- Distinguish authentication, authorization, not-found, throttling, and server errors.
- Retry only transient failures such as HTTP 429 and selected 5xx responses.
- Never print the access token.

### 7.4 Convert each downloaded YAML

For each valid row:

1. Download the source YAML into a job-scoped temporary directory.
2. Call `insert_servicenow_job(source_text, excel_row)` from the existing module.
3. If no Cherwell task exists, mark the row `NO_CHANGE`.
4. If conversion succeeds, write the converted YAML to the artifact staging directory.
5. Record the source URL, row number, output filename, resolved source version, and status.

Do not invoke the PowerShell GUI and do not call `ServiceNow Pipeline Generator.cmd` from Azure Pipelines.

### 7.5 Generate a processing report

Publish both a human-readable CSV report and a machine-readable JSON report.

Suggested report columns:

| Column | Description |
|---|---|
| `excel_row` | Original workbook row |
| `pipeline_url` | Source URL, excluding any credentials |
| `organization` | Parsed organization |
| `project` | Parsed project |
| `repository` | Parsed repository |
| `yaml_path` | Parsed source path |
| `source_version` | Requested branch, tag, or commit |
| `resolved_commit` | Exact commit retrieved |
| `status` | `CONVERTED`, `NO_CHANGE`, or `FAILED` |
| `output_file` | Artifact-relative output path |
| `message` | Sanitized error or informational message |
| `sha256` | SHA-256 of generated output |

Do not place access tokens, authorization headers, or secret-variable values in the report.

### 7.6 Define exit behavior

Recommended behavior:

- Exit `0` when all rows are `CONVERTED` or `NO_CHANGE`.
- Exit nonzero when any row is `FAILED`.
- Always write the report before exiting.
- Configure artifact publishing with `condition: always()` so the report is available after partial failures.

## 8. Authentication and authorization

### 8.1 Do not use a PAT for same-organization access

Use Azure Pipelines' `System.AccessToken`, which is a short-lived job access token. Map it only into the orchestration step:

```yaml
env:
  SYSTEM_ACCESSTOKEN: $(System.AccessToken)
```

The token's effective permissions come from the pipeline job's authorization scope and the Build Service identity's permissions.

### 8.2 Grant read permission to target repositories

For each target repository:

1. Open **Project settings**.
2. Open **Repositories**.
3. Select the target repository.
4. Open **Security**.
5. Locate or add the converter project's Build Service identity. It is normally named similar to:

   ```text
   {Converter Project Name} Build Service ({Organization Name})
   ```

6. Set **Read** to **Allow**.
7. Do not grant Contribute, Force push, Delete repository, or policy-bypass permissions for the first release.

If target repositories are in other projects in the same organization, an administrator may also need to add the Build Service identity to those projects or explicitly grant repository access. Keep job authorization scope limited to the current project where possible, then grant only the individual repository permissions required.

### 8.3 Cross-organization repositories

Do not solve cross-organization access by adding a PAT to the workbook. Use an approved Azure DevOps service connection and a separately reviewed authentication design. Cross-organization support should be treated as a separate release because the trust boundary and permissions differ.

## 9. Azure Pipeline definition

After `azure_devops_batch_converter.py` is implemented, use a pipeline similar to the following:

```yaml
trigger: none
pr: none

parameters:
  - name: inputWorkbook
    displayName: Workbook path in converter repository
    type: string
    default: requests/incoming/request.xlsx

  - name: worksheet
    displayName: Worksheet name
    type: string
    default: Changes

pool:
  vmImage: ubuntu-latest

steps:
  - checkout: self
    clean: true

  - task: UsePythonVersion@0
    displayName: Select Python
    inputs:
      versionSpec: "3.12"

  - bash: |
      set -euo pipefail
      python -m unittest discover -s tests -v
    displayName: Run converter tests

  - bash: |
      set -euo pipefail
      mkdir -p "$OUTPUT_DIRECTORY"

      python scripts/azure_devops_batch_converter.py \
        --excel "$BUILD_SOURCESDIRECTORY/$INPUT_WORKBOOK" \
        --sheet "$WORKSHEET" \
        --output "$OUTPUT_DIRECTORY" \
        --allowed-organization "$ALLOWED_ORGANIZATION"
    displayName: Download and convert pipelines
    env:
      SYSTEM_ACCESSTOKEN: $(System.AccessToken)
      INPUT_WORKBOOK: ${{ parameters.inputWorkbook }}
      WORKSHEET: ${{ parameters.worksheet }}
      ALLOWED_ORGANIZATION: contoso
      OUTPUT_DIRECTORY: $(Build.ArtifactStagingDirectory)/servicenow-conversion

  - task: PublishPipelineArtifact@1
    displayName: Publish conversion results
    condition: always()
    inputs:
      targetPath: $(Build.ArtifactStagingDirectory)/servicenow-conversion
      artifact: servicenow-conversion
```

Runtime parameters are passed through environment variables rather than embedded directly into Bash commands. The orchestration script must still ensure that the workbook resolves inside `$(Build.SourcesDirectory)` and does not use an absolute path or `..` traversal.

## 10. Create the Azure DevOps pipeline

An Azure DevOps project administrator should:

1. Push the converter repository to Azure Repos.
2. Open **Pipelines** and select **New pipeline**.
3. Select **Azure Repos Git**.
4. Select the converter repository.
5. Select **Existing Azure Pipelines YAML file**.
6. Select `/azure-pipelines.yml` from the default branch.
7. Save the pipeline without running it until repository permissions are configured.
8. Rename the pipeline to a recognizable name, for example `ServiceNow Pipeline Converter`.
9. Configure who can view and run the pipeline.
10. Confirm that secret values are not exposed through queue-time variables.

For YAML pipelines, `System.AccessToken` is made available by explicitly mapping it into the relevant task's environment, as shown above.

## 11. End-user operating procedure

### 11.1 Prepare the workbook

1. Confirm the workbook is `.xlsx`, `.xlsm`, or `.csv`.
2. Confirm the worksheet is named `Changes`.
3. Confirm row 1 contains `pipeline_url` and the required ServiceNow headers.
4. Confirm every `pipeline_url` opens the desired YAML file when viewed by the user in Azure Repos.
5. Save and close the workbook.
6. Give it a unique request name, for example:

   ```text
   CHG0123456_2026-07-21.xlsx
   ```

### 11.2 Upload through the VDI browser

1. Open the converter repository in Azure DevOps.
2. Browse to `requests/incoming/`.
3. Select the upload-files action.
4. Select the workbook from the VDI.
5. Commit it to a request branch or directly to the intake branch according to the team's repository policy.
6. Copy the repository-relative path, for example:

   ```text
   requests/incoming/CHG0123456_2026-07-21.xlsx
   ```

If workbook contents must not be retained in Git history, use the secure-file intake alternative in section 14 instead.

### 11.3 Run the pipeline

1. Open **Pipelines**.
2. Select **ServiceNow Pipeline Converter**.
3. Select **Run pipeline**.
4. Select the branch containing the uploaded workbook.
5. Enter the workbook path in **Workbook path in converter repository**.
6. Confirm the worksheet name.
7. Select **Run**.

### 11.4 Review the result

1. Open the completed run.
2. Review test and conversion logs.
3. Open the `servicenow-conversion` artifact.
4. Download the artifact.
5. Review `conversion-report.csv` first.
6. Review each generated YAML before using it in a target repository.
7. Raise a normal pull request manually for accepted output during phase 1.

## 12. Expected failure handling

| Failure | Expected result |
|---|---|
| Workbook missing | Run fails before processing |
| Workbook path outside repository | Run fails validation |
| Worksheet missing | Run fails with a clear workbook error |
| `pipeline_url` header missing | Run fails workbook validation |
| Invalid URL in one row | Row is `FAILED`; other rows continue |
| URL outside allowlisted organization | Row is `FAILED`; no request is sent |
| Build identity lacks repository Read | Row is `FAILED` with sanitized 401/403 message |
| YAML path does not exist | Row is `FAILED` with 404 message |
| File is not `.yml` or `.yaml` | Row is `FAILED` before download |
| YAML contains no Cherwell task | Row is `NO_CHANGE` |
| Conversion succeeds | Row is `CONVERTED` and output is published |
| Some rows fail | Run fails, but report and successful outputs are still published |

## 13. Test plan

### 13.1 Unit tests

Add tests for:

- Valid Azure Repos URLs.
- URL-encoded project, repository, branch, and file names.
- Missing `path` query parameter.
- Non-YAML paths.
- HTTP instead of HTTPS.
- Incorrect hosts such as `dev.azure.com.example.com`.
- Disallowed organizations and projects.
- Embedded credentials.
- Traversal attempts.
- Branch parsing.
- Blank workbook rows.
- Multiple workbook rows.
- Missing and duplicate headers.
- Per-row conversion and output naming.
- Sanitization of errors and reports.
- Partial success behavior.

Mock HTTP responses in unit tests; tests must not require a real Azure DevOps token.

### 13.2 Integration tests

Create a non-production Azure Repo containing:

- A YAML file with a Cherwell task.
- A YAML file without a Cherwell task.
- A conditional-stage example.
- A restricted file or repository used to verify authorization failure.

Run a workbook that references these files and verify the artifact report and outputs.

### 13.3 Security tests

Verify rejection of:

- `http://` URLs.
- Hosts outside `dev.azure.com`.
- Loopback and private-network targets.
- URL user information such as `https://user:secret@...`.
- Redirects to another host.
- Paths containing `..` after decoding.
- Oversized YAML responses.
- Workbooks exceeding the row limit.
- Formula or malformed-cell edge cases.
- Attempts to expose `SYSTEM_ACCESSTOKEN` in logs.

## 14. Secure-file intake alternative

If the workbook contains information that must not be stored in Git history:

1. Upload it under **Pipelines > Library > Secure files**.
2. Authorize only the converter pipeline and approved users.
3. Pass the approved secure-file name or identifier to `DownloadSecureFile@1`.
4. Use `$(taskName.secureFilePath)` as the orchestration script's input.
5. Let Azure Pipelines delete the temporary agent copy at the end of the job.

Example:

```yaml
- task: DownloadSecureFile@1
  name: requestWorkbook
  displayName: Download request workbook
  inputs:
    secureFile: ${{ parameters.secureWorkbook }}

- bash: |
    set -euo pipefail
    python scripts/azure_devops_batch_converter.py \
      --excel "$REQUEST_WORKBOOK" \
      --sheet "$WORKSHEET" \
      --output "$OUTPUT_DIRECTORY" \
      --allowed-organization "$ALLOWED_ORGANIZATION"
  env:
    SYSTEM_ACCESSTOKEN: $(System.AccessToken)
    REQUEST_WORKBOOK: $(requestWorkbook.secureFilePath)
    WORKSHEET: ${{ parameters.worksheet }}
    ALLOWED_ORGANIZATION: contoso
    OUTPUT_DIRECTORY: $(Build.ArtifactStagingDirectory)/servicenow-conversion
```

Secure Files requires an upload to the Library before the run; the **Run pipeline** screen is not a general-purpose file attachment form.

## 15. Phase 2: automatic pull requests

Only implement this after the read-only artifact workflow is accepted.

Additional requirements include:

- A controlled branch naming convention.
- Build Service permission to create branches and contribute changes.
- Permission to create pull requests.
- Branch policies requiring human review.
- Collision handling when multiple workbook rows target the same file.
- Idempotency when a request is rerun.
- A decision about whether one workbook produces one PR per repository or one PR per row.
- Audit linkage between workbook request, pipeline run, commit, and PR.

Do not grant force-push, branch-policy bypass, or direct-write permission to protected branches.

## 16. Monitoring, retention, and audit

For every run, retain or record:

- Azure DevOps run ID.
- Submitter identity.
- Source workbook name and commit.
- Workbook SHA-256.
- Source repository and resolved commit for each row.
- Converter repository commit.
- Output hashes.
- Conversion report.
- Final run status.

Agree with security and compliance teams on artifact, log, and workbook retention. Avoid printing full workbook rows in logs because they may contain operational data.

## 17. Rollout plan

### Stage 1: development

- Implement batch workbook reading, URL validation, REST downloading, and reports.
- Add unit tests.
- Run only against a dedicated test repository.

### Stage 2: controlled pilot

- Grant read permission to a small repository allowlist.
- Have selected users upload test workbooks.
- Publish artifacts only.
- Compare converted output with manual conversion.

### Stage 3: production read-only service

- Expand the approved repository list.
- Apply retention, monitoring, and support procedures.
- Keep output artifact-only.

### Stage 4: optional pull-request automation

- Complete a separate permission and security review.
- Add branch and PR creation.
- Require human reviewers through branch policies.

## 18. Acceptance criteria for the first release

The release is ready when all of the following are true:

- A VDI user can upload a workbook and run the process using only a browser.
- No PowerShell execution-policy bypass is used.
- No PAT is present in workbooks, source code, pipeline YAML, logs, or artifacts.
- Only allowlisted Azure DevOps URLs can be requested.
- The pipeline identity has read-only access to approved repositories.
- Every populated workbook row produces a deterministic status.
- Successful converted YAML files are published as an artifact.
- Partial failures still produce a report.
- Existing converter and orchestration tests pass.
- Security testing confirms that access tokens are not logged.
- Operations documentation identifies the support owner and escalation path.

## 19. Microsoft references

- Azure DevOps Git Items REST API: <https://learn.microsoft.com/en-us/rest/api/azure/devops/git/items/get?view=azure-devops-rest-7.1>
- `System.AccessToken` predefined variable: <https://learn.microsoft.com/en-us/azure/devops/pipelines/build/variables?view=azure-devops>
- Job access tokens and authorization scope: <https://learn.microsoft.com/en-us/azure/devops/pipelines/process/access-tokens?view=azure-devops>
- Use Python Version task: <https://learn.microsoft.com/en-us/azure/devops/pipelines/tasks/reference/use-python-version-v0?view=azure-pipelines>
- Publish Pipeline Artifact task reference: <https://learn.microsoft.com/en-us/azure/devops/pipelines/tasks/reference/publish-pipeline-artifact-v1?view=azure-pipelines>
- Download Secure File task reference: <https://learn.microsoft.com/en-us/azure/devops/pipelines/tasks/reference/download-secure-file-v1?view=azure-pipelines>
- Multi-repository checkout limitations: <https://learn.microsoft.com/en-us/azure/devops/pipelines/repos/multi-repo-checkout?view=azure-devops>
- Azure Repos pull requests: <https://learn.microsoft.com/en-us/azure/devops/repos/git/pull-requests?view=azure-devops>

