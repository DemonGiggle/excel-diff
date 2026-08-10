# SignPath onboarding

The repository-side SignPath preparation is stored here. After SignPath
Foundation approves the project:

## Before applying

1. Push the MIT license, policies, lock files, workflow, and artifact
   configuration from this change to the default branch.
2. Confirm multi-factor authentication is enabled on the `DemonGiggle` GitHub
   account.
3. Add this link to the existing v0.1 and v0.2 GitHub release notes:
   `https://github.com/DemonGiggle/excel-diff/blob/main/CODE_SIGNING_POLICY.md`.
4. Confirm the existing portable ZIP release remains publicly downloadable;
   SignPath requires the project to have already released the artifact type it
   wants signed.
5. Apply at https://signpath.org/apply with these project details:
   - Project name/handle: `Excel Compare`
   - Homepage and repository: `https://github.com/DemonGiggle/excel-diff`
   - License: `MIT`
   - Current release: `https://github.com/DemonGiggle/excel-diff/releases/tag/v0.2`
   - Artifact: self-contained Windows x64 portable ZIP containing
     `ExcelDiff.exe`
   - Privacy: offline, no telemetry, no network transfers, workbooks opened
     read-only

## After approval

1. Create or select the SignPath project for Excel Compare.
2. Upload `artifact-configuration.xml` as an artifact configuration with slug
   `portable-windows-v1`.
3. Create a signing policy with slug `release-signing`, using the SignPath
   Foundation certificate and requiring manual approval.
4. Connect this GitHub repository as the trusted build system.
5. Create a SignPath API token for a submitter that can use that policy.
6. Add the token as the GitHub Actions secret `SIGNPATH_API_TOKEN`.
7. Add `SIGNPATH_ORGANIZATION_ID` and `SIGNPATH_PROJECT_SLUG` as GitHub Actions
   repository variables.
8. Enable multi-factor authentication for every GitHub and SignPath team member.
9. Create a release tag such as `v0.3.0`, then run the **Build and sign release**
   workflow from that tag and enter `0.3.0` as the version.
10. Approve the pending signing request in SignPath. Download the resulting
    signed ZIP and checksum from the completed GitHub Actions run and attach
    both to the matching GitHub release.

The workflow deliberately requires a tag and a manual SignPath approval. Never
place a signing token in the repository or use it for unreviewed pull-request
builds.
