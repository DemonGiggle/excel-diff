# Code signing policy

Free code signing provided by [SignPath.io](https://signpath.io/), certificate
by [SignPath Foundation](https://signpath.org/).

## Signed artifacts

Official signed Windows releases are built from the public source repository by
GitHub Actions and submitted directly to SignPath. SignPath signs only the
project-owned `ExcelDiff.exe` and `ExcelDiff.dll` files. Framework and
third-party binaries included in the portable package are not re-signed by this
project.

Each signing request is tied to a tagged source revision, uses the repository's
version-controlled build workflow and artifact configuration, and requires
manual approval in SignPath. Release downloads should be obtained from:

https://github.com/DemonGiggle/excel-diff/releases

## Team roles

- Committers and reviewers: [DemonGiggle](https://github.com/DemonGiggle)
- Approvers: [DemonGiggle](https://github.com/DemonGiggle)

Contributions from people who are not committers must be submitted as pull
requests and reviewed by a committer before merging. Every release signing
request must be approved by an approver. All team members with repository or
SignPath access must use multi-factor authentication.

## Privacy

See the project's [privacy policy](PRIVACY.md). In summary: this program will
not transfer any information to other networked systems unless specifically
requested by the user or the person installing or operating it.

## Verification

On Windows, the signatures can be inspected with PowerShell:

```powershell
Get-AuthenticodeSignature .\ExcelDiff.exe
Get-AuthenticodeSignature .\ExcelDiff.dll
```

For an official signed release, both commands must report `Status` as `Valid`
and identify SignPath Foundation as the signer. Published releases also include
a SHA-256 checksum for the portable ZIP.
