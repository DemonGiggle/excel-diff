# Third-party notices

Excel Compare's distributed Windows build contains the following open-source
components. They are not authored or signed as project-owned binaries.

| Component | License | Project |
| --- | --- | --- |
| .NET 8 Windows Desktop Runtime | MIT | https://github.com/dotnet/runtime |
| DocumentFormat.OpenXml 3.5.1 and DocumentFormat.OpenXml.Framework 3.5.1 | MIT | https://github.com/dotnet/Open-XML-SDK |
| ExcelDataReader 3.9.0 | MIT | https://github.com/ExcelDataReader/ExcelDataReader |
| System.Text.Encoding.CodePages 8.0.0 | MIT | https://github.com/dotnet/runtime |

Additional transitive Microsoft runtime libraries in the self-contained build
are distributed under the licenses declared by their NuGet packages. Exact
resolved versions are recorded in the committed `packages.lock.json` files.

The original license notices remain authoritative. Source and license details
are available from the linked upstream projects and package metadata on
https://www.nuget.org/.
