# Third-party notices

Excel Scenario Analysis Tool is distributed under the MIT License. It uses or was informed by the following open-source projects.

## xlsx-review

The workbook comparison design was informed by [drpedapati/xlsx-review](https://github.com/drpedapati/xlsx-review) at commit [`c39bcf7e6817b79da08852ccc48c950b097fd1b9`](https://github.com/drpedapati/xlsx-review/commit/c39bcf7e6817b79da08852ccc48c950b097fd1b9). Excel Scenario Analysis Tool is an independent C# implementation and has no runtime dependency on xlsx-review.

Copyright (c) 2026 CinciNeuro / Henry Bloomingdale

MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## Handy

[cjpais/Handy](https://github.com/cjpais/Handy) at commit [`00d255492d3aa3585c762142d54273447b6752bc`](https://github.com/cjpais/Handy/commit/00d255492d3aa3585c762142d54273447b6752bc) was used only as an interaction and visual-principles reference.

Copyright (c) 2025 CJ Pais

Handy is MIT licensed. No Handy source code, name, logo, hand icon, illustrations, or brand assets are included in Excel Scenario Analysis Tool.

## Microsoft Open XML SDK

Excel Scenario Analysis Tool uses [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK), licensed under the MIT License. Copyright (c) .NET Foundation and Contributors.

## Microsoft.Data.Sqlite

Excel Scenario Analysis Tool uses [Microsoft.Data.Sqlite](https://github.com/dotnet/efcore), licensed under the MIT License. Copyright (c) .NET Foundation and Contributors.

## SQLitePCLRaw 2.1.12 and SQLite

Microsoft.Data.Sqlite depends on [SQLitePCLRaw](https://github.com/ericsink/SQLitePCL.raw) 2.1.12, including its `bundle_e_sqlite3`, `core`, `lib.e_sqlite3`, and `provider.e_sqlite3` packages.

SQLitePCLRaw is licensed under the Apache License, Version 2.0. Copyright 2014–2024 SourceGear, LLC. The full license is included at `licenses/SQLitePCLRaw-2.1.12-Apache-2.0.txt` in the installed application. The bundled SQLite library is dedicated to the public domain by its authors.

## .NET 10.0.11 self-contained runtime

The standalone ARM64 application includes the Microsoft .NET and Windows Desktop runtimes. The release build copies `LICENSE.TXT`, `THIRD-PARTY-NOTICES.TXT`, and the Windows Desktop `LICENSE` directly from the exact 10.0.11 runtime packages into the installed application under `licenses/dotnet/`.

## Microsoft Windows SDK for .NET 10.0.26100.57

The published WPF application includes `Microsoft.Windows.SDK.NET.dll` and `WinRT.Runtime.dll` from `Microsoft.Windows.SDK.NET.Ref` 10.0.26100.57. Microsoft identifies the package's license at [aka.ms/WinSDKLicenseURL](https://aka.ms/WinSDKLicenseURL). The exact license retrieved for this package version is included as `licenses/dotnet/Microsoft-Windows-SDK-10.0.26100.57-License.rtf`; its SHA-256 is `dd07eb178e00c6bba4148457fc00ff77cd4887eb521d504186fe59c9ec8bbe62`.
