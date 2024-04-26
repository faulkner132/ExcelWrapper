# Change Log

Releases
- [2.0.0](#200-release)
- [1.5.1](#151-release)
- [1.5.0](#150-release)
- [1.4.0](#140-release)
- [1.3.0](#130-release)
- [1.2.0](#120-release)
- [1.1.0](#110-release)
- [1.0.0](#100-release)

---

## 2.0.0 Release
_April 26, 2024_

**Possible compatibility issues: See Notes section below.**

Updates:
- Library updated to C# language targeting both .NET Framework 4.0 and .NET 6

Notes:
- `Object` property names have been renamed to `RawObject`
- `Util.TryGetCellRowColumn` renamed to `Util.GetCellRowColumn`
- Optional parameters in many methods were updated to nullable types
- Color parameters now accept `System.Drawing.Color` values
- `Worksheet.GetHeaderIndexes` replaces `toUpper` parameter with `comparison`

<br/>

## 1.5.1 Release
_September 10, 2019_

Fixes:
- No code changes: previous build was against Debug instead of Release configuration

<br/>

## 1.5.0 Release
_August 24, 2019_

**Possible compatibility issue: See Updates section below.**

New Methods:
- `Worksheet.GetValue<T>` (overloaded)

Updates:
- `Workbook.Save` adds support for the `Constants.SaveFormat.AutoDetermine` save format. This was set as the default save format

<br/>

## 1.4.0 Release
_November 19, 2016_

**Possible compatibility issue: See Updates section below.**

New Methods:
- `Worksheet.FindAll`
- `Worksheet.GetHeaderIndexes`

Updates:
- `Worksheet.Find` adds the `searchRange` parameter
- `Worksheet.FindPrevious` parameter was renamed from `afterCell` to `beforeCell`

<br/>

## 1.3.0 Release
_November 12, 2015_

No anticipated compatibility issues with the previous version.

New Methods:
- `Excel.ToggleStatusBar`
- `Excel.StatusText`
- `Worksheet.InsertPicture`
- `Range.Hyperlink`

<br/>

## 1.2.0 Release
_July 17, 2015_

No anticipated compatibility issues with the previous version.

New Methods:
- `Worksheet.Sort`

<br/>

## 1.1.0 Release
_March 9, 2014_

No anticipated compatibility issues with the previous version.

New Methods:
- `Excel.WorkbookExists`
- `Workbook.WorksheetExists`

Changes:
- `Excel.OpenWorkbook` supports providing a password for opening protected files
- `Workbook.Save` supports providing a password for saving with protection

Fixes:
- Documentation updated with many issues to this file addressed
- `Workbook.AddWorksheet` will verify if the new sheet name already exists before adding a new worksheet

<br/>

## 1.0.0 Release
_February 3, 2014_

Initial release