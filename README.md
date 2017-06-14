# ExcelWrapper
Late-bound .NET object for interacting with Excel.

Supports common tasks for manipulating Excel data without requiring a reference to interop COM components.


Features:
- Straight-forward get and set methods for values and formulas - specify by location (e.g. A1, B23, BA21, etc.) or cell row/column coordinates.
- Full support for creating, opening, and saving workbooks.
- Full support for adding, deleting, and moving worksheets.
- Support for common range operations such as setting borders, fonts, colors, and text formatting.
- Support for calling the native Excel Find function.
- Convenience methods for common operations (column number to letter and vice versa, formula reference builder (R1C1 format), etc.).
- Late-bound, so Excel interop libraries are not required to be referenced.
- Direct access to the underlying Excel COM objects to invoke any non-implemented functions.
- Implements IDisposable to prevent 'ghost' Excel instances.
- Full intellisense comments included.

_Note:_ While explicit references to Excel libraries are not required, this component does require that Excel is installed on any system which utilizes it. This library is not a replacement for Excel, it is a late-bound interface to the Excel COM objects.


## Usage Example

```vb.net
' Create invisible Excel.Application object.
Using xl As New ExcelWrapper.Excel

    ' Add a workbook.
    ' This workbook will not have any sheets.
    Dim xlWorkbook As ExcelWrapper.Workbook = xl.AddWorkbook

    ' Add some sheets to the workbook.
    Dim xlWorksheet1 As ExcelWrapper.Worksheet = xlWorkbook.AddWorksheet("My Sheet")
    Dim xlWorksheet2 As ExcelWrapper.Worksheet =
        xlWorkbook.AddWorksheet("Another Sheet", after:=xlWorksheet1)
    Dim xlWorksheet3 As ExcelWrapper.Worksheet =
        xlWorkbook.AddWorksheet("Before Others", before:=xlWorksheet1)

    ' Show the Excel object.
    xl.Visible = True

    With xlWorksheet1
        ' Bring the sheet to the front.
        .SetActive()

        ' Set values.
        .SetValue(1, 1, 5) ' Row 1, Column 1, Value = 5
        .SetValue(1, "B", 10) ' Row 1, Column B, Value = 10

        ' Manually enter formula instead of using the formula builder.
        .SetFormula("C1", "=RC[-2]+RC[-1]") ' Cell C1, Formula = A1+B1


        Dim xlRange As ExcelWrapper.Range = .GetRange("A3:C5")

        ' Set Cells A3 - C5 to the Excel formula to generate a random number.
        xlRange.Formula = "=RAND()"

        ' Place a border around the top and right sides.
        xlRange.SetBorder(style:=ExcelWrapper.Constants.BorderStyle.Single,
                            weight:=ExcelWrapper.Constants.BorderWeight.Normal,
                            color:=RGB(0, 0, 0), modifyTop:=True, modifyRight:=True)

        ' Set the background color to blue.
        xlRange.SetProperties(backColor:=RGB(0, 0, 255))


        ' Build an Excel formula reference (R1C1 format) for A3 - C5
        ' with respect to the destination Cell, D5.
        Dim formulaString As String =
            String.Format("=AVERAGE({0}:{1})",
                            ExcelWrapper.Util.GetFormulaReference(5, "D", 3, "A"),
                            ExcelWrapper.Util.GetFormulaReference(5, 4, 5, 3))

        ' Row 4, Column 4 (Cell D5), Formula =AVERAGE(R[-2]C[-3]:RC[-1])
        .SetFormula(5, 4, formulaString)

        ' Write the generated formula to D6.
        .SetValue("D6", "'" & formulaString)


        ' Set the first row to bold, italics, and red.
        .GetRange("1:1").SetFont(bold:=True, italic:=True, color:=RGB(255, 0, 0))

        ' Set Column D to be 3x the width.
        ' Can use either D or 4 to reference the column.
        .SetColumnWidth("D", .GetColumnWidth(4) * 3)

        ' Center the text of column D.
        .GetRange("D:D").SetProperties(
            horizontalAlign:=ExcelWrapper.Constants.HorizontalAlignment.Center)

    End With

    ' Hide Excel.
    xl.Visible = False

    ' Save.
    xlWorkbook.Save("Test File.xlsx")

    ' End Using will release all Excel resources and display the hidden application.
End Using
```