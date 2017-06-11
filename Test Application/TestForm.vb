Public Class TestForm

    Private Sub TestForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = My.Application.Info.Title

    End Sub

    Private Sub RunTestButton_Click(sender As Object, e As EventArgs) Handles RunTestButton.Click

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
                ' with respect to the destination Cell, D4.
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

                Dim sortColumns As New Dictionary(Of Object, ExcelWrapper.Constants.SortDirection)
                sortColumns.Add("A", ExcelWrapper.Constants.SortDirection.Ascending)
                sortColumns.Add("C", ExcelWrapper.SortDirection.Descending)

                .SetValue("A8", "Test 3")
                .SetValue("A9", "Test 2")
                .SetValue("A10", "Test 1")
                .SetValue("B8", "4")
                .SetValue("B9", "5")
                .SetValue("B10", "6")
                .SetValue("C8", 5)
                .SetValue("C9", 6)
                .SetValue("C10", 7)

                .Sort(sortColumns, hasHeaders:=True, sortRange:= .GetRange("A8:C10"))

            End With

            ' Hide Excel.
            xl.Visible = False

            ' Save.
            xlWorkbook.Save("Test File.xlsx")

            ' End Using will release all Excel resources and display the hidden application.
        End Using

    End Sub
End Class
