// Usage of this code is through the MIT License. See LICENSE file.

using System;
using ExcelWrapper;

Console.WriteLine("Create invisible 'Excel.Application' object.");
using var xl = new Excel();

Console.WriteLine("Add a workbook with no logical sheets.");
var xlWorkbook = xl.AddWorkbook();

Console.WriteLine("Add some sheets to the workbook.");
var xlWorksheet = xlWorkbook.AddWorksheet("My Sheet");
_ = xlWorkbook.AddWorksheet("After Sheet", after: xlWorksheet);
_ = xlWorkbook.AddWorksheet("Before Others", before: xlWorksheet);

Console.WriteLine("Show the Excel object.");
xl.Visible = true;

Console.WriteLine("Bring the sheet to the front.");
xlWorksheet.SetActive();

Console.WriteLine("Set values.");
Console.WriteLine("\tRow 1, Column 1, Value = 5");
xlWorksheet.SetValue(1, 1, 5);
Console.WriteLine("\tRow 1, Column B, Value = 10");
xlWorksheet.SetValue(1, "B", 10);

Console.WriteLine("Manually enter formula instead of using the formula builder.");
Console.WriteLine("\tCell C1, Formula = A1+B1");
xlWorksheet.SetFormula("C1", "=RC[-2]+RC[-1]");

var xlRange = xlWorksheet.GetRange("A3:C5");

Console.WriteLine("Set Cells A3-C5 to the Excel formula to generate a random number.");
xlRange.Formula = "=RAND()";

Console.WriteLine("Place a border around the top and right sides.");
xlRange.SetBorder(style: Constants.BorderStyle.Single, weight: Constants.BorderWeight.Thick, color: System.Drawing.Color.FromArgb(4529), modifyTop: true, modifyRight: true);

Console.WriteLine("Set the background color to yellow.");
xlRange.SetProperties(backColor: System.Drawing.Color.Yellow);

Console.WriteLine("Build an Excel formula reference (R1C1 format) for A3 - C5 with respect to the destination Cell, D5.");
var formulaString = string.Format("=AVERAGE({0}:{1})",
    Util.GetFormulaReference(5, "D", 3, "A"),
    Util.GetFormulaReference(5, 4, 5, 3));

Console.WriteLine($"Row 5, Column 4 (Cell D5), Formula {formulaString}");
xlWorksheet.SetFormula(5, 4, formulaString);

Console.WriteLine("Write the generated formula to D6.");
xlWorksheet.SetValue("D6", "'" + formulaString);

Console.WriteLine("Set the first row to bold, italics, and red.");
xlWorksheet.GetRange("1:1").SetFont(bold: true, italic: true, color: System.Drawing.Color.Red);

Console.WriteLine("Set Column D to be 3x the width.");
xlWorksheet.SetColumnWidth("D", xlWorksheet.GetColumnWidth(4) * 3);

Console.WriteLine("Center the text of column D.");
xlWorksheet.GetRange("D:D").SetProperties(horizontalAlign: Constants.HorizontalAlignment.Center);

Console.WriteLine("Enter some hardcoded values in A8-C10");
xlWorksheet.SetValue("A8", "Test 3");
xlWorksheet.SetValue("A9", "Test 2");
xlWorksheet.SetValue("A10", "Test 1");
xlWorksheet.SetValue("B8", "5");
xlWorksheet.SetValue("B9", "6");
xlWorksheet.SetValue("B10", "4");
xlWorksheet.SetValue("C8", 5);
xlWorksheet.SetValue("C9", 6);
xlWorksheet.SetValue("C10", 7);

Console.WriteLine("For rows 8-10, sort by column A ascending and column C descending");
xlWorksheet.Sort(new System.Collections.Generic.Dictionary<object, Constants.SortDirection>
{
    { "A", Constants.SortDirection.Ascending },
    { "C", Constants.SortDirection.Descending },
}, false, xlWorksheet.GetRange("A8:C10"));

Console.WriteLine("Copy values using Generics from A8-C10 to A12-C14");
for (var column = 1; column <= 3; column++)
{
    for (var row = 8; row <= 10; row++)
    {
        xlWorksheet.SetValue(row + 4, column, xlWorksheet.GetValue<string>(row, column));
    }
}

var applicationDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)!;

Console.WriteLine("Insert image.");
xlWorksheet.InsertPicture(System.IO.Path.Combine(applicationDirectory, "SampleImage.jpg"), xlWorksheet.GetRange("J4"), true, widthPercent: 50); 

Console.WriteLine("Save.");
xlWorkbook.Save(System.IO.Path.Combine(applicationDirectory, "Test File.xlsx"));

Console.WriteLine("Hide Excel.");
xl.Visible = false;

Console.WriteLine("End 'using' will release all Excel resources and display the hidden application.");