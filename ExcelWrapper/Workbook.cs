// Usage of this code is through the MIT License. See LICENSE file.

namespace ExcelWrapper;

/// <summary> Late-bound object wrapper to the <c>Excel.Workbook</c> COM object. </summary>
public class Workbook
{
    /// <summary>
    /// When a new <c>Excel.Workbook</c> is created, this is the reference for the <c>Excel.Worksheet.</c>
    /// <para>It is automatically deleted by the <see cref="AddWorksheet"/> unless it is explicitly referenced first via
    /// <see cref="GetWorksheet"/>.</para>
    /// </summary>
    ///
    /// <remarks> When an existing <c>Excel.Workbook</c> is opened, this is not used. </remarks>
    private dynamic _defaultWorksheet = null;

    /// <summary> Creates a new <c>Excel.Workbook</c> late-bound object. </summary>
    ///
    /// <remarks>
    /// The constructor is internal only to enforce using the <see cref="Excel"/> object.
    /// <para> When a new workbook is created, it is created with only a single sheet which is marked for deletion (Excel
    /// requires a workbook to have at least one sheet). This sheet is not really intended to be used (as it is expected
    /// any sheets needed will be explicitly created), so it will be automatically removed when a new sheet is added
    /// unless a <see cref="Worksheet"/> object reference is linked to it first (using the <see cref="GetWorksheet"/>
    /// method).</para>
    /// </remarks>
    ///
    /// <param name="owner"> Parent <see cref="Excel"/> object. </param>
    /// <param name="excelWorkbook"> <c>Excel.Workbook</c> object for this instance. </param>
    /// <param name="existingFileName"> (Optional) Excel workbook file name for the respective <c>Excel.Workbook</c>
    ///     object. </param>
    internal Workbook(Excel owner, dynamic excelWorkbook, string existingFileName = "")
    {
        Excel = owner;
        RawObject = excelWorkbook;
        FileName = existingFileName;

        if (string.IsNullOrEmpty(existingFileName))
        {
            // Clear any existing worksheets excel for the first.
            // Cannot remove all worksheets because Excel requires there to be one.
            for (var i = Convert.ToInt32(RawObject.Worksheets.Count); i >= 2; i -= 1)
            {
                RawObject.Sheets(i).Delete();
            }

            // The remaining worksheet will be cleaned up automatically when the user creates a sheet.

            // Set a simple marker to note this sheet should be deleted.
            _defaultWorksheet = RawObject.Sheets(1);
        }
    }

    /// <summary> Gets the owning <see cref="ExcelWrapper.Excel"/> object. </summary>
    ///
    /// <value> <see cref="ExcelWrapper.Excel"/> object which contains the current instance. </value>
    public Excel Excel { get; }

    /// <summary> Gets or sets the raw <c>Excel.Workbook</c> object. </summary>
    ///
    /// <value> <c>Excel.Workbook</c> object. </value>
    public dynamic RawObject { get; private set; }

    /// <summary> Internal list of all <see cref="Worksheet"/>s which have been explicitly referenced. </summary>
    private List<Worksheet> Worksheets { get; } = [];

    /// <summary> Returns whether the <c>Excel.Workbook</c> object is <see langword="null"/>. </summary>
    ///
    /// <value> <see langword="bool"/> indicating if the <c>Excel.Workbook</c> object is empty. </value>
    public bool IsNull => RawObject is null;

    /// <summary> Gets the name of the workbook. Note this can be different from the <see cref="FileName"/>. </summary>
    ///
    /// <value> Name of the <c>Excel.Workbook</c> object. </value>
    public string Name => RawObject.Name.ToString();

    /// <summary>
    /// Gets or sets the file name of the respective workbook, or an empty string if it has not been saved.
    /// </summary>
    ///
    /// <value> File name of the current workbook, or an empty string if it has not been saved. </value>
    public string FileName { get; private set; }

    /// <summary> Gets the total number of <c>Excel.Worksheets</c> contained in the Excel workbook. </summary>
    ///
    /// <remarks> This does not necessarily equal the internal object which contains <see cref="Worksheet"/>s. </remarks>
    ///
    /// <value> Total number of Excel worksheets. </value>
    public int WorksheetCount => Convert.ToInt32(RawObject.Sheets.Count);

    /// <summary> Adds and returns a new <see cref="Worksheet"/> to the current object. </summary>
    ///
    /// <remarks>
    /// Precedence to the location of the new worksheet is as follows:
    /// <paramref name="before"/>, <paramref name="after"/>, <paramref name="toEnd"/>
    /// </remarks>
    ///
    /// <param name="name"> (Optional) Name of the new <see cref="Worksheet"/>. </param>
    /// <param name="before"> (Optional) <see cref="Worksheet"/> to add the new sheet before. </param>
    /// <param name="after"> (Optional) <see cref="Worksheet"/> to add the new sheet after. </param>
    /// <param name="toEnd"> (Optional) When <see langword="true"/>, the new worksheet is added as the last sequentially. </param>
    ///
    /// <returns> New <see cref="Worksheet"/> instance. </returns>
    public Worksheet AddWorksheet(string name = "", Worksheet before = default, Worksheet after = default, bool toEnd = true)
    {
        CleanWorksheets();

        Worksheet newWorksheet;

        // Check if new sheet being requested already exists.
        if (WorksheetExists(name))
        {
            newWorksheet = GetWorksheet(name);
        }
        else
        {
            // Does not exist, create it.
            dynamic xlWorksheet;
            if (before is { })
            {
                xlWorksheet = RawObject.Sheets.Add(Before: before.RawObject);
            }
            else if (after is { })
            {
                xlWorksheet = RawObject.Sheets.Add(After: after.RawObject);
            }
            else if (toEnd)
            {
                xlWorksheet = RawObject.Sheets.Add(After: RawObject.Sheets(WorksheetCount));
            }
            else
            {
                xlWorksheet = RawObject.Sheets.Add();
            }
            newWorksheet = new Worksheet(this, xlWorksheet, name: name);

            Worksheets.Add(newWorksheet);
        }

        if (_defaultWorksheet is { })
        {
            // The original sheet which is marked for deletion is still in the workbook.
            // See what action should be done with it.

            if (Worksheets.All(worksheet => worksheet.Name != _defaultWorksheet.Name))
            {
                // Original sheet is not being used, delete it.
                _defaultWorksheet.Delete();
            }

            // No more action should be done with the original sheet.
            _defaultWorksheet = null;
        }

        return newWorksheet;
    }

    /// <summary> Moves the <paramref name="sheetToMove"/> within the <c>Excel.Workbook</c>. </summary>
    ///
    /// <remarks>
    /// In the event neither <paramref name="moveBefore"/> or <paramref name="moveAfter"/>, no action is taken. If both
    /// <paramref name="moveBefore"/> and <paramref name="moveAfter"/> are provide, <paramref name="moveBefore"/> takes
    /// priority.
    /// </remarks>
    ///
    /// <param name="sheetToMove"> <see cref="Worksheet"/> object to move. </param>
    /// <param name="moveBefore"> (Optional) <see cref="Worksheet"/> object to move <paramref name="sheetToMove"/> before.
    ///     If this is specified, do not provide a value for <paramref name="moveAfter"/>. </param>
    /// <param name="moveAfter"> (Optional) <see cref="Worksheet"/> object to move <paramref name="sheetToMove"/> after.
    ///     If this is specified, do not provide a value for <paramref name="moveBefore"/>. </param>
    public void MoveWorksheet(Worksheet sheetToMove, Worksheet moveBefore = default, Worksheet moveAfter = default)
    {
        if (moveBefore is { })
        {
            RawObject.Sheets(sheetToMove.Name).Move(Before: moveBefore.RawObject);
        }
        else if (moveAfter is { })
        {
            RawObject.Sheets(sheetToMove.Name).Move(After: moveAfter.RawObject);
        }
    }

    /// <summary> Returns a <see cref="Worksheet"/> object for the requested <paramref name="reference"/>. </summary>
    ///
    /// <remarks>
    /// If the specified <paramref name="reference"/> does not exist, an exception will be thrown. Use <see cref="WorksheetExists"/>
    /// to verify if <paramref name="reference"/> does exist.
    /// </remarks>
    ///
    /// <param name="reference"> Excel worksheet name or index number (1 based). </param>
    ///
    /// <returns> <see cref="Worksheet"/> object. </returns>
    public Worksheet GetWorksheet(object reference)
    {
        CleanWorksheets();

        var xlWorksheet = RawObject.Sheets(reference);

        if (Worksheets.FirstOrDefault(worksheet => worksheet.Name == xlWorksheet.Name) is not { } targetWorksheet)
        {
            targetWorksheet = new Worksheet(this, xlWorksheet);
            Worksheets.Add(targetWorksheet);
        }
        return targetWorksheet;
    }

    /// <summary> Returns whether an Excel.Worksheet object at the given <paramref name="reference"/> exists. </summary>
    ///
    /// <remarks> This check is performed against the raw <c>Excel.Workbook</c> object. </remarks>
    ///
    /// <param name="reference"> <inheritdoc cref="GetWorksheet" path="/param[@name='reference']"/> </param>
    public bool WorksheetExists(object reference)
    {
        try
        {
            // To check, attempt to use the reference.
            _ = RawObject.Sheets(reference);

            // If we get here, all is well.
            return true;
        }
        catch (Exception)
        {
            // Any exception indicates it does not exist.
            return false;
        }
    }

    /// <summary> Saves the current Excel workbook to disk. </summary>
    ///
    /// <param name="fileName"> Output file name. This will set <see cref="FileName"/>. </param>
    /// <param name="format"> (Optional) Format to write the workbook as. Any <c>XlFileFormat</c> value is accepted. If
    ///     <see cref="Constants.SaveFormat.AutoDetermine"/> is used, then filenames ending in 'xls' or 'csv' are
    ///     automatically saved in their respective format with everything else using <see cref="Constants.SaveFormat.Default"/>.</param>
    /// <param name="password"> (Optional) When provided, sets the protection password on the file. This is only
    ///     applicable for native Excel formats. The Excel limit is 15 chars. </param>
    public void Save(string fileName, Constants.SaveFormat format = Constants.SaveFormat.AutoDetermine, string password = "")
    {
        FileName = fileName;

        if (format == Constants.SaveFormat.AutoDetermine)
        {
            format = System.IO.Path.GetExtension(fileName).ToLower() switch
            {
                ".csv" => Constants.SaveFormat.CSV,
                ".xls" => Constants.SaveFormat.Legacy,
                _ => Constants.SaveFormat.Default,
            };
        }

        RawObject.SaveAs(fileName, FileFormat: format, Password: password);
    }

    /// <summary> Closes the current Excel workbook. </summary>
    ///
    /// <remarks> Any linked <see cref="Worksheet"/> objects will be released/set to <see langword="null"/>. </remarks>
    public void Close()
    {
        ReleaseAllWorksheets();
        RawObject.Close();
        RawObject = null;
        Excel.CleanWorkbooks();
    }

    /// <summary> Activates/brings to front the respective Excel workbook. </summary>
    public void SetActive() => RawObject.Activate();

    /// <summary> Removes any dead/empty <see cref="Worksheet"/> objects from the internal list. </summary>
    internal void CleanWorksheets()
    {
        for (var i = Worksheets.Count - 1; i >= 0; i--)
        {
            if (Worksheets[i].IsNull)
            {
                Worksheets.RemoveAt(i);
            }
        }
    }

    /// <summary> Sets any <see cref="Worksheet"/> objects linked to this object to <see langword="null"/>. </summary>
    private void ReleaseAllWorksheets()
    {
        for (var i = 0; i < Worksheets.Count; i++)
        {
            Worksheets[i] = default;
        }
        Worksheets.Clear();
    }
}