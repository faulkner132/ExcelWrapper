// Usage of this code is through the MIT License. See LICENSE file.

namespace ExcelWrapper;

/// <summary> Late-bound object wrapper to the <c>Excel.Worksheet</c> COM object. </summary>
public class Worksheet
{
    /// <summary> Creates a new <c>Excel.Worksheet</c> late-bound object. </summary>
    ///
    /// <remarks>
    /// The constructor is internal only to enforce using the respective <see cref="ExcelWrapper.Workbook"/> object.
    /// </remarks>
    ///
    /// <param name="owner"> Parent <see cref="ExcelWrapper.Workbook"/> object. </param>
    /// <param name="excelWorksheeet"> <c>Excel.Worksheet</c> object for this instance. </param>
    /// <param name="name"> (Optional) Name to assign to the new sheet. </param>
    internal Worksheet(Workbook owner, dynamic excelWorksheeet, string name = "")
    {
        Workbook = owner;
        RawObject = excelWorksheeet;

        if (!string.IsNullOrEmpty(name))
        {
            Name = name;
        }
    }

    /// <summary> Gets the owning <see cref="Workbook"/> object. </summary>
    ///
    /// <value> <see cref="Workbook"/> object which contains the current instance. </value>
    public Workbook Workbook { get; }

    /// <summary> Gets or sets the raw <c>Excel.Worksheet</c> object. </summary>
    ///
    /// <value> Raw <c>Excel.Worksheet</c> object. </value>
    public dynamic RawObject { get; private set; }

    /// <summary> Returns whether the <c>Excel.Worksheet</c> object is <see langword="null"/>. </summary>
    ///
    /// <value> <see langword="bool"/> indicating if the <c>Excel.Worksheet</c> object is <see langword="null"/>. </value>
    public bool IsNull => RawObject is null;

    /// <summary> Gets or sets the name of the <c>Excel.Worksheet</c>. </summary>
    ///
    /// <remarks> Names must be unique in each <c>Excel.Workbook</c>. The Excel limit is 31 characters. </remarks>
    ///
    /// <value> Worksheet name. Values longer than 31 characters will be truncated. </value>
    public string Name
    {
        get => RawObject.Name.ToString();
        set => RawObject.Name = value.Substring(0, Math.Min(value.Length, 31));
    }

    /// <summary> Gets the <paramref name="column"/> width. </summary>
    ///
    /// <param name="column"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    ///
    /// <returns> Current column width. </returns>
    public float GetColumnWidth(object column) => Convert.ToSingle(RawObject.Columns(column).ColumnWidth);

    /// <summary> Set the <paramref name="column"/> width. </summary>
    ///
    /// <param name="column"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    /// <param name="width"> Width to set. </param>
    public void SetColumnWidth(object column, float width) => RawObject.Columns(column).ColumnWidth = width;

    /// <summary> Gets the <paramref name="row"/> height. </summary>
    ///
    /// <param name="row"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> </param>
    ///
    /// <returns> Current row height. </returns>
    public float GetRowHeight(int row) => Convert.ToSingle(RawObject.Rows(row).RowHeight);

    /// <summary> Sets the <paramref name="row"/> height. </summary>
    ///
    /// <param name="row"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> </param>
    /// <param name="height"> Height to set. </param>
    public void SetRowHeight(int row, float height) => RawObject.Rows(row).RowHeight = height;

    /// <summary>
    /// Returns the Excel formula string in the specified cell. If there is no formula, the value of the cell is returned.
    /// </summary>
    ///
    /// <param name="row"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> </param>
    /// <param name="column"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    ///
    /// <returns> Excel formula string if applicable, otherwise the cell value as a string. </returns>
    public string GetFormula(int row, object column) => RawObject.Cells(row, column).FormulaR1C1.ToString();

    /// <summary> Sets the Excel formula string for the specified cell. </summary>
    ///
    /// <remarks> A static value can be provided for <paramref name="formula"/> as well. </remarks>
    ///
    /// <param name="row"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> </param>
    /// <param name="column"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    /// <param name="formula"> Excel formula string to apply to the cell. </param>
    public void SetFormula(int row, object column, string formula) => RawObject.Cells(row, column).FormulaR1C1 = formula;

    /// <param name="location"> <inheritdoc cref="GetValue(string)" path="/param[@name='location']"/> </param>
    ///
    /// <inheritdoc cref="GetFormula(int, object)"/>
    public string GetFormula(string location) => RawObject.Range(location).FormulaR1C1.ToString();

    /// <param name="location"> <inheritdoc cref="GetValue(string)" path="/param[@name='location']"/> </param>
    /// <param name="formula"> <inheritdoc cref="SetFormula(int, object, string)" path="/param[@name='formula']"/> </param>
    ///
    /// <inheritdoc cref="SetFormula(int, object, string)"/>
    public void SetFormula(string location, string formula) => RawObject.Range(location).FormulaR1C1 = formula;

    /// <summary>
    /// Returns the value of the specified cell as a raw object. If no value is present, this will be <see langword="null"/>.
    /// </summary>
    ///
    /// <param name="row"> Row number. </param>
    /// <param name="column"> Column letter or number (1 based). </param>
    ///
    /// <returns> Static value of the specified cell. If there is no value, <see langword="null"/> is returned. </returns>
    public object GetValue(int row, object column) => RawObject.Cells(row, column).Value;

    /// <summary>
    /// Returns the value of the specified cell as type <typeparamref name="T"/>. If the value cannot be converted to the
    /// provided type then the default value of <typeparamref name="T"/> (or an empty string if <typeparamref name="T"/>
    /// is <see langword="string"/>) is returned.
    /// </summary>
    ///
    /// <typeparam name="T"> Type to convert the value to. </typeparam>
    /// <param name="row"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> </param>
    /// <param name="column"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    ///
    /// <returns>
    /// Value of the specified cell as <typeparamref name="T"/>, or the default value of <typeparamref name="T"/>.
    /// </returns>
    public T GetValue<T>(int row, object column) => GetValue<T>(GetValue(row, column));

    /// <summary> Sets the value of the specified cell. </summary>
    ///
    /// <param name="row"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> </param>
    /// <param name="column"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    /// <param name="value"> Value to set. </param>
    public void SetValue(int row, object column, object value) => RawObject.Cells(row, column).Value = value;

    /// <summary>
    /// Returns the value of the specified cell as a raw object. If no value is present, this will be <see langword="null"/>.
    /// </summary>
    ///
    /// <param name="location"> Cell location (e.g. A5, B22, AD32, etc.) or named reference. </param>
    ///
    /// <returns> Static value of the specified cell. If there is no value, <see langword="null"/> is returned. </returns>
    public object GetValue(string location) => RawObject.Range(location).Value;

    /// <param name="location"> <inheritdoc cref="GetValue(string)" path="/param[@name='location']"/> </param>
    ///
    /// <inheritdoc cref="GetValue{T}(int, object)"/>
    public T GetValue<T>(string location) => GetValue<T>(GetValue(location));

    /// <summary> Sets the value of the specified cell. </summary>
    ///
    /// <param name="location"> <inheritdoc cref="GetValue(string)" path="/param[@name='location']"/> </param>
    /// <param name="value"> <inheritdoc cref="SetValue(int, object, object)" path="/param[@name='value']"/> </param>
    public void SetValue(string location, object value) => RawObject.Range(location).Value = value;

    /// <summary> Helper method for <see cref="GetValue{T}(int, object)"/> and <see cref="GetValue{T}(string)"/> </summary>
    ///
    /// <param name="value"> Value to convert. </param>
    ///
    /// <inheritdoc cref="GetValue{T}(int, object)"/>
    private static T GetValue<T>(object value)
    {
        var returnValue = default(T);

        if (value is null)
        {
            return typeof(T) == typeof(string)
                ? (T)Convert.ChangeType(string.Empty, typeof(T))
                : returnValue;
        }

        // Apply special conversion rules.
        if (typeof(T) == typeof(string))
        {
            value = value.ToString();
        }
        else if (typeof(T) == typeof(bool))
        {
            _ = bool.TryParse(value.ToString(), out var boolValue);
            value = boolValue;
        }

        // Attempt conversion to the requested type.
        try
        {
            returnValue = (T)Convert.ChangeType(value, typeof(T));
        }
        catch (Exception)
        {
            // Ignore conversion failure, use default value.
        }
        return returnValue;
    }

    /// <summary> Returns a <see cref="Range"/> object for all cells in the current Worksheet. </summary>
    ///
    /// <returns> <see cref="Range"/> object. </returns>
    public Range GetAllCells() => new(RawObject.Cells);

    /// <summary> Returns the last column with data in the specified <paramref name="row"/>. </summary>
    ///
    /// <param name="row"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> This will be used in
    ///     determination. </param>
    ///
    /// <returns> Last column number (1 based) containing data. </returns>
    public int GetMaxColumn(int row) => Convert.ToInt32(RawObject.Cells(row, RawObject.Columns.Count).End(-4159).Column); // xlToLeft

    /// <summary> Returns the last row with data in the specified <paramref name="column"/>. </summary>
    ///
    /// <param name="column"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> This will be used in
    ///     determination. </param>
    ///
    /// <returns> Last row number containing data. </returns>
    public int GetMaxRow(object column) => Convert.ToInt32(RawObject.Cells(RawObject.Rows.Count, column).End(-4162).Row); // xlUp

    /// <summary> Returns a <see cref="Range"/> object for the provided Excel range string. </summary>
    ///
    /// <param name="range"> Range string (e.g. A2:D4, A:B, 2:5, etc.) or named reference (single or multi-cell). </param>
    ///
    /// <returns> <see cref="Range"/> object. </returns>
    public Range GetRange(string range) => new(RawObject.Range(range));

    /// <summary> Returns a <see cref="Range"/> object for the area between the provided coordinates. </summary>
    ///
    /// <param name="startRow"> First cell. <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> </param>
    /// <param name="startColumn"> First cell. <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    /// <param name="endRow"> Second cell. <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> </param>
    /// <param name="endColumn"> Second cell. <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    ///
    /// <returns> <see cref="Range"/> object. </returns>
    public Range GetRange(int startRow, object startColumn, int endRow, object endColumn)
    {
        var rangeStrings = new List<string>();
        foreach (var rangeCoordinate in new[]
            {
                new[] { startRow, startColumn },
                [endRow, endColumn],
            })
        {
            var columnLetter = int.TryParse(rangeCoordinate[1].ToString(), out var columnIndex)
                ? Util.GetColumnLetters(columnIndex)
                : (string)rangeCoordinate[1];

            rangeStrings.Add(columnLetter + rangeCoordinate[0]);
        }

        return GetRange(string.Join(":", rangeStrings.ToArray()));
    }

    /// <summary> Returns a <see cref="Range"/> object of the specified <paramref name="column"/>. </summary>
    ///
    /// <param name="column"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    ///
    /// <returns> <see cref="Range"/> object. </returns>
    public Range GetColumn(object column) => new(RawObject.Columns(Util.GetColumnNumber(column.ToString())));

    /// <summary> Returns a <see cref="Range"/> object of the specified <paramref name="row"/>. </summary>
    ///
    /// <param name="row"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> </param>
    ///
    /// <returns> <see cref="Range"/> object. </returns>
    public Range GetRow(int row) => new(RawObject.Rows(row));

    /// <summary> Deletes the provided <paramref name="column"/>. </summary>
    ///
    /// <param name="column"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    public void DeleteColumn(object column) => RawObject.Columns(Util.GetColumnNumber(column.ToString())).Delete();

    /// <summary> Deletes the provided <paramref name="row"/>. </summary>
    ///
    /// <param name="row"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> </param>
    public void DeleteRow(int row) => RawObject.Rows(row).Delete();

    /// <summary> Inserts a new column. </summary>
    ///
    /// <param name="beforeColumn"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> The new
    ///     column is inserted before this. </param>
    public void InsertColumn(object beforeColumn) => RawObject.Columns(Util.GetColumnNumber(beforeColumn.ToString())).Insert();

    /// <summary> Inserts a new row. </summary>
    ///
    /// <param name="beforeRow"> <inheritdoc cref="GetValue(int, object)" path="/param[@name='row']"/> The new row is
    ///     inserted before this. </param>
    public void InsertRow(int beforeRow) => RawObject.Rows(beforeRow).Insert();

    /// <summary> Finds the first cell match for the search parameters. </summary>
    ///
    /// <remarks>
    /// Excel searches will wrap so to prevent processing the same results cell locations should be considered. Use the <see cref="Range.CompareFirstCell"/>
    /// returned <see cref="Range"/> object.
    /// </remarks>
    ///
    /// <param name="valueToFind"> Value to find. </param>
    /// <param name="afterCell"> (Optional) Non-inclusive cell to begin the search from. If not provided, this is set to
    ///     the upper left-most position. </param>
    /// <param name="partialMatch"> (Optional) Specifies if partial matches be allowed. </param>
    /// <param name="searchByRows"> (Optional) Search by rows first (<c>true</c>) or columns first (<c>false</c>). </param>
    /// <param name="searchNext"> (Optional) Search forward [top to bottom / left to right] (<c>true</c>) or backward
    ///     [bottom to top / right to left]. </param>
    /// <param name="searchRange"> (Optional) <see cref="Range"/> to apply the search to. If not provided, all cells are
    ///     searched. </param>
    ///
    /// <returns>
    /// Single cell <see cref="Range"/> object. If there is no match found, the <see cref="Range.RawObject"/> will have a
    /// value of <see langword="null"/> (use <see cref="Range.IsNull"/>).
    /// </returns>
    ///
    /// <seealso cref="FindAll"/>
    public Range Find(object valueToFind, Range afterCell = default, bool partialMatch = false, bool searchByRows = true, bool searchNext = true, Range searchRange = default)
    {
        _lastSearchRange = searchRange ?? GetAllCells();

        var xlRange = _lastSearchRange.RawObject.Find(valueToFind,
            afterCell is null ? Type.Missing : afterCell.RawObject,
            -4163, // xlValues
            partialMatch ? 2 : 1, // xlWhole = 1 / xlPart = 2
            searchByRows ? 1 : 2, // xlByRows = 1 / xlByColumns = 2
            searchNext ? 1 : 2); // xlNext = 1 / xlPrevious = 2

        return new Range(xlRange);
    }

    /// <summary> Last range searched by the <see cref="Find"/> command. </summary>
    ///
    /// <remarks> This is used by the <see cref="FindNext"/> and <see cref="FindPrevious"/> commands. </remarks>
    private Range _lastSearchRange;

    /// <summary>
    /// Finds the next match for the last defined search parameters searching forward (top to bottom / left to right).
    /// </summary>
    ///
    /// <remarks>
    /// Excel searches will wrap so prevent processing the same results, cell locations should be considered. Use the
    /// <see cref="Range.CompareFirstCell"/> returned <see cref="Range"/> object.
    /// <para>This should be used after the <see cref="Find"/> function which sets the respective search range.</para>
    /// </remarks>
    ///
    /// <param name="afterCell"> Non-inclusive cell to continue the search from. </param>
    ///
    /// <returns>
    /// Single cell <see cref="Range"/> object. If there is no match found, the <see cref="Range.RawObject"/> will have a
    /// value of <see langword="null"/> (use <see cref="Range.IsNull"/>).
    /// </returns>
    public Range FindNext(Range afterCell) => new((_lastSearchRange ?? GetAllCells()).RawObject.FindNext(afterCell.RawObject));

    /// <summary>
    /// Finds the next match for the last defined search parameters searching backward (bottom to top / right to left).
    /// </summary>
    ///
    /// <param name="beforeCell"> <inheritdoc cref="FindNext" path="/param[@name='afterCell']"/> </param>
    ///
    /// <inheritdoc cref="FindNext"/>
    public Range FindPrevious(Range beforeCell) => new((_lastSearchRange ?? GetAllCells()).RawObject.FindPrevious(beforeCell.RawObject));

    /// <summary>
    /// Returns all cells which match the search parameters. If no results are found, an empty array is returned.
    /// </summary>
    ///
    /// <remarks>
    /// This is wrapper for the <see cref="Find"/> function and uses the default parameters of that method with regard to
    /// the traversal.
    /// </remarks>
    ///
    /// <returns> Array of <see cref="Range"/> objects. If no results are found, the array is empty. </returns>
    ///
    /// <inheritdoc cref="Find"/>
    public Range[] FindAll(object valueToFind, bool partialMatch = false, Range searchRange = default)
    {
        searchRange ??= GetAllCells();

        var results = new List<Range>();
        var result = Find(valueToFind, partialMatch: partialMatch, searchRange: searchRange);

        while (!result.IsNull
            // When the range exists, we have wrapped around.
            && !results.Any(range => result.CompareFirstCell(range)))
        {
            results.Add(result);
            result = FindNext(result);
        }

        return results.ToArray();
    }

    /// <summary> Sorts columns by value. </summary>
    ///
    /// <param name="sortColumns"> Index of columns and the respective  sort order in the order to apply them. Provide
    ///     the <inheritdoc cref="GetValue(int, object)" path="/param[@name='column']"/> </param>
    /// <param name="hasHeaders"> (Optional) Specifies if the <paramref name="sortRange"/> contains a header row. </param>
    /// <param name="sortRange"> (Optional) <see cref="Range"/> to apply the sort to. If left empty, <see cref="GetAllCells"/>
    ///     is used. </param>
    public void Sort(Dictionary<object, Constants.SortDirection> sortColumns, bool hasHeaders = true, Range sortRange = default)
    {
        RawObject.Sort.SortFields.Clear();
        foreach (var sortColumn in sortColumns)
        {
            RawObject.Sort.SortFields.Add(Key: GetColumn(sortColumn.Key).RawObject, Order: sortColumn.Value, SortOn: 0); // xlSortOnValues
        }
        RawObject.Sort.SetRange((sortRange ?? GetAllCells()).RawObject);
        RawObject.Sort.Header = hasHeaders ? 1 : 2;
        RawObject.Sort.MatchCase = false;
        RawObject.Sort.Orientation = 1; // xlTopToBottom
        RawObject.Sort.SortMethod = 1; // xlPinYin
        RawObject.Sort.Apply();
    }

    /// <summary> Pastes the current clipboard contents to the provided <paramref name="range"/>. </summary>
    ///
    /// <param name="range"> <see cref="Range"/> object to receive the copied information. </param>
    public void Paste(Range range) => RawObject.Paste(range.RawObject);

    /// <summary>
    /// Embeds the specified image file into the sheet at the specified cell location. The image is will be saved inside
    /// the Excel sheet.
    /// </summary>
    ///
    /// <remarks>
    /// When <paramref name="sizeProportional"/> is <see langword="true"/>, priority for scaling is as follows (the first
    /// which provides a value is used):
    /// <list type="number">
    /// <item><paramref name="widthFixed"/></item>
    /// <item><paramref name="heightFixed"/></item>
    /// <item><paramref name="widthPercent"/></item>
    /// <item><paramref name="heightPercent"/></item>
    /// </list>
    /// </remarks>
    ///
    /// <param name="pictureFile"> Image file to insert into the sheet. </param>
    /// <param name="location"> Cell where this image should be placed. The object will be placed in the upper left
    ///     corner of the provided location. </param>
    /// <param name="sizeProportional"> (Optional) Specifies if resizing should be done preserving the aspect ratio. If
    ///     not provided, no resizing is performed. </param>
    /// <param name="widthFixed"> (Optional) Scales the image width to the fixed size, respective to <paramref name="sizeProportional"/>.</param>
    /// <param name="heightFixed"> (Optional) Scales the image height to the fixed size, respective to <paramref name="sizeProportional"/>.</param>
    /// <param name="widthPercent"> (Optional) Scales the image width to the specified percent, respective to <paramref name="sizeProportional"/>.</param>
    /// <param name="heightPercent"> (Optional) Scales the image height to the specified percent, respective to <paramref name="sizeProportional"/>.</param>
    public void InsertPicture(string pictureFile, Range location,
        bool? sizeProportional = null, int? widthFixed = null, int? heightFixed = null, float? widthPercent = null, float? heightPercent = null)
    {
        if (!System.IO.File.Exists(pictureFile))
        {
            return;
        }

        var xlPicture = RawObject.Shapes.AddPicture(pictureFile, Constants.MsoTriState.False, Constants.MsoTriState.True, 1, 1, -1, -1);

        switch (sizeProportional)
        {
            case true:
            {
                xlPicture.LockAspectRatio = Constants.MsoTriState.True;
                if (widthFixed is { })
                {
                    xlPicture.Width = widthFixed;
                }
                else if (heightFixed is { })
                {
                    xlPicture.Height = heightFixed;
                }
                else if (widthPercent is { })
                {
                    xlPicture.ScaleWidth(widthPercent / 100.0f, Constants.MsoTriState.True);
                }
                else if (heightPercent is { })
                {
                    xlPicture.ScaleHeight(heightPercent / 100.0f, Constants.MsoTriState.True);
                }
                break;
            }
            case false:
            {
                xlPicture.LockAspectRatio = Constants.MsoTriState.False;
                if (widthFixed is { } || heightFixed is { })
                {
                    if (widthFixed is { })
                    {
                        xlPicture.Width = widthFixed;
                    }
                    if (heightFixed is { })
                    {
                        xlPicture.Height = heightFixed;
                    }
                }
                else if (widthPercent is { } || heightPercent is { })
                {
                    if (widthPercent is { })
                    {
                        xlPicture.ScaleWidth(widthPercent / 100.0f, Constants.MsoTriState.True);
                    }
                    if (heightPercent is { })
                    {
                        xlPicture.ScaleHeight(heightPercent / 100.0f, Constants.MsoTriState.True);
                    }
                }
                break;
            }
        }

        xlPicture.Left = location.RawObject.Left;
        xlPicture.Top = location.RawObject.Top;
    }

    /// <summary> Activates/brings to front the respective Excel worksheet. </summary>
    public void SetActive() => RawObject.Activate();

    /// <summary> Deletes the current Excel worksheet object. </summary>
    public void Delete()
    {
        RawObject.Delete();
        RawObject = null;
        Workbook.CleanWorksheets();
    }

    /// <summary>
    /// Generates a lookup of column numbers by (text) value on the respective row. This allows for effectively
    /// retrieving column values based on a header text location.
    /// </summary>
    ///
    /// <remarks>
    /// In the event of duplicate header text (consider the transforms and comparisons applied), only the first
    /// occurrence is added to the lookup.
    /// </remarks>
    ///
    /// <param name="row"> (Optional) Row number to generate the lookup for. </param>
    /// <param name="trimText"> (Optional) Specifies if the header value should be trimmed before adding to the lookup. </param>
    /// <param name="comparison"> (Optional) Specifies the string comparison to use when determining if header text is
    ///     duplicated. </param>
    ///
    /// <returns> Lookup with the column text as the key and the column number as the value. </returns>
    public Dictionary<string, int> GetHeaderIndexes(int row = 1, bool trimText = true, StringComparison comparison = StringComparison.CurrentCulture)
    {
        var headerIndexes = new Dictionary<string, int>();

        var maxColumn = GetMaxColumn(row);
        for (var columnNumber = 1; columnNumber <= maxColumn; columnNumber++)
        {
            if (GetValue<string>(row, columnNumber) is not { } value)
            {
                continue;
            }

            if (trimText)
            {
                value = value.Trim();
            }

            if (!headerIndexes.Keys.Any(key => key.Equals(value, comparison)))
            {
                headerIndexes.Add(value, columnNumber);
            }
        }

        return headerIndexes;
    }
}