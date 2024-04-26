// Usage of this code is through the MIT License. See LICENSE file.

namespace ExcelWrapper;

/// <summary> Late-bound object wrapper to the <c>Excel.Range</c> COM object. </summary>
public class Range
{
    /// <summary> Creates a new <c>Excel.Range</c> late-bound object. </summary>
    ///
    /// <param name="excelRange"> <inheritdoc cref="RawObject" path="/summary"/> </param>
    internal Range(object excelRange)
    {
        RawObject = excelRange;
    }

    /// <summary> Gets the raw <c>Excel.Range</c> object. </summary>
    ///
    /// <value> Raw <c>Excel.Range</c> object. </value>
    public dynamic RawObject { get; }

    /// <summary> Returns whether the <c>Excel.Range</c> object is <see langword="null"/>. </summary>
    ///
    /// <value> <see langword="bool"/> indicating if the <c>Excel.Range</c> object is empty. </value>
    public bool IsNull => RawObject is null;

    /// <summary> Gets the column number (1 based) of the upper left most cell. </summary>
    ///
    /// <value> Column number (1 based). </value>
    public int Column => Convert.ToInt32(RawObject.Column);

    /// <summary> Gets the row number of the upper left most cell. </summary>
    ///
    /// <value> Row number. </value>
    public int Row => Convert.ToInt32(RawObject.Row);

    /// <summary>
    /// Gets or sets the value of the current <c>Excel.Range</c>.
    /// <para>If the range is a single cell, the cell value is returned.
    /// Otherwise, if the range is multiple cells, a 2D array of values is returned with the 0 index value being the
    /// column offset from the upper left most cell and the 1 index being the row offset.</para>
    /// </summary>
    ///
    /// <remarks>
    /// If the <c>Excel.Range</c> is a single cell, the cell formula is returned. Otherwise, if the range is multiple
    /// cells, a 2D array of formulas is returned with the 0 index value being the column offset from the upper left most
    /// cell and the 1 index being the row offset.
    /// </remarks>
    ///
    /// <value> Value to set for all cells in the <c>Excel.Range</c>. </value>
    public object Value
    {
        get => RawObject.Value;
        set => RawObject.Value = value;
    }

    /// <summary>
    /// Gets or sets the Excel formula of the current <c>Excel.Range</c>. A static value can be provided for the value.
    /// Likewise, if the value is static, it is returned.
    /// <para></para>
    /// </summary>
    ///
    /// <inheritdoc cref="Value" path="/remarks"/>
    /// <returns> Formula to set for all cells in the <c>Excel.Range</c>. </returns>
    public string Formula
    {
        get => RawObject.FormulaR1C1.ToString();
        set => RawObject.FormulaR1C1 = value;
    }

    /// <summary>
    /// Gets or sets a hyperlink on the current <c>Excel.Range</c>. An empty string indicates no hyperlink is present and
    /// is also used to remove the hyperlink.
    /// </summary>
    ///
    /// <remarks> Only a single hyperlink is supported on the <c>Excel.Range</c>. </remarks>
    ///
    /// <value> Target URL for the respective <c>Excel.Range</c>. </value>
    public string Hyperlink
    {
        get => (RawObject.Hyperlinks?.Count ?? 0) != 0 ? RawObject.Hyperlinks(1).Address.ToString() : "";
        set
        {
            var xlWorksheet = RawObject.Parent;

            // Clear any existing hyperlinks in the current range.
            foreach (var cell in (IEnumerable)RawObject)
            {
                ((dynamic)cell).Hyperlinks.Delete();
            }

            if (!string.IsNullOrEmpty(value))
            {
                xlWorksheet.Hyperlinks.Add(RawObject, value);
            }
        }
    }

    /// <summary> Sets the border properties of the <c>Excel.Range</c>. </summary>
    ///
    /// <param name="style"> (Optional) Border style to apply. This will accept any <c>XlLineStyle</c> value. </param>
    /// <param name="weight"> (Optional) Border weight to apply. This will accept any valid <c>XlBorderWeight</c> value. </param>
    /// <param name="color"> (Optional) Border color to apply, if specified. <inheritdoc cref="GetExcelColor" path="/param[@name='color']"/></param>
    /// <param name="modifyTop"> (Optional) When <see langword="true"/>, options provided apply to the top border. </param>
    /// <param name="modifyBottom"> (Optional) When <see langword="true"/>, options provided apply to the bottom border. </param>
    /// <param name="modifyLeft"> (Optional) When <see langword="true"/>, options provided apply to the left border. </param>
    /// <param name="modifyRight"> (Optional) When <see langword="true"/>, options provided apply to the right border. </param>
    public void SetBorder(Constants.BorderStyle style = 0, Constants.BorderWeight weight = 0, object color = null,
        bool modifyTop = false, bool modifyBottom = false, bool modifyLeft = false, bool modifyRight = false)
    {
        foreach (var borderEdge in new[]
            {
                new object[] { modifyTop, Constants.BorderEdge.Top },
                [modifyBottom, Constants.BorderEdge.Bottom],
                [modifyLeft, Constants.BorderEdge.Left],
                [modifyRight, Constants.BorderEdge.Right],
            }
            .Where(border => (bool)border[0])
            .Select(border => RawObject.Borders(border[1])))
        {
            if (style != 0)
            {
                borderEdge.LineStyle = style;
            }
            if (weight != 0)
            {
                borderEdge.Weight = weight;
            }
            if (GetExcelColor(color) is { } excelColor)
            {       
                borderEdge.Color = excelColor;
            }
        }
    }

    /// <summary> Sets the font properties of the <c>Excel.Range</c>. </summary>
    ///
    /// <param name="fontName"> (Optional) When provided, sets the font family. </param>
    /// <param name="fontSize"> (Optional) When provided, sets the font size. </param>
    /// <param name="bold"> (Optional) When provided, changes the bold status. </param>
    /// <param name="italic"> (Optional) When provided, changes the italic status. </param>
    /// <param name="color"> (Optional) Font color to apply, if specified. <inheritdoc cref="GetExcelColor" path="/param[@name='color']"/></param>
    public void SetFont(string fontName = null, float? fontSize = null, bool? bold = null, bool? italic = null, object color = null)
    {
        if (!string.IsNullOrEmpty(fontName))
        {
            RawObject.Font.Name = fontName;
        }
        if (fontSize is { })
        {
            RawObject.Font.Size = fontSize;
        }
        if (GetExcelColor(color) is { } excelColor)
        {
            RawObject.Font.Color = excelColor;
        }
        if (bold is { })
        {
            RawObject.Font.Bold = bold;
        }
        if (italic is { })
        {
            RawObject.Font.Italic = italic;
        }
    }

    /// <summary> Sets miscellaneous <c>Excel.Range</c> properties. </summary>
    ///
    /// <param name="horizontalAlign"> (Optional) Horizontal text alignment to apply. </param>
    /// <param name="verticalAlign"> (Optional) Vertical text alignment to apply. </param>
    /// <param name="style"> (Optional) When provided, sets the display style.
    ///     <para>Send a valid string excepted by Excel, e.g. "Currency", "Percent", "Comma", etc.</para> </param>
    /// <param name="numberFormat"> (Optional) When provided, sets the formatting properties.
    ///     <para>Send a valid string excepted by Excel, e.g. "#,##0.00", "m/d/yyyy", "@" (for string), etc.</para> </param>
    /// <param name="wrap"> (Optional) When provided, sets the text wrap status for all cells in the range. </param>
    /// <param name="merge"> (Optional) When provided, merges all cells in the range. </param>
    /// <param name="backColor"> (Optional) Background color to apply, if specified. <inheritdoc cref="GetExcelColor" path="/param[@name='color']"/>.</param>
    public void SetProperties(Constants.HorizontalAlignment horizontalAlign = 0, Constants.VerticalAlignment verticalAlign = 0,
        string style = null, string numberFormat = null, bool? wrap = null, bool? merge = null, object backColor = null)
    {
        if (horizontalAlign != 0)
        {
            RawObject.HorizontalAlignment = horizontalAlign;
        }
        if (verticalAlign != 0)
        {
            RawObject.VerticalAlignment = verticalAlign;
        }
        if (!string.IsNullOrEmpty(style))
        {
            RawObject.Style = style;
        }
        if (!string.IsNullOrEmpty(numberFormat))
        {
            RawObject.NumberFormat = numberFormat;
        }
        if (GetExcelColor(backColor) is { } excelBackColor)
        {
            RawObject.Interior.Color = excelBackColor;
        }
        if (wrap is { })
        {
            RawObject.WrapText = wrap;
        }
        if (merge is { })
        {
            RawObject.MergeCells = merge;
        }
    }

    /// <summary>
    /// Gets the ready to use Excel color from the provided <paramref name="color"/>. Otherwise, <see langword="null"/>
    /// is returned.
    /// </summary>
    ///
    /// <param name="color"> This will accept any valid Excel color value or <see cref="System.Drawing.Color"/>. </param>
    private static int? GetExcelColor(object color)
    {
        if (color is { })
        {
            return color is System.Drawing.Color drawingColor
                ? System.Drawing.ColorTranslator.ToOle(drawingColor)
                : (int)color;
        }
        return null;
    }

    /// <summary> Clears the contents of the <c>Excel.Range</c> object. </summary>
    public void DeleteContents() => RawObject.ClearContents();

    /// <summary> Copies the <c>Excel.Range</c> object to the clipboard. </summary>
    public void Copy() => RawObject.Copy();

    /// <summary>
    /// Compares the upper left most cell between this instance and the passed in <paramref name="range"/> and returns if
    /// they are the same.
    /// </summary>
    ///
    /// <param name="range"> <see cref="Range"/> object to compare to the current instance. </param>
    ///
    /// <returns> <see langword="bool"/> indicating if both <see cref="Range"/>s start in the same cell. </returns>
    public bool CompareFirstCell(Range range) => (Row == range.Row) && (Column == range.Column);
}