''' <summary>
''' Late-bound object wrapper to the Excel.Range COM object.
''' </summary>
''' <remarks></remarks>
Public Class Range


#Region " Instance "

    ''' <summary>
    ''' Raw Excel.Range object used by this wrapper.
    ''' </summary>
    ''' <remarks></remarks>
    Private _range As Object



    ''' <summary>
    ''' Creates a new Excel.Range late-bound object.
    ''' </summary>
    ''' <param name="excelRange">Excel.Range object for this instance.</param>
    ''' <remarks></remarks>
    Friend Sub New(excelRange As Object)

        _range = excelRange

    End Sub

#End Region


#Region " Properties "

    ''' <summary>
    ''' Gets the raw Excel.Range object.
    ''' </summary>
    ''' <returns>Raw Excel.Range object.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property [Object] As Object
        Get
            Return _range
        End Get
    End Property

    ''' <summary>
    ''' Returns whether the Excel.Range object is empty/null.
    ''' </summary>
    ''' <returns>Boolean indicating if the Excel.Range object is empty.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsNull As Boolean
        Get
            Return _range Is Nothing
        End Get
    End Property

    ''' <summary>
    ''' Gets the column of the upper left most cell.
    ''' </summary>
    ''' <returns>Column number.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Column As Integer
        Get
            Return _range.Column
        End Get
    End Property

    ''' <summary>
    ''' Gets the row of the upper left most cell.
    ''' </summary>
    ''' <returns>Row number.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Row As Integer
        Get
            Return _range.Row
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets the value of the current range.
    ''' </summary>
    ''' <value>Value to set for all cells in the range.</value>
    ''' <returns>Current range value(s).</returns>
    ''' <remarks>
    ''' If the range is a single cell, the cell value is returned.
    ''' Otherwise if the range is multiple cells, a 2D array of values is returned
    ''' with the 0 index value being the column offset from the upper left most cell and the 1 index being the row offset.
    ''' </remarks>
    Public Property Value As Object
        Get
            Return _range.Value
        End Get
        Set(value As Object)
            _range.Value = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the Excel formula of the current range.
    ''' </summary>
    ''' <value>Formula to set for all cells in the range.</value>
    ''' <returns>Current range formula(s).</returns>
    ''' <remarks>
    ''' A static value can be provided for the value. Likewise, if the value is static, it is returned.
    ''' If the range is a single cell, the cell formula is returned.
    ''' Otherwise if the range is multiple cells, a 2D array of formulas is returned
    ''' with the 0 index value being the column offset from the upper left most cell and the 1 index being the row offset.
    ''' </remarks>
    Public Property Formula As String
        Get
            Return _range.FormulaR1C1
        End Get
        Set(value As String)
            _range.FormulaR1C1 = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets a hyperlink on the current range.
    ''' An empty string indicates no hyperlink is present and is also used to remove the hyperlink.
    ''' </summary>
    ''' <value>Target URL for the respective range.</value>
    ''' <returns>Target URL address.</returns>
    ''' <remarks>Only a single hyperlink is supported on the range.</remarks>
    Public Property Hyperlink As String
        Get
            If _range.Hyperlinks.Count > 0 Then
                Return _range.Hyperlinks(1).Address
            Else
                Return ""
            End If
        End Get
        Set(value As String)
            Dim xlWorksheet As Object = _range.Parent

            ' Clear any existing hyperlinks in the current range.
            For Each cell As Object In _range
                cell.Hyperlinks.Delete()
            Next

            If value <> "" Then
                xlWorksheet.Hyperlinks.Add(_range, value)
            End If
        End Set
    End Property

#End Region


#Region " Modifiers "

    ''' <summary>
    ''' Sets the border properties of the range.
    ''' </summary>
    ''' <param name="style">Border style to apply. This will accept any 'XlLineStyle' value. Use 0 to ignore.</param>
    ''' <param name="weight">Border weight to apply. This will accept any valid 'XlBorderWeight' value. Use 0 to ignore.</param>
    ''' <param name="color">Border color to apply, if specified. This will accept any valid Excel color value, recommended format is <see cref="RGB"/> function.</param>
    ''' <param name="modifyTop">When True, options set apply to the top border.</param>
    ''' <param name="modifyBottom">When True, options set apply to the bottom border.</param>
    ''' <param name="modifyLeft">When True, options set apply to the left border.</param>
    ''' <param name="modifyRight">When True, options set apply to the right border.</param>
    ''' <remarks></remarks>
    Public Sub SetBorder(Optional style As Constants.BorderStyle = 0, Optional weight As Constants.BorderWeight = 0,
                         Optional color As Object = Nothing,
                         Optional modifyTop As Boolean = False, Optional modifyBottom As Boolean = False,
                         Optional modifyLeft As Boolean = False, Optional modifyRight As Boolean = False)

        For Each mashup As Object() In {
            New Object() {modifyTop, Constants.BorderEdge.Top},
            New Object() {modifyBottom, Constants.BorderEdge.Bottom},
            New Object() {modifyLeft, Constants.BorderEdge.Left},
            New Object() {modifyRight, Constants.BorderEdge.Right}}

            If mashup(0) Then
                With _range.Borders(mashup(1))

                    If style <> 0 Then
                        .LineStyle = style
                    End If

                    If weight <> 0 Then
                        .Weight = weight
                    End If

                    If color IsNot Nothing Then
                        .Color = color
                    End If

                End With
            End If

        Next
    End Sub

    ''' <summary>
    ''' Sets the font properties of the range.
    ''' </summary>
    ''' <param name="fontName">When provided, sets the font family.</param>
    ''' <param name="fontSize">When provided, sets the font size. Specify with a numeric value.</param>
    ''' <param name="bold">When provided, changes the bold status. Specify with a boolean.</param>
    ''' <param name="italic">When provided, changes the italic status. Specify with a boolean.</param>
    ''' <param name="color">Font color to apply, if specified. This will accept any valid Excel color value, recommended format is <see cref="RGB"/> function.</param>
    ''' <remarks></remarks>
    Public Sub SetFont(Optional fontName As Object = Nothing, Optional fontSize As Object = Nothing,
                       Optional bold As Object = Nothing, Optional italic As Object = Nothing,
                       Optional color As Object = Nothing)

        If fontName IsNot Nothing AndAlso TypeOf fontName Is String Then
            _range.Font.Name = fontName
        End If

        If fontSize IsNot Nothing AndAlso IsNumeric(fontSize) Then
            _range.Font.Size = Convert.ToSingle(fontSize)
        End If

        If color IsNot Nothing Then
            _range.Font.Color = color
        End If

        If bold IsNot Nothing AndAlso TypeOf bold Is Boolean Then
            _range.Font.Bold = bold
        End If

        If italic IsNot Nothing AndAlso TypeOf italic Is Boolean Then
            _range.Font.Italic = italic
        End If
    End Sub

    ''' <summary>
    ''' Sets miscellanous range properties.
    ''' </summary>
    ''' <param name="horizontalAlign">Horizontal text alignment to apply. Use 0 to ignore.</param>
    ''' <param name="verticalAlign">Vertical text alignment to apply. Use 0 to ignore.</param>
    ''' <param name="style">
    ''' When provided, sets the display style.
    ''' Send a valid string excepted by Excel, e.g. "Currency", "Percent", "Comma", etc.
    ''' </param>
    ''' <param name="numberFormat">
    ''' When provided, sets the formatting properties.
    ''' Send a valid string excepted by Excel, e.g. "#,##0.00", "m/d/yyyy", "@" (for string), etc.
    ''' </param>
    ''' <param name="wrap">When provided, sets the text wrap status for all cells in the range. Specify with a boolean.</param>
    ''' <param name="merge">When provided, merges all cells in the range. Specify with a boolean.</param>
    ''' <param name="backColor">Background color to apply, if specified. This will accept any valid Excel color value, recommended format is <see cref="RGB"/> function.</param>
    ''' <remarks></remarks>
    Public Sub SetProperties(Optional horizontalAlign As Constants.HorizontalAlignment = 0,
                             Optional verticalAlign As Constants.VerticalAlignment = 0,
                             Optional style As Object = Nothing, Optional numberFormat As Object = Nothing,
                             Optional wrap As Object = Nothing, Optional merge As Object = Nothing,
                             Optional backColor As Object = Nothing)

        If horizontalAlign <> 0 Then
            _range.HorizontalAlignment = horizontalAlign
        End If

        If verticalAlign <> 0 Then
            _range.VerticalAlignment = verticalAlign
        End If

        If style IsNot Nothing AndAlso TypeOf style Is String Then
            _range.Style = style
        End If

        If numberFormat IsNot Nothing AndAlso TypeOf numberFormat Is String Then
            _range.NumberFormat = numberFormat
        End If

        If backColor IsNot Nothing Then
            _range.Interior.Color = backColor
        End If

        If wrap IsNot Nothing AndAlso TypeOf wrap Is Boolean Then
            _range.WrapText = wrap
        End If

        If merge IsNot Nothing AndAlso TypeOf merge Is Boolean Then
            _range.MergeCells = merge
        End If

    End Sub

#End Region


#Region " Methods "

    ''' <summary>
    ''' Clears the contents of the Excel range object.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DeleteContents()

        _range.ClearContents()

    End Sub

    ''' <summary>
    ''' Copies the Excel range object to the clipboard.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Copy()
        _range.Copy()
    End Sub

    ''' <summary>
    ''' Compares the upper left most cell between this instance and the passed in <paramref name="range"/> and returns if they are the same.
    ''' </summary>
    ''' <param name="range"><see cref="Range"/> object to compare to the current instance.</param>
    ''' <returns>Boolean indicating if both <see cref="Range"/>s start in the same cell.</returns>
    ''' <remarks></remarks>
    Public Function CompareFirstCell(range As Range) As Boolean

        Return Row = range.Row And Column = range.Column

    End Function

#End Region


End Class
