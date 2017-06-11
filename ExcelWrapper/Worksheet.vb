Imports System.Drawing

''' <summary>
''' Late-bound object wrapper to the Excel.Worksheet COM object.
''' </summary>
''' <remarks></remarks>
Public Class Worksheet


#Region " Instance "

    ''' <summary>
    ''' Parent instance of <see cref="Workbook"/>.
    ''' </summary>
    ''' <remarks></remarks>
    Private _workbook As Workbook

    ''' <summary>
    ''' Raw Excel.Worksheet object used by this wrapper.
    ''' </summary>
    ''' <remarks></remarks>
    Private _worksheet As Object

    ''' <summary>
    ''' Last range searched by the <see cref="Find"/> command.
    ''' This is used by the <see cref="FindNext"/> and <see cref="FindPrevious"/> commands.
    ''' </summary>
    ''' <remarks></remarks>
    Private _lastSearchRange As Range


    ''' <summary>
    ''' Creates a new Excel.Worksheet latebound object.
    ''' </summary>
    ''' <param name="owner">Parent <see cref="Workbook"/> object.</param>
    ''' <param name="excelWorksheeet">Excel.Worksheet object for this instance.</param>
    ''' <param name="name">Name to assign to the new sheet.</param>
    ''' <remarks>The constructor is internal only to enforce using the respective <see cref="Workbook"/> object.</remarks>
    Friend Sub New(ByRef owner As Workbook, ByRef excelWorksheeet As Object, Optional name As String = "")

        _workbook = owner
        _worksheet = excelWorksheeet

        If name <> "" Then
            Me.Name = name
        End If

    End Sub

#End Region


#Region " Properties "

    ''' <summary>
    ''' Gets the owning <see cref="Workbook"/> object.
    ''' </summary>
    ''' <returns><see cref="Workbook"/> object which contains the current instance.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Workbook As Workbook
        Get
            Return _workbook
        End Get
    End Property

    ''' <summary>
    ''' Gets the raw Excel.Worksheet object.
    ''' </summary>
    ''' <returns>Raw Excel.Worksheet object.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property [Object] As Object
        Get
            Return _worksheet
        End Get
    End Property

    ''' <summary>
    ''' Returns whether the Excel.Worksheet object is empty/null.
    ''' </summary>
    ''' <returns>Boolean indicating if the Excel.Worksheet object is empty.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsNull As Boolean
        Get
            Return _worksheet Is Nothing
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets the name of the Worksheet.
    ''' </summary>
    ''' <value>Worksheet name. Values longer than 31 characters will be truncated.</value>
    ''' <returns>Name of the current Worksheet.</returns>
    ''' <remarks>Names must be unique in each Workbook. The Excel limit is 31 characters.</remarks>
    Public Property Name As String
        Get
            Return _worksheet.Name
        End Get
        Set(value As String)
            _worksheet.Name = value.Substring(0, Math.Min(value.Length, 31))
        End Set
    End Property

    ''' <summary>
    ''' Gets the <paramref name="column"/> width.
    ''' </summary>
    ''' <param name="column">Column letter or number.</param>
    ''' <returns>Current column width.</returns>
    ''' <remarks></remarks>
    Public Function GetColumnWidth(column As Object) As Single
        Return _worksheet.Columns(column).ColumnWidth
    End Function
    ''' <summary>
    ''' Set the <paramref name="column"/> width.
    ''' </summary>
    ''' <param name="column">Column letter or number.</param>
    ''' <param name="width">Width to set.</param>
    ''' <remarks></remarks>
    Public Sub SetColumnWidth(column As Object, width As Single)
        _worksheet.Columns(column).ColumnWidth = width
    End Sub

    ''' <summary>
    ''' Gets the <paramref name="row"/> height.
    ''' </summary>
    ''' <param name="row">Row.</param>
    ''' <returns>Current row height.</returns>
    ''' <remarks></remarks>
    Public Function GetRowHeight(row As Integer) As Single
        Return _worksheet.Rows(row).RowHeight
    End Function
    ''' <summary>
    ''' Sets the <paramref name="row"/> height.
    ''' </summary>
    ''' <param name="row">Row.</param>
    ''' <param name="height">Height to set.</param>
    ''' <remarks></remarks>
    Public Sub SetRowHeight(row As Integer, height As Single)
        _worksheet.Rows(row).RowHeight = height
    End Sub

#End Region


#Region " Values "

    ''' <summary>
    ''' Returns the Excel formula string in the specified cell. If there is no formula, the value of the cell is returned.
    ''' </summary>
    ''' <param name="row">Row.</param>
    ''' <param name="column">Column letter or number.</param>
    ''' <returns>Excel formula string if applicable, otherwise the cell value as a string.</returns>
    ''' <remarks></remarks>
    Public Function GetFormula(row As Integer, column As Object) As String
        Return _worksheet.Cells(row, column).FormulaR1C1
    End Function
    ''' <summary>
    ''' Sets the Excel formula string for the specified cell.
    ''' </summary>
    ''' <param name="row">Row.</param>
    ''' <param name="column">Column letter or number.</param>
    ''' <param name="formula">Excel formula string to apply to the cell.</param>
    ''' <remarks>A static value can be provided for <paramref name="formula"/> as well.</remarks>
    Public Sub SetFormula(row As Integer, column As Object, formula As String)
        _worksheet.Cells(row, column).FormulaR1C1 = formula
    End Sub

    ''' <summary>
    ''' Returns the Excel formula string in the specified cell. If there is no formula, the value of the cell is returned.
    ''' </summary>
    ''' <param name="location">Cell location (e.g. A5, B22, AD32, etc.) or named reference.</param>
    ''' <returns>Excel formula string if applicable, otherwise the cell value as a string.</returns>
    ''' <remarks></remarks>
    Public Function GetFormula(location As String) As String
        Return _worksheet.Range(location).FormulaR1C1
    End Function
    ''' <summary>
    ''' Returns the Excel formula string in the specified cell. If there is no formula, the value of the cell is returned.
    ''' </summary>
    ''' <param name="location">Cell location (e.g. A5, B22, AD32, etc.) or named reference.</param>
    ''' <param name="formula">Excel formula string to apply to the cell.</param>
    ''' <remarks>A static value can be provided for <paramref name="formula"/> as well.</remarks>
    Public Sub SetFormula(location As String, formula As String)
        _worksheet.Range(location).FormulaR1C1 = formula
    End Sub

    ''' <summary>
    ''' Returns the value of the specified cell.
    ''' </summary>
    ''' <param name="row">Row.</param>
    ''' <param name="column">Column letter or number.</param>
    ''' <returns>Static value of the specified cell. If there is no value, null is returned.</returns>
    ''' <remarks></remarks>
    Public Function GetValue(row As Integer, column As Object) As Object
        Return _worksheet.Cells(row, column).Value
    End Function
    ''' <summary>
    ''' Sets the value of the specified cell.
    ''' </summary>
    ''' <param name="row">Row.</param>
    ''' <param name="column">Column letter or number.</param>
    ''' <param name="value">Value to set.</param>
    ''' <remarks></remarks>
    Public Sub SetValue(row As Integer, column As Object, value As Object)
        _worksheet.Cells(row, column).Value = value
    End Sub
    ''' <summary>
    ''' Returns the value of the specified cell.
    ''' </summary>
    ''' <param name="location">Cell location (e.g. A5, B22, AD32, etc.) or named reference.</param>
    ''' <returns>Static value of the specified cell. If there is no value, null is returned.</returns>
    ''' <remarks></remarks>
    Public Function GetValue(location As String) As Object
        Return _worksheet.Range(location).Value
    End Function
    ''' <summary>
    ''' Sets the value of the specified cell.
    ''' </summary>
    ''' <param name="location">Cell location (e.g. A5, B22, AD32, etc.) or named reference.</param>
    ''' <param name="value">Value to set.</param>
    ''' <remarks></remarks>
    Public Sub SetValue(location As String, value As Object)
        _worksheet.Range(location).Value = value
    End Sub

#End Region


#Region " Accessors "

    ''' <summary>
    ''' Returns a <see cref="Range"/> object for all cells in the current Worksheet.
    ''' </summary>
    ''' <returns><see cref="Range"/> object.</returns>
    ''' <remarks></remarks>
    Public Function GetAllCells() As Range

        Dim xlRange As Object = _worksheet.Cells
        Return New Range(xlRange)

    End Function

    ''' <summary>
    ''' Returns the last column with data in the specified <paramref name="row"/>.
    ''' </summary>
    ''' <param name="row">Row containing the data to search.</param>
    ''' <returns>Last column number containing data.</returns>
    ''' <remarks></remarks>
    Public Function GetMaxColumn(row As Integer) As Integer
        ' xlToLeft = -4159
        Return _worksheet.Cells(row, _worksheet.Columns.Count).End(-4159).Column
    End Function

    ''' <summary>
    ''' Returns the last row with data in the specified <paramref name="column"/>.
    ''' </summary>
    ''' <param name="column">Column containing the data to search.</param>
    ''' <returns>Last row number containing data.</returns>
    ''' <remarks></remarks>
    Public Function GetMaxRow(column As Object) As Integer
        ' xlUp = -4162
        Return _worksheet.Cells(_worksheet.Rows.Count, column).End(-4162).Row
    End Function

    ''' <summary>
    ''' Returns a <see cref="Range"/> object for the provided Excel range string.
    ''' </summary>
    ''' <param name="range">Range string (e.g. A2:D4, A:B, 2:5, etc.) or named reference (single or multi-cell).</param>
    ''' <returns><see cref="Range"/> object.</returns>
    ''' <remarks></remarks>
    Public Function GetRange(range As String) As Range
        Dim xlRange As Object = _worksheet.Range(range)
        Return New Range(xlRange)
    End Function

    ''' <summary>
    ''' Returns a <see cref="Range"/> object for the area between the provided coordinates.
    ''' </summary>
    ''' <param name="startRow">Row of first cell.</param>
    ''' <param name="startColumn">Column letter or number of the first cell.</param>
    ''' <param name="endRow">Row of the second cell.</param>
    ''' <param name="endColumn">Column letter or number of the second cell.</param>
    ''' <returns><see cref="Range"/> object.</returns>
    ''' <remarks></remarks>
    Public Function GetRange(startRow As Integer, startColumn As Object, endRow As Integer, endColumn As Object) As Range
        Dim rangeStrings As New List(Of String)
        For Each rangeCoordinate As Object() In {New Object() {startRow, startColumn}, New Object() {endRow, endColumn}}
            Dim columnLetter As String = rangeCoordinate(1)
            columnLetter =
                If(IsNumeric(columnLetter), Util.GetColumnLetters(columnLetter), columnLetter)

            rangeStrings.Add(columnLetter & rangeCoordinate(0))
        Next

        Return GetRange(String.Join(":", rangeStrings.ToArray))
    End Function

    ''' <summary>
    ''' Returns a <see cref="Range"/> object of the specified <paramref name="column"/>.
    ''' </summary>
    ''' <param name="column">Column letter or number.</param>
    ''' <returns><see cref="Range"/> object</returns>
    ''' <remarks></remarks>
    Public Function GetColumn(column As Object) As Range
        If Not IsNumeric(column) Then
            column = Util.GetColumnNumber(column)
        End If

        Dim xlRange As Object = _worksheet.Columns(column)
        Return New Range(xlRange)
    End Function

    ''' <summary>
    ''' Returns a <see cref="Range"/> object of the specified <paramref name="row"/>.
    ''' </summary>
    ''' <param name="row">Row.</param>
    ''' <returns><see cref="Range"/> object</returns>
    ''' <remarks></remarks>
    Public Function GetRow(row As Integer) As Range
        Dim xlRange As Object = _worksheet.Rows(row)
        Return New Range(xlRange)
    End Function

#End Region


#Region " Methods "

    ''' <summary>
    ''' Deletes a column.
    ''' </summary>
    ''' <param name="column">Column number or letter to delete.</param>
    ''' <remarks></remarks>
    Public Sub DeleteColumn(column As Object)

        _worksheet.Columns(Util.GetColumnNumber(column)).Delete()

    End Sub

    ''' <summary>
    ''' Deletes a row.
    ''' </summary>
    ''' <param name="row">Row to delete.</param>
    ''' <remarks></remarks>
    Public Sub DeleteRow(row As Integer)

        _worksheet.Rows(row).Delete()

    End Sub

    ''' <summary>
    ''' Inserts a new column.
    ''' </summary>
    ''' <param name="beforeColumn">Current column number or letter to insert the new column before.</param>
    ''' <remarks></remarks>
    Public Sub InsertColumn(beforeColumn As Object)

        _worksheet.Columns(Util.GetColumnNumber(beforeColumn)).Insert()

    End Sub

    ''' <summary>
    ''' Inserts a new row.
    ''' </summary>
    ''' <param name="beforeRow">Current row to insert the new row before.</param>
    ''' <remarks></remarks>
    Public Sub InsertRow(beforeRow As Integer)

        _worksheet.Rows(beforeRow).Insert()

    End Sub

    ''' <summary>
    ''' Finds the first cell match for the search parameters.
    ''' </summary>
    ''' <param name="valueToFind">Value to find.</param>
    ''' <param name="afterCell">Non-inclusive cell to begin the search from.
    ''' If not provided, this is set to the the upper left-most position.</param>
    ''' <param name="partialMatch">Should partial matches be allowed?</param>
    ''' <param name="searchByRows">Search by rows first (<c>true</c>) or columns first (<c>false</c>).</param>
    ''' <param name="searchNext">Search forward [top to bottom / left to right] (<c>true</c>) or backward [bottom to top / right to left].</param>
    ''' <param name="searchRange">Range to apply the search to.
    ''' If not provided, all cells are searched.</param>
    ''' <returns>
    ''' Single cell <see cref="Range"/> object.
    ''' If there is no match found, the <see cref="Range.[Object]"/> will have a value of null (use <see cref="Range.IsNull"/>).
    ''' </returns>
    ''' <remarks>
    ''' Excel searches will wrap so to prevent processing the same results, cell locations should be considered.
    ''' Use the <see cref="Range.CompareFirstCell"/> returned <see cref="Range"/> object.
    ''' <seealso cref="FindAll"/>
    ''' </remarks>
    Public Function Find(valueToFind As Object, Optional afterCell As Range = Nothing,
                         Optional partialMatch As Boolean = False, Optional searchByRows As Boolean = True,
                         Optional searchNext As Boolean = True, Optional searchRange As Range = Nothing) As Range
        ' xlValues = -4163
        ' xlWhole = 1 / xlPart = 2
        ' xlByRows = 1 / xlByColumns = 2
        ' xlNext = 1 / xlPrevious = 2

        _lastSearchRange = If(searchRange, GetAllCells)

        Dim xlRange As Object = _lastSearchRange.Object.Find(
            valueToFind, If(afterCell Is Nothing, Type.Missing, afterCell.Object), -4163,
            IIf(partialMatch, 2, 1), IIf(searchByRows, 1, 2), IIf(searchNext, 1, 2))

        Return New Range(xlRange)

    End Function
    ''' <summary>
    ''' Finds the next match for the last defined search parameters searching forward (top to bottom / left to right).
    ''' </summary>
    ''' <param name="afterCell">Non-inclusive cell to continue the search from.</param>
    ''' <returns>
    ''' Single cell <see cref="Range"/> object.
    ''' If there is no match found, the <see cref="Range.[Object]"/> will have a value of null (use <see cref="Range.IsNull"/>).
    ''' </returns>
    ''' <remarks>
    ''' Excel searches will wrap so prevent processing the same results, cell locations should be considered.
    ''' Use the <see cref="Range.CompareFirstCell"/> returned <see cref="Range"/> object.
    ''' <para>This should be used after the <see cref="Find"/> function which sets the respective search range.</para>
    ''' </remarks>
    Public Function FindNext(afterCell As Range) As Range
        Dim xlRange As Object = If(_lastSearchRange, GetAllCells).Object.FindNext(afterCell.Object)
        Return New Range(xlRange)
    End Function
    ''' <summary>
    ''' Finds the next match for the last defined search parameters searching backward (bottom to top / right to left).
    ''' </summary>
    ''' <param name="beforeCell">Non-inclusive cell to continue the search from.</param>
    ''' <returns>
    ''' Single cell <see cref="Range"/> object.
    ''' If there is no match found, the <see cref="Range.[Object]"/> will have a value of null (use <see cref="Range.IsNull"/>).
    ''' <para>This should be used after the <see cref="Find"/> function which sets the respective search range.</para>
    ''' </returns>
    ''' <remarks>
    ''' Excel searches will wrap so prevent processing the same results, cell locations should be considered.
    ''' Use the <see cref="Range.CompareFirstCell"/> returned <see cref="Range"/> object.
    ''' </remarks>
    Public Function FindPrevious(beforeCell As Range) As Range
        Dim xlRange As Object = If(_lastSearchRange, GetAllCells).Object.FindPrevious(beforeCell.Object)
        Return New Range(xlRange)
    End Function

    ''' <summary>
    ''' Returns all cells which match the search parameters.
    ''' If no results are found, an empty array is returned.
    ''' </summary>
    ''' <param name="valueToFind">Value to find.</param>
    ''' <param name="partialMatch">Should partial matches be allowed?</param>
    ''' <param name="searchRange">Range to apply the search to.
    ''' If not provided, all cells are searched.</param>
    ''' <returns>Array of <see cref="Range"/> objects. If no results are found, the array is empty.</returns>
    ''' <remarks>This is wrapper for the <see cref="Find"/> function and uses the default parameters of that method with regards to the traversal.</remarks>
    Public Function FindAll(valueToFind As Object, Optional partialMatch As Boolean = False,
                            Optional searchRange As Range = Nothing) As Range()

        searchRange = If(searchRange, GetAllCells)

        Dim results As New List(Of Range)
        Dim result = Find(valueToFind, partialMatch:=partialMatch, searchRange:=searchRange)

        ' Function for evaluating if the find result already exists in the range.
        ' The value to check cannot be empty.
        Dim checkForExistance = Function(check As Range)
                                    For Each x In results
                                        If result.CompareFirstCell(x) Then
                                            Return True
                                        End If
                                    Next
                                    Return False
                                End Function

        While Not result.IsNull AndAlso Not checkForExistance(result)
            ' If the result is the first search cell then put it as the first result.
            If result.CompareFirstCell(searchRange) Then
                results.Insert(0, result)
            Else
                results.Add(result)
            End If

            result = FindNext(result)
        End While

        Return results.ToArray

    End Function

    ''' <summary>
    ''' Sorts columns by value.
    ''' </summary>
    ''' <param name="sortColumns">Index of columns and sort order in the order to apply them.</param>
    ''' <param name="hasHeaders">Specifies if the <paramref name="sortRange"/> contains a header row.</param>
    ''' <param name="sortRange">Range to apply the sort to. If left empty, <see cref="GetAllCells"/> is used.</param>
    ''' <remarks></remarks>
    Public Sub Sort(sortColumns As Dictionary(Of Object, SortDirection),
                    Optional hasHeaders As Boolean = True, Optional sortRange As Range = Nothing)

        With Me.Object.Sort
            With .SortFields
                .Clear()
                For Each sortColumn In sortColumns
                    .Add(Key:=GetColumn(sortColumn.Key).Object, Order:=sortColumn.Value, SortOn:=0) 'xlSortOnValues
                Next
            End With

            .SetRange(If(sortRange Is Nothing, GetAllCells, sortRange).Object)
            .Header = IIf(hasHeaders, 1, 2)
            .MatchCase = False
            .Orientation = 1 'xlTopToBottom
            .SortMethod = 1 'xlPinYin
            .Apply()
        End With

    End Sub

    ''' <summary>
    ''' Pastes the current clipboard contents to the provided <paramref name="range"/>.
    ''' </summary>
    ''' <param name="range"><see cref="Range"/> object to receive the copied information.</param>
    ''' <remarks></remarks>
    Public Sub Paste(range As Range)
        _worksheet.Paste(range.Object)
    End Sub

    ''' <summary>
    ''' Embeds the specified image file into the sheet at the specified cell location.
    ''' The image is will be saved inside of the Excel sheet.
    ''' </summary>
    ''' <param name="pictureFile">Image file to insert into the sheet.</param>
    ''' <param name="location">Cell where this image should be placed.
    ''' The object will be placed in the upper left corner of the provided location.</param>
    ''' <param name="sizeProportional">Specifies if resizing should be done preserving the aspect ratio.</param>
    ''' <param name="widthFixed">Scales the image width, proportionally, to the fixed width.</param>
    ''' <param name="heightFixed">Scales the image height, proportionally, to the fixed height.</param>
    ''' <param name="widthPercent">Scales the image width, proportionally, to the percent width.</param>
    ''' <param name="heightPercent">Scales the image height, proportionally, to the percent height.</param>
    ''' <remarks>When <paramref name="sizeProportional"/> is <c>True</c>, priority for scaling is as follows (the first which provides a non-zero value is used):
    ''' <list type="number">
    ''' <item><paramref name="widthFixed"/></item>
    ''' <item><paramref name="heightFixed"/></item>
    ''' <item><paramref name="widthPercent"/></item>
    ''' <item><paramref name="heightPercent"/></item>
    ''' </list></remarks>
    Public Sub InsertPicture(pictureFile As String, location As Range,
                             Optional sizeProportional As Boolean = True,
                             Optional widthFixed As Integer = 0, Optional heightFixed As Integer = 0,
                             Optional widthPercent As Single = 0, Optional heightPercent As Single = 0)

        If Not My.Computer.FileSystem.FileExists(pictureFile) Then
            Exit Sub
        End If

        ' Read the image dimensions.
        Dim image As Bitmap = New Bitmap(pictureFile)

        With _worksheet.Shapes.AddPicture(pictureFile,
                                          Constants.MsoTriState.False,
                                          Constants.MsoTriState.True,
                                          1, 1,
                                          image.Width, image.Height)

            If sizeProportional Then
                .LockAspectRatio = Constants.MsoTriState.True

                If widthFixed <> 0 Then
                    .Width = widthFixed
                ElseIf heightFixed <> 0 Then
                    .Height = heightFixed
                ElseIf widthPercent <> 0 Then
                    .ScaleWidth(Convert.ToSingle(widthPercent / 100.0), Constants.MsoTriState.True)
                ElseIf heightPercent <> 0 Then
                    .ScaleHeight(Convert.ToSingle(heightPercent / 100.0), Constants.MsoTriState.True)
                End If

            Else
                .LockAspectRatio = Constants.MsoTriState.False

                If widthFixed <> 0 Or heightFixed <> 0 Then
                    If widthFixed <> 0 Then
                        .Width = widthFixed
                    End If
                    If heightFixed <> 0 Then
                        .Height = heightFixed
                    End If

                ElseIf widthPercent <> 0 Or heightPercent <> 0 Then
                    If widthPercent <> 0 Then
                        .ScaleWidth(Convert.ToSingle(widthPercent / 100.0), Constants.MsoTriState.True)
                    End If
                    If heightPercent <> 0 Then
                        .ScaleHeight(Convert.ToSingle(heightPercent / 100.0), Constants.MsoTriState.True)
                    End If

                End If
            End If

            .Left = location.Object.Left
            .Top = location.Object.Top
        End With
    End Sub


    ''' <summary>
    ''' Activates/brings to front the respective Excel worksheet.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetActive()
        _worksheet.Activate()
    End Sub

    ''' <summary>
    ''' Deletes the current Excel worksheet object.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Delete()
        _worksheet.Delete()
        _worksheet = Nothing
        _workbook.CleanWorksheets()
    End Sub

    ''' <summary>
    ''' Generates a lookup of column numbers by (text) value on the respective row.
    ''' This allows for effectively retrieving column values based on a header text location.
    ''' </summary>
    ''' <param name="row">Row number to generate the lookup for.</param>
    ''' <param name="trimText">Specifies if the header value should be trimmed before adding to the lookup.</param>
    ''' <param name="toUpper">Specifies if the header value should be converted to upper case before adding to the lookup.</param>
    ''' <returns>Key value pair with the column text as the key and the column number as the value.</returns>
    ''' <remarks>In the event of duplicate header text (consider the transforms applied), only the first occurance is added to the lookup.</remarks>
    Public Function GetHeaderIndexes(Optional row As Integer = 1,
                                     Optional trimText As Boolean = True, Optional toUpper As Boolean = False) As Dictionary(Of String, Integer)

        Dim headerIndexes = New Dictionary(Of String, Integer)

        For columnNumber = 1 To GetMaxColumn(row)
            Dim value = TryCast(GetValue(row, columnNumber), String)
            If value Is Nothing Then
                Continue For
            End If

            value = value.ToString

            If trimText Then
                value = value.Trim
            End If
            If toUpper Then
                value = value.ToUpper
            End If

            If Not headerIndexes.ContainsKey(value) Then
                headerIndexes.Add(value, columnNumber)
            End If
        Next

        Return headerIndexes

    End Function

#End Region


End Class
