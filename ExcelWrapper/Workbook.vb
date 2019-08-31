''' <summary>
''' Late-bound object wrapper to the Excel.Workbook COM object.
''' </summary>
''' <remarks></remarks>
Public Class Workbook


#Region " Instance "

    ''' <summary>
    ''' Parent instance of <see cref="Excel"/>.
    ''' </summary>
    ''' <remarks></remarks>
    Private ReadOnly _excel As Excel

    ''' <summary>
    ''' Raw Excel.Workbook object used by this wrapper.
    ''' </summary>
    ''' <remarks></remarks>
    Private _workbook As Object

    ''' <summary>
    ''' Internal list of all <see cref="Worksheet"/>s which have been explicitly referenced.
    ''' </summary>
    ''' <remarks></remarks>
    Private _worksheets As List(Of Worksheet)

    ''' <summary>
    ''' File name of the <see cref="_workbook"/>, if applicable.
    ''' </summary>
    ''' <remarks></remarks>
    Private _fileName As String

    ''' <summary>
    ''' When a new Excel.Workbook is created, this is the reference for the Excel.Worksheet.
    ''' It is automatically deleted by the <see cref="AddWorksheet"/> unless it is explicitly referenced first via <see cref="GetWorksheet"/>.
    ''' </summary>
    ''' <remarks>When an existing Excel.Workbook is opened, this is not used.</remarks>
    Private _defaultWorksheet As Object = Nothing



    ''' <summary>
    ''' Creates a new Excel.Workbook latebound object.
    ''' </summary>
    ''' <param name="owner">Parent <see cref="Excel"/> object.</param>
    ''' <param name="excelWorkbook">Excel.Workbook object for this instance.</param>
    ''' <param name="existingFileName">Excel workbook file name for the respective Excel.Workbook object.</param>
    ''' <remarks>
    ''' The constructor is internal only to enforce using the <see cref="Excel"/> object.
    ''' 
    ''' When a new workbook is created, it is created with only a single sheet which is marked for deletion
    ''' (Excel requires a workbook to have at least one sheet).
    ''' This sheet is not really intended to be used (as it is expected any sheets needed will be explicitly created),
    ''' so it will be automatically removed when a new sheet is added unless a <see cref="Worksheet"/> object reference
    ''' is linked to it first (using the <see cref="GetWorksheet"/> method).
    ''' </remarks>
    Friend Sub New(ByRef owner As Excel, ByRef excelWorkbook As Object, Optional existingFileName As String = "")

        _excel = owner
        _workbook = excelWorkbook
        _worksheets = New List(Of Worksheet)
        _fileName = existingFileName

        If existingFileName = "" Then
            ' Clear any existing worksheets excel for the first.
            ' Cannot remove all worksheets because Excel requires there to be one.
            For i As Integer = _workbook.Worksheets.Count To 2 Step -1
                _workbook.Sheets(i).Delete()
            Next

            ' The remaining worksheet will be cleaned up automatically when the user creates a sheet.

            ' Set a simple marker to note this sheet should be deleted.
            _defaultWorksheet = _workbook.Sheets(1)
        End If

    End Sub

#End Region


#Region " Properties "

    ''' <summary>
    ''' Gets the owning <see cref="Excel"/> object.
    ''' </summary>
    ''' <returns><see cref="Excel"/> object which contains the current instance.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Excel As Excel
        Get
            Return _excel
        End Get
    End Property

    ''' <summary>
    ''' Gets the raw Excel.Workbook object.
    ''' </summary>
    ''' <returns>Excel.Workbook object.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property [Object] As Object
        Get
            Return _workbook
        End Get
    End Property

    ''' <summary>
    ''' Returns whether the Excel.Workbook object is empty/null.
    ''' </summary>
    ''' <returns>Boolean indicating if the Excel.Workbook object is empty.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsNull As Boolean
        Get
            Return _workbook Is Nothing
        End Get
    End Property

    ''' <summary>
    ''' Gets the name of the workbook. Note this can be different from the <see cref="FileName"/>.
    ''' </summary>
    ''' <returns>Name of the Excel.Workbook object.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Name As String
        Get
            Return _workbook.Name
        End Get
    End Property

    ''' <summary>
    ''' Gets the file name of the respective workbook, if it has been saved.
    ''' </summary>
    ''' <value></value>
    ''' <returns>File name of the current workbook, or an empty string if it has not been saved.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileName As String
        Get
            Return _fileName
        End Get
    End Property

    ''' <summary>
    ''' Gets the total number of Excel.Worksheets contained in the Excel workbook.
    ''' </summary>
    ''' <returns>Total number of Excel worksheets.</returns>
    ''' <remarks>This does not necessarily equal the internal object which contains <see cref="Worksheet"/>s.</remarks>
    Public ReadOnly Property WorksheetCount As Integer
        Get
            Return _workbook.Sheets.Count
        End Get
    End Property

#End Region


#Region " Methods "

    ''' <summary>
    ''' Adds and returns a new <see cref="Worksheet"/> to the current object. 
    ''' </summary>
    ''' <param name="name">Name of the new <see cref="Worksheet"/>.</param>
    ''' <param name="before"><see cref="Worksheet"/> to add the new sheet before.</param>
    ''' <param name="after"><see cref="Worksheet"/> to add the new sheet after.</param>
    ''' <param name="toEnd">When True, the new worksheet is added as the last sequentially.</param>
    ''' <returns>New <see cref="Worksheet"/> instance.</returns>
    ''' <remarks>
    ''' Precedence to the location of the new worksheet is as follows:
    ''' <paramref name="before"/>, <paramref name="after"/>, <paramref name="toEnd"/>
    ''' </remarks>
    Public Function AddWorksheet(Optional name As String = "",
                                 Optional before As Worksheet = Nothing,
                                 Optional after As Worksheet = Nothing,
                                 Optional toEnd As Boolean = True) As Worksheet

        CleanWorksheets()

        Dim newWorksheet As Worksheet = Nothing

        ' Check if new sheet being requested already exists.
        If WorksheetExists(name) Then
            ' Already exists, set the reference.
            newWorksheet = GetWorksheet(name)

        Else
            ' Does not exist, create it.

            Dim xlWorksheet As Object
            If before IsNot Nothing Then
                xlWorksheet = _workbook.Sheets.Add(before.Object)
            ElseIf after IsNot Nothing Then
                xlWorksheet = _workbook.Sheets.Add(, after.Object)
            ElseIf toEnd Then
                xlWorksheet = _workbook.Sheets.Add(, _workbook.Sheets(WorksheetCount))
            Else
                xlWorksheet = _workbook.Sheets.Add()
            End If
            newWorksheet = New Worksheet(Me, xlWorksheet, name:=name)

            ' Add to the internal object before returning.
            _worksheets.Add(newWorksheet)
        End If


        If _defaultWorksheet IsNot Nothing Then
            ' The original sheet which is marked for deletion is still in the workbook.
            ' See what action should be done with it.

            ' Check if the original sheet is being used.
            If _worksheets.Find(Function(x) x.Name = _defaultWorksheet.Name) Is Nothing Then
                ' Original sheet is not being used, delete it.
                _defaultWorksheet.Delete()
            End If

            ' No more action should be done with the original sheet.
            _defaultWorksheet = Nothing
        End If


        Return newWorksheet

    End Function

    ''' <summary>
    ''' Moves the <paramref name="sheetToMove"/> within the Workbook.
    ''' </summary>
    ''' <param name="sheetToMove"><see cref="Worksheet"/> object to move.</param>
    ''' <param name="moveBefore">
    ''' <see cref="Worksheet"/> object to move <paramref name="sheetToMove"/> before.
    ''' If this is specified, do not provide a value for <paramref name="moveAfter"/>.
    ''' </param>
    ''' <param name="moveAfter">
    ''' <see cref="Worksheet"/> object to move <paramref name="sheetToMove"/> after.
    ''' If this is specified, do not provide a value for <paramref name="moveBefore"/>.
    ''' </param>
    ''' <remarks>
    ''' In the event neither <paramref name="moveBefore"/> or <paramref name="moveAfter"/>, no action is taken.
    ''' If both <paramref name="moveBefore"/> and <paramref name="moveAfter"/> are provide, <paramref name="moveBefore"/> takes priority.
    ''' </remarks>
    Public Sub MoveWorksheet(sheetToMove As Worksheet,
                             Optional moveBefore As Worksheet = Nothing, Optional moveAfter As Worksheet = Nothing)

        If moveBefore IsNot Nothing Then
            _workbook.Sheets(sheetToMove.Name).Move(moveBefore.Object)

        ElseIf moveAfter IsNot Nothing Then
            _workbook.Sheets(sheetToMove.Name).Move(, moveAfter.Object)

        End If

    End Sub

    ''' <summary>
    ''' Returns a <see cref="Worksheet"/> object for the requested <paramref name="reference"/>.
    ''' </summary>
    ''' <param name="reference">Excel worksheet name or index number.</param>
    ''' <returns><see cref="Worksheet"/> object.</returns>
    ''' <remarks>
    ''' If the specified <paramref name="reference"/> does not exist, an exception will be thrown.
    ''' Use <see cref="WorksheetExists"/> to verify if <paramref name="reference"/> does exist.
    ''' </remarks>
    Public Function GetWorksheet(reference As Object) As Worksheet

        CleanWorksheets()

        Dim xlWorksheet As Object = _workbook.Sheets(reference)

        ' Check if the worksheet is already in the "known" list.
        Dim existingWorksheet As Worksheet = _worksheets.Find(Function(x) x.Name = xlWorksheet.Name)

        If existingWorksheet Is Nothing Then
            ' Isn't in the list.
            Dim newWorksheet As New Worksheet(Me, xlWorksheet)
            _worksheets.Add(newWorksheet)
            Return newWorksheet
        Else
            ' Already in the list.
            Return existingWorksheet
        End If

    End Function


    ''' <summary>
    ''' Returns whether or not an Excel.Worksheet object at the given <paramref name="reference"/> exists.
    ''' </summary>
    ''' <param name="reference">Worksheet name or index.</param>
    ''' <returns></returns>
    ''' <remarks>This check is performed against the raw Excel.Workbook object.</remarks>
    Public Function WorksheetExists(reference As Object) As Boolean

        Try
            ' To check, attempt to use the reference.
            Dim check As Object = _workbook.Sheets(reference)

            ' If we get here, all is well.
            Return True

        Catch ex As Exception
            ' Any exception indicates it does not exist.
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Saves the current Excel workbook to disk.
    ''' </summary>
    ''' <param name="fileName">Output file name. This will set <see cref="FileName"/>.</param>
    ''' <param name="format">Format to write the workbook as. Any 'XlFileFormat' value is accepted.
    '''                      If <see cref="SaveFormat.AutoDetermine"/> is used,
    '''                      then filenames ending in 'xls' or 'csv' are automatically saved in their respective format
    '''                      with everything else using <see cref="SaveFormat.Default"/>. </param>
    ''' <param name="password">When provided, sets the protection password on the file.
    '''                        This is only applicable for native Excel formats. The Excel limit is 15 chars.</param>
    ''' <remarks></remarks>
    Public Sub Save(fileName As String,
                    Optional format As SaveFormat = SaveFormat.AutoDetermine,
                    Optional password As String = "")
        _fileName = fileName

        If format = SaveFormat.AutoDetermine Then
            Select Case System.IO.Path.GetExtension(fileName).ToLower
                Case ".csv"
                    format = SaveFormat.CSV
                Case ".xls"
                    format = SaveFormat.Legacy
                Case Else
                    format = SaveFormat.Default
            End Select
        End If

        _workbook.SaveAs(fileName, FileFormat:=format, Password:=password)
    End Sub

    ''' <summary>
    ''' Closes the current Excel workbook.
    ''' </summary>
    ''' <remarks>Any linked <see cref="Worksheet"/> objects will be released/set to null.</remarks>
    Public Sub Close()

        ReleaseAllWorksheets()
        _workbook.Close()
        _workbook = Nothing
        Excel.CleanWorkbooks()

    End Sub

    ''' <summary>
    ''' Activates/brings to front the respective Excel workbook.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetActive()

        _workbook.Activate()

    End Sub


    ''' <summary>
    ''' Removes any dead/empty <see cref="Worksheet"/> objects from the internal list.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub CleanWorksheets()

        For i As Integer = _worksheets.Count - 1 To 0 Step -1
            If _worksheets(i).IsNull Then
                _worksheets.RemoveAt(i)
            End If
        Next

    End Sub

    ''' <summary>
    ''' Sets any <see cref="Worksheet"/> objects linked to this object to null.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReleaseAllWorksheets()

        For Each Worksheet As Worksheet In _worksheets
            Worksheet = Nothing
        Next

        _worksheets.Clear()

    End Sub

#End Region


End Class
