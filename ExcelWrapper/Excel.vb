''' <summary>
''' Late-bound object wrapper to the Excel.Application COM object.
''' </summary>
''' <remarks>
''' Because this library is late-bound, it does not require a project reference to the Excel COM object libraries.
''' However, this component requires that Excel is installed on any machine which it is utilized.
''' </remarks>
Public Class Excel
    Implements IDisposable


#Region " Instance "

    ''' <summary>
    ''' Raw Excel.Application object used by this wrapper.
    ''' </summary>
    ''' <remarks></remarks>
    Private _excel As Object

    ''' <summary>
    ''' Internal list of all <see cref="Workbook"/>s which have been explicitly referenced.
    ''' </summary>
    ''' <remarks></remarks>
    Private _workbooks As List(Of Workbook)



    ''' <summary>
    ''' Creates a new Excel.Application object.
    ''' </summary>
    ''' <remarks>The application is created invisibily with no workbooks open.</remarks>
    Public Sub New()

        _excel = CreateObject("Excel.Application")
        _workbooks = New List(Of Workbook)

        _excel.DisplayAlerts = False

        Me.Visible = False

        ' Close any existing workbooks.
        For i As Integer = _excel.Workbooks.Count To 1 Step -1
            _excel.Workbooks(i).Close()
        Next

    End Sub

#End Region


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing AndAlso _excel IsNot Nothing Then
                ' If the Excel object is still alive and hidden, show it before releasing.
                If Not Visible Then
                    Visible = True
                End If

                ReleaseAllWorkbooks()
                _excel.DisplayAlerts = True
                _excel = Nothing
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region


#Region " Properties "

    ''' <summary>
    ''' Gets the raw Excel.Application object.
    ''' </summary>
    ''' <returns>Excel.Application object.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property [Object] As Object
        Get
            Return _excel
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets the status bar text displayed along the bottom of the Excel window.
    ''' </summary>
    ''' <value>New text value.</value>
    ''' <returns>Existing text value.</returns>
    ''' <remarks>This method will not reset the status bar if an empty value is passed.
    ''' To reset the status bar, use the <see cref="ToggleStatusBar"/> method.</remarks>
    Public Property StatusText As String
        Get
            Return _excel.StatusBar
        End Get
        Set(value As String)
            _excel.StatusBar = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the visibility of Excel application instance.
    ''' </summary>
    ''' <value>New visibility status.</value>
    ''' <returns>Current visibility status.</returns>
    ''' <remarks></remarks>
    Public Property Visible As Boolean
        Get
            Return _excel.Visible
        End Get
        Set(value As Boolean)
            _excel.Visible = value
        End Set
    End Property

    ''' <summary>
    ''' Gets the total number of Excel.Workbooks contained in the Excel application.
    ''' </summary>
    ''' <returns>Total number of Excel workbooks.</returns>
    ''' <remarks>This does not necessarily equal the internal object which contains <see cref="Workbook"/>s.</remarks>
    Public ReadOnly Property WorkbookCount As Integer
        Get
            Return _excel.Workbooks.Count
        End Get
    End Property

#End Region


#Region " Methods "

    ''' <summary>
    ''' Creates and returns a new <see cref="Workbook"/> object.
    ''' </summary>
    ''' <returns>New <see cref="Workbook"/> instance.</returns>
    ''' <remarks></remarks>
    Public Function AddWorkbook() As Workbook

        CleanWorkbooks()

        Dim xlWorkbook As Object = _excel.Workbooks.Add
        Dim newWorkbook As New Workbook(Me, xlWorkbook)

        ' Add to the internal object before returning.
        _workbooks.Add(newWorkbook)

        Return newWorkbook
    End Function

    ''' <summary>
    ''' Opens the specified <paramref name="fileName"/> and returns the <see cref="Workbook"/> object.
    ''' </summary>
    ''' <param name="fileName">File to open in Excel.</param>
    ''' <param name="password">Password required to open the file (if applicable.</param>
    ''' <returns><see cref="Workbook"/> instance for the newly opened file.</returns>
    ''' <remarks></remarks>
    Public Function OpenWorkbook(fileName As String, Optional password As String = "") As Workbook

        CleanWorkbooks()

        Dim xlWorkbook As Object = _excel.Workbooks.Open(fileName, Password:=password)
        Dim newWorkbook As New Workbook(Me, xlWorkbook, existingFileName:=fileName)

        ' Add to the internal object before returning.
        _workbooks.Add(newWorkbook)

        Return newWorkbook

    End Function

    ''' <summary>
    ''' Returns a <see cref="Workbook"/> object for the requested <paramref name="reference"/>.
    ''' </summary>
    ''' <param name="reference">Excel workbook name or index number.</param>
    ''' <returns><see cref="Workbook"/> object.</returns>
    ''' <remarks>
    ''' If the specified <paramref name="reference"/> does not exist, an exception will be thrown.
    ''' Use <see cref="WorkbookExists"/> to verify if <paramref name="reference"/> does exist.
    ''' </remarks>
    Public Function GetWorkbook(reference As Object) As Workbook

        CleanWorkbooks()

        Dim xlWorkbook As Object = _excel.Workbooks(reference)


        ' Check if the worksheet is already in the "known" list.
        Dim existingWorkbook As Workbook = _workbooks.Find(Function(x) x.Name = xlWorkbook.Name)

        If existingWorkbook Is Nothing Then
            ' Isn't in the list.
            Dim newWorkbook As New Workbook(Me, xlWorkbook)
            _workbooks.Add(newWorkbook)
            Return newWorkbook
        Else
            ' Already in the list.
            Return existingWorkbook
        End If

    End Function


    ''' <summary>
    ''' Returns whether or not an Excel.Workbook object at the given <paramref name="reference"/> exists.
    ''' </summary>
    ''' <param name="reference">Workbook name or index.</param>
    ''' <returns></returns>
    ''' <remarks>This check is performed against the raw Excel.Application object.</remarks>
    Public Function WorkbookExists(reference As Object) As Boolean

        Try
            ' To check, attempt to use the reference.
            Dim check As Object = _excel.Workbooks(reference)

            ' If we get here, all is well.
            Return True

        Catch ex As Exception
            ' Any exception indicates it does not exist.
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Initalizes/resets the status bar displayed along the bottom of the Excel window.
    ''' If neither <paramref name="forceOn"/> or <paramref name="forceOff"/> are provided, display is toggled.
    ''' </summary>
    ''' <param name="forceOn">Specifies that the status bar should be reset to the default state and displayed.
    ''' If this value is provided, it trumps <paramref name="forceOff"/>, if provided.</param>
    ''' <param name="forceOff">Specifies the status bar should be reset and hidden.</param>
    ''' <remarks></remarks>
    Public Sub ToggleStatusBar(Optional forceOn As Boolean = False, Optional forceOff As Boolean = False)

        _excel.StatusBar = False

        _excel.DisplayStatusBar =
            IIf(forceOn, True,
                IIf(forceOff, False,
                    Not _excel.DisplayStatusBar))

    End Sub


    ''' <summary>
    ''' Closes the current Excel application.
    ''' </summary>
    ''' <remarks>Any linked <see cref="Workbook"/> and subsequent <see cref="Worksheet"/> objects will be released/set to null.</remarks>
    Public Sub Close()

        ReleaseAllWorkbooks()
        _excel.Quit()
        _excel = Nothing

    End Sub


    ''' <summary>
    ''' Removes any dead/empty <see cref="Workbook"/> objects from the internal list.
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub CleanWorkbooks()

        For i As Integer = _workbooks.Count - 1 To 0 Step -1
            If _workbooks(i).IsNull Then
                _workbooks.RemoveAt(i)
            End If
        Next

    End Sub

    ''' <summary>
    ''' Sets any <see cref="Workbook"/> objects linked to this object to null.
    ''' </summary>
    ''' <remarks>This action will also trigger subsequently linked <see cref="Worksheet"/> objects to be released.</remarks>
    Private Sub ReleaseAllWorkbooks()

        For Each workbook As Workbook In _workbooks
            workbook = Nothing
        Next

        _workbooks.Clear()

    End Sub

#End Region


End Class
