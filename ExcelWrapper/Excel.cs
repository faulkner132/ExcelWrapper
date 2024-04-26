// Usage of this code is through the MIT License. See LICENSE file.

namespace ExcelWrapper;

/// <summary> Late-bound object wrapper to the <c>Excel.Application</c> COM object. </summary>
///
/// <remarks>
/// Because this library is late-bound, it does not require a project reference to the Excel COM object libraries.
/// However, this component requires that Excel is installed on any machine which it is utilized.
/// </remarks>
public class Excel : IDisposable
{
    /// <summary> Internal list of all <see cref="Workbook"/>s which have been explicitly referenced. </summary>
    private List<Workbook> Workbooks { get; } = [];

    /// <summary> Creates a new <c>Excel.Application</c> object. </summary>
    ///
    /// <remarks> The application is created invisibly with no workbooks open. </remarks>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("CodeQuality", "IDE0079:Remove unnecessary suppression", Justification = "Only supported on Windows")]
    public Excel()
    {
#pragma warning disable CA1416
        RawObject = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!;
#pragma warning restore CA1416

        RawObject.DisplayAlerts = false;

        Visible = false;

        // Close any existing workbooks.
        for (int i = Convert.ToInt32(RawObject.Workbooks.Count); i >= 1; i -= 1)
        {
            RawObject.Workbooks(i).Close();
        }
    }

    #region IDisposable Support

    /// <summary> Gets or sets a value indicating whether this object has been disposed. </summary>
    private bool HasBeenDisposed { get; set; }

    /// <inheritdoc/>
    public void Dispose()
    {
        if (!HasBeenDisposed)
        {
            if (RawObject is { })
            {
                // If the Excel object is still alive and hidden, show it before releasing.
                if (!Visible)
                {
                    Visible = true;
                }

                ReleaseAllWorkbooks();
                RawObject.DisplayAlerts = true;
                RawObject = null;
            }
        }
        HasBeenDisposed = true;

        GC.SuppressFinalize(this);
    }

    #endregion

    /// <summary> Gets or sets the raw <c>Excel.Application</c> object. </summary>
    ///
    /// <value> <c>Excel.Application</c> object. </value>
    public dynamic RawObject { get; private set; }

    /// <summary> Gets or sets the status bar text displayed along the bottom of the Excel window. </summary>
    ///
    /// <remarks>
    /// This method will not reset the status bar if an empty value is passed. To reset the status bar, use the <see cref="ToggleStatusBar"/>
    /// method.
    /// </remarks>
    ///
    /// <value> New text value. </value>
    public string StatusText
    {
        get => RawObject.StatusBar.ToString();
        set => RawObject.StatusBar = value;
    }

    /// <summary> Gets or sets the visibility of Excel instance. </summary>
    ///
    /// <value> New visibility status. </value>
    public bool Visible
    {
        get => Convert.ToBoolean(RawObject.Visible);
        set => RawObject.Visible = value;
    }

    /// <summary> Gets the total number of <c>Excel.Workbooks</c> contained in the Excel application. </summary>
    ///
    /// <remarks> This does not necessarily equal the internal object which contains <see cref="Workbook"/>s. </remarks>
    ///
    /// <value> Total number of Excel workbooks. </value>
    public int WorkbookCount => Convert.ToInt32(RawObject.Workbooks.Count);

    /// <summary> Creates and returns a new <see cref="Workbook"/> object. </summary>
    ///
    /// <returns> New <see cref="Workbook"/> instance. </returns>
    public Workbook AddWorkbook()
    {
        CleanWorkbooks();

        var xlWorkbook = RawObject.Workbooks.Add;
        var newWorkbook = new Workbook(this, xlWorkbook);
        Workbooks.Add(newWorkbook);

        return newWorkbook;
    }

    /// <summary> Opens the specified <paramref name="fileName"/> and returns the <see cref="Workbook"/> object. </summary>
    ///
    /// <param name="fileName"> File to open in Excel. </param>
    /// <param name="password"> (Optional) Password required to open the file (if applicable). </param>
    ///
    /// <returns> <see cref="Workbook"/> instance for the newly opened file. </returns>
    public Workbook OpenWorkbook(string fileName, string password = "")
    {
        CleanWorkbooks();

        var xlWorkbook = RawObject.Workbooks.Open(fileName, Password: password);
        var newWorkbook = new Workbook(this, xlWorkbook, existingFileName: fileName);
        Workbooks.Add(newWorkbook);

        return newWorkbook;
    }

    /// <summary> Returns a <see cref="Workbook"/> object for the requested <paramref name="reference"/>. </summary>
    ///
    /// <remarks>
    /// If the specified <paramref name="reference"/> does not exist, an exception will be thrown. Use <see cref="WorkbookExists"/>
    /// to verify if <paramref name="reference"/> does exist.
    /// </remarks>
    ///
    /// <param name="reference"> Excel workbook name or index number (1 based). </param>
    ///
    /// <returns> <see cref="Workbook"/> object. </returns>
    public Workbook GetWorkbook(object reference)
    {
        CleanWorkbooks();

        var xlWorkbook = RawObject.Workbooks(reference);

        if (Workbooks.FirstOrDefault(workbook => workbook.Name == xlWorkbook.Name) is { } existingWorkbook)
        {
            return existingWorkbook;
        }

        var newWorkbook = new Workbook(this, xlWorkbook);
        Workbooks.Add(newWorkbook);

        return newWorkbook;
    }

    /// <summary>
    /// Returns whether an <c>Excel.Workbook</c> object at the given <paramref name="reference"/> exists.
    /// </summary>
    ///
    /// <remarks> This check is performed against the raw <c>Excel.Application</c> object. </remarks>
    ///
    /// <param name="reference"> <inheritdoc cref="GetWorkbook" path="/param[@name='reference']"/> </param>
    public bool WorkbookExists(object reference)
    {
        try
        {
            // To check, attempt to use the reference.
            _ = RawObject.Workbooks(reference);

            // If we get here, all is well.
            return true;
        }
        catch (Exception)
        {
            // Any exception indicates it does not exist.
            return false;
        }
    }

    /// <summary>
    /// Initializes/resets the status bar displayed along the bottom of the Excel window. If neither <paramref name="forceOn"/>
    /// or <paramref name="forceOff"/> are provided, display is toggled.
    /// </summary>
    ///
    /// <param name="forceOn"> (Optional) Specifies that the status bar should be reset to the default state and
    ///     displayed. If this value is provided, it takes priority over <paramref name="forceOff"/>, if provided. </param>
    /// <param name="forceOff"> (Optional) Specifies if the status bar should be reset and hidden. </param>
    public void ToggleStatusBar(bool forceOn = false, bool forceOff = false)
    {
        RawObject.StatusBar = false;
        RawObject.DisplayStatusBar = forceOn
            ? true
            : (forceOff ? false : !RawObject.DisplayStatusBar);
    }

    /// <summary> Closes the current Excel application. </summary>
    ///
    /// <remarks>
    /// Any linked <see cref="Workbook"/> and subsequent <see cref="Worksheet"/> objects will be released/set to <see langword="null"/>.
    /// </remarks>
    public void Close()
    {
        ReleaseAllWorkbooks();
        RawObject.Quit();
        RawObject = null;
    }


    /// <summary> Removes any dead/empty <see cref="Workbook"/> objects from the internal list. </summary>
    internal void CleanWorkbooks()
    {
        for (var i = Workbooks.Count - 1; i >= 0; i -= 1)
        {
            if (Workbooks[i].IsNull)
            {
                Workbooks.RemoveAt(i);
            }
        }
    }

    /// <summary> Sets any <see cref="Workbook"/> objects linked to this object to <see langword="null"/>. </summary>
    ///
    /// <remarks>
    /// This action will also trigger subsequently linked <see cref="Worksheet"/> objects to be released.
    /// </remarks>
    private void ReleaseAllWorkbooks()
    {
        for (var i = 0; i < Workbooks.Count; i++)
        {
            Workbooks[i] = default;
        }
        Workbooks.Clear();
    }
}