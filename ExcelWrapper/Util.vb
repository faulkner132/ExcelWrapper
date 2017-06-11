''' <summary>
''' Miscellanous Excel related methods.
''' </summary>
''' <remarks></remarks>
Public Class Util

    ''' <summary>
    ''' Converts a number value into the respective column letter in Excel.
    ''' </summary>
    ''' <param name="number">Numeric value to get the respective column letters for.</param>
    ''' <returns>String value of the respective column letters.</returns>
    ''' <remarks>This function supports evaluation up to "ZZ" despite any Excel maximums.</remarks>
    Public Shared Function GetColumnLetters(number As Integer) As String

        Dim prefixCount As Integer = 0
        While number > 26
            prefixCount += 1
            number -= 26
        End While

        Dim columnLetter As String = Chr(Asc("A") + number - 1)

        ' Add the prefix if there is one.
        If prefixCount > 0 Then
            columnLetter = Chr(Asc("A") + prefixCount - 1) & columnLetter
        End If

        Return columnLetter

    End Function

    ''' <summary>
    ''' Converts a string value into the respective column number in Excel.
    ''' </summary>
    ''' <param name="letters">String value to get the respective column number for.</param>
    ''' <returns>Numeric value of the respective column.</returns>
    ''' <remarks>This function supports evaluation up to "ZZ" despite any Excel maximums.</remarks>
    Public Shared Function GetColumnNumber(letters As String) As Integer

        ' Check if a number value as passed in.
        If IsNumeric(letters) Then
            Return Convert.ToInt32(letters)
        End If

        letters = letters.ToUpper

        ' Evaluate the right most letter.
        ' If there is only a single char, this is all we will need to do.
        Dim columnNumber As Integer = (Asc(Right(letters, 1)) + 1) - Asc("A")

        If letters.Length > 1 Then
            ' Consider the prefix to be "26" for each letter.
            columnNumber += ((Asc(Left(letters, 1)) + 1) - Asc("A")) * 26
        End If

        Return columnNumber

    End Function

    ''' <summary>
    ''' Breaks a cell reference into individual row and column components and returns whether the <paramref name="location"/> was processed correctly.
    ''' </summary>
    ''' <param name="location">Excel cell reference, e.g. B4, G22, AC32, etc.</param>
    ''' <param name="Row">Output variable for the respective row number.</param>
    ''' <param name="Column">Output variable the for respective column number.</param>
    ''' <returns>Boolean indicating if the <paramref name="location"/> was processed correctly.</returns>
    ''' <remarks>A return value of false will happen if an invalid <paramref name="location"/> is specified.</remarks>
    Public Shared Function GetCellRowColumn(location As String, ByRef Row As Integer, ByRef Column As Integer) As Boolean

        location = location.Trim.ToUpper

        Dim index As Integer = 0

        Dim columnChars As String = ""
        While index < location.Length AndAlso Char.IsLetter(location(index))
            columnChars &= location(index)
            index += 1
        End While

        Dim rowChars As String = ""
        While index < location.Length AndAlso Char.IsNumber(location(index))
            rowChars &= location(index)
            index += 1
        End While

        Row = Convert.ToInt32(rowChars)
        Column = GetColumnNumber(columnChars)

        ' If the index is now the same as the location lenght, it means all chars were processed.
        Return index = location.Length

    End Function

    ''' <summary>
    ''' Returns an Excel reference formula (e.g. R[-2]C[2]) respective to the cell locations.
    ''' </summary>
    ''' <param name="baseRow">Row where the formula should be calculated from.</param>
    ''' <param name="baseColumn">Column where the formula should be calculated from.</param>
    ''' <param name="referenceRow">Row of the cell to reference in the formula.</param>
    ''' <param name="referenceColumn">Column of the cell to reference in the formula.</param>
    ''' <returns>Excel reference formula string.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetFormulaReference(baseRow As Integer, baseColumn As Object, referenceRow As Integer, referenceColumn As Object) As String

        Dim rowDifference As Integer = referenceRow - baseRow
        Dim columnDifference As Integer = GetColumnNumber(referenceColumn) - GetColumnNumber(baseColumn)

        Return String.Format("R{0}C{1}",
                             If(rowDifference <> 0, String.Format("[{0}]", rowDifference), ""),
                             If(columnDifference <> 0, String.Format("[{0}]", columnDifference), ""))

    End Function

End Class
