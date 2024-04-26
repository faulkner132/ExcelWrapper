// Usage of this code is through the MIT License. See LICENSE file.

namespace ExcelWrapper;

/// <summary> Miscellaneous Excel related methods. </summary>
public class Util
{
    /// <summary> Converts a number value into the respective column letter in Excel. </summary>
    ///
    /// <remarks> This function supports evaluation up to "ZZ" despite any Excel maximums. </remarks>
    ///
    /// <param name="number"> Numeric value to get the respective column letters for. This is expected to be a positive
    ///     value. </param>
    ///
    /// <returns> String value of the respective column letters. </returns>
    public static string GetColumnLetters(int number)
    {
        var prefixCount = number / 26;

        return convertToLetter(number % 26)
            // Add the prefix if there is one.
            + (prefixCount > 0 ? convertToLetter(prefixCount) : "");

        static string convertToLetter(int letterIndex) => ((char)(('A' + letterIndex) - 1)).ToString();
    }

    /// <summary> Converts a string value into the respective column number (1 based) in Excel. </summary>
    ///
    /// <remarks> This function supports evaluation up to "ZZ" despite any Excel maximums. </remarks>
    ///
    /// <param name="letters"> String value to get the respective column number for. This is expected to be a non-empty,
    ///     trimmed value.
    ///     <para>If this is a number, it is returned as provided.</para> </param>
    ///
    /// <returns> Numeric value of the respective column. </returns>
    public static int GetColumnNumber(string letters)
    {
        if (int.TryParse(letters, out var result))
        {
            return result;
        }

        letters = letters?.ToUpper() ?? "";

        // Evaluate the right-most letter.
        // If there is only a single char, this is all we will need to do.
        var columnNumber = convertToNumber(letters[letters.Length - 1]);

        if (letters.Length > 1)
        {
            // Consider the prefix to be "26" for each letter.
            columnNumber += convertToNumber(letters[0]) * 26;
        }

        return columnNumber;

        static int convertToNumber(char letter) => (letter + 1) - 'A';
    }

    /// <summary>
    /// Breaks a cell reference into individual row and column components and returns whether the <paramref name="location"/>
    /// was processed correctly. Calculated values are returned via <see langword="out"/> parameters on success.
    /// </summary>
    ///
    /// <remarks> A return value of false will happen if an invalid <paramref name="location"/> is specified. </remarks>
    ///
    /// <param name="location"> Valid Excel cell reference, e.g. B4, G22, AC32, etc. </param>
    /// <param name="row"> [out] Calculated row number. </param>
    /// <param name="column"> [out] Calculated column number (1 based). </param>
    ///
    /// <returns> <see langword="bool"/> indicating if the <paramref name="location"/> was processed correctly. </returns>
    public static bool TryGetCellRowColumn(string location, out int row, out int column)
    {
        location = location.Trim().ToUpper();

        var index = 0;

        string getOffsetChars(Func<char, bool> condition)
        {
            var offsetChars = "";
            while ((index < location.Length) && condition(location[index]))
            {
                offsetChars += location[index].ToString();
                index++;
            }
            return offsetChars;
        }

        // Order of processing matters since 'index' is modified.
        column = GetColumnNumber(getOffsetChars(char.IsLetter));
        row = Convert.ToInt32(getOffsetChars(char.IsNumber));

        // If the index is now the same as the location length, it means all chars were processed.
        return index == location.Length;
    }

    /// <summary> Returns an Excel reference formula (e.g. R[-2]C[2]) respective to the cell locations. </summary>
    ///
    /// <param name="baseRow"> Row where the formula should be calculated from. </param>
    /// <param name="baseColumn"> Column letter or number (1 based) where the formula should be calculated from. </param>
    /// <param name="referenceRow"> Row of the cell to reference in the formula. </param>
    /// <param name="referenceColumn"> Column letter or number (1 based) of the cell to reference in the formula. </param>
    ///
    /// <returns> Excel reference formula string. </returns>
    public static string GetFormulaReference(int baseRow, object baseColumn, int referenceRow, object referenceColumn)
    {
        var rowDifference = referenceRow - baseRow;
        var columnDifference = GetColumnNumber(referenceColumn.ToString()) - GetColumnNumber(baseColumn.ToString());

        return $"R{(rowDifference != 0 ? $"[{rowDifference}]" : "")}C{(columnDifference != 0 ? $"[{columnDifference}]" : "")}";
    }
}