// Usage of this code is through the MIT License. See LICENSE file.

namespace ExcelWrapper;

/// <summary> Container for holding Excel constant values. </summary>
[System.Diagnostics.CodeAnalysis.SuppressMessage("CodeQuality", "IDE0079:Remove unnecessary suppression", Justification = "Ignore ReSharper flags")]
[System.Diagnostics.CodeAnalysis.SuppressMessage("ReSharper", "MissingXmlDoc", Justification = "Passthrough values")]
#pragma warning disable CS1591
public static class Constants
{
    public enum SaveFormat
    {
        AutoDetermine = int.MinValue,
        CSV = 6,
        Default = 51,
        Legacy = 39,
    }

    public enum BorderEdge
    {
        Top = 8,
        Bottom = 9,
        Left = 7,
        Right = 10,
    }

    public enum BorderStyle
    {
        None = -4142,
        Single = 1,
        Double = -4119,
        Dash = -4115,
        Dot = -4118,
    }

    public enum BorderWeight
    {
        Normal = -4138,
        Hairline = 1,
        Thin = 2,
        Thick = 4,
    }

    public enum HorizontalAlignment
    {
        Left = -4131,
        Right = -4152,
        Center = -4108,
    }

    public enum VerticalAlignment
    {
        Top = -4160,
        Bottom = -4107,
        Middle = -4108,
    }

    public enum SortDirection
    {
        Ascending = 1,
        Descending = 2,
    }

    public enum MsoTriState
    {
        False = 0,
        True = -1,
    }
}