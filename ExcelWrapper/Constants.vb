''' <summary>
''' Container for holding Excel constant values.
''' </summary>
''' <remarks></remarks>
Public Module Constants

    Public Enum SaveFormat As Integer
        AutoDetermine = Integer.MinValue
        CSV = 6
        [Default] = 51
        Legacy = 39
    End Enum

    Public Enum BorderEdge As Integer
        Top = 8
        Bottom = 9
        Left = 7
        Right = 10
    End Enum

    Public Enum BorderStyle As Integer
        None = -4142
        [Single] = 1
        [Double] = --4119
        Dash = -4115
        Dot = -4118
    End Enum

    Public Enum BorderWeight As Integer
        Normal = -4138
        Hairline = 1
        Thin = 2
        Thick = 4
    End Enum

    Public Enum HorizontalAlignment As Integer
        Left = -4131
        Right = -4152
        Center = -4108
    End Enum

    Public Enum VerticalAlignment As Integer
        Top = -4160
        Bottom = -4107
        Middle = -4108
    End Enum

    Public Enum SortDirection As Integer
        Ascending = 1
        Descending = 2
    End Enum

    Public Enum MsoTriState As Integer
        [False] = 0
        [True] = -1
    End Enum

End Module
