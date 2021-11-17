Attribute VB_Name = "Module5" ' Name of the module when imported into excel. 
Public Sub Formatting()
' This macro doesn't do anything. However I just wanted to place in code that was used on previous projects regarding formatting.
Cells.Select
Cells.EntireColumn.AutoFit
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select 'Basically go from A1 to the end of the table. Can always be changed out to another row/column
    Selection.Interior.Color = RGB(30, 102, 40) ' RGB is easy to look up on google
    Selection.Font.Color = RGB(255, 255, 255)
    Columns("N:N").Select ' Select a column fully
    Selection.Insert Shift:=xlToRight 'Insert a column to the left of it
    Selection.ClearFormats ' Clear any formats

    'Example of going to the bottom of a table and offsetting one column and then filling the interior. (This is sort of like adding a border but deeper color and changing width)
    Range("M1").Select ' Select cell first
    Selection.End(xlDown).Select ' Go to the bottom row of this. BE WARE EMPTY ROWS SUCH AS NULLS CAN THROW THIS OFF. CHOSE A COLUMN THAT ALWAYS HAS DATA IN IT
    ActiveCell.Offset(0, 1).Select ' OFFset to the column you want 
    Range(Selection, Selection.End(xlUp)).Select ' Go to the top of that column selecting everything within it.
    With Selection.Interior ' Set the interior to a dark Color (this can always be done with RGB but decided to show a different way (basically how excel records the macros))
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Columns("N:N").ColumnWidth = 2 ' Setting the column width to 2 (not pixels believe it is a different measurement)
End Sub