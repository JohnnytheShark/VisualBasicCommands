Public Sub ImportFiles()
' Imports the first sheet from every workbook that user selected.
    Dim TextFile As Workbook
    Dim OpenFiles() As Variant
    Dim i As Integer
    Dim SheetCount As Integer
    Dim CurrentBook As Workbook
    
    Set CurrentBook = ActiveWorkbook
    
    OpenFiles = GetFiles()
    Application.ScreenUpdating = False

    For i = 1 To Application.CountA(OpenFiles)
        Set TextFile = Workbooks.Open(OpenFiles(i))
        SheetCount = TextFile.Sheets.Count

        For j = 1 To SheetCount
            TextFile.Sheets(j).Range("A1").CurrentRegion.Copy
            CurrentBook.Activate
            CurrentBook.Worksheets.Add
            ActiveSheet.Paste
            ActiveSheet.Name = TextFile.Sheets(j).Name
        Next j

        Application.CutCopyMode = False

        TextFile.Close
    Next i
    Application.ScreenUpdating = True


End Sub

Public Function GetFiles() As Variant
    GetFiles = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=True)
End Function