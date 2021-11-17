Public Sub ImportFiles()
    Dim TextFile As Workbook
    Dim OpenFiles() As Variant
    Dim i As Integer
    Dim CurrentBook As Workbook
    
    Set CurrentBook = ActiveWorkbook
    
    OpenFiles = GetFiles()
    Application.ScreenUpdating = False

    For i = 1 To Application.CountA(OpenFiles)
        Set TextFile = Workbooks.Open(OpenFiles(i))

        TextFile.Sheets(1).Range("A1").CurrentRegion.Copy
        CurrentBook.Activate
        CurrentBook.Worksheets.Add
        ActiveSheet.Paste
        ActiveSheet.Name = TextFile.Sheets(1).Name

        Application.CutCopyMode = False

        TextFile.Close
    Next i
    Application.ScreenUpdating = True


End Sub

Public Function GetFiles() As Variant
    GetFiles = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=True)
End Function

