Public Sub ImportFiles()
    Dim TextFile As Workbook
    Dim OpenFiles() As Variant
    Dim i as Integer

    OpenFiles = Application.GetOpenFilename(title:="Select File(s) to Import", MultiSelect:=True)
    Application.ScreenUpdating = False

    For i = 1 to Application.CountA(OpenFiles)
        Set Textfile = Workbooks.Open(Openfiles(i))

        Textfile.Sheets(1).range("A1").currentregion.copy
        Workbooks(1).Activate
        Workbooks(1).Worksheets.Add
        ActiveSheet.Paste
        ActiveSheet.name = Textfile.Name

        Application.CutCopyMode = False

        TextFile.close
    Next i 
    Application.ScreenUpdating = True


End Sub