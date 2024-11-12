Sub ChoosingFile()
''Declarations''
    Dim myfile As String, myfile2 As String, FName As String
    Dim myfilearray As FileDialog
    Dim filechosen As Long
    Dim wb As Workbook    
    Dim myrng As Range
''
    
        Set myfilearray = Application.FileDialog(msoFileDialogFilePicker)
            myfilearray.InitialView = msoFileDialogViewList
            myfilearray.AllowMultiSelect = False
            ' title that shows on top of window
            myfilearray.Title = "Choose Guideline File"
            myfilearray.Filters.Clear
            ' . extension of file
            myfilearray.Filters.Add "excel filse", "*.xl**"
        filechosen = myfilearray.Show

        If filechosen <> -1 Then
            MsgBox "You have chosen nothing!"
            BaseTotalRows = 1
            Exit Sub
        End If
        
        myfile = myfilearray.SelectedItems(1)
        myfile2 = Split(myfile, "\")(UBound(Split(myfile, "\")))
B_events_false
        If IsOpen(CStr(myfile2)) = 1 Then
            Workbooks(myfile2).Close
        End If
A_events_true

        Set GuidelineFile = Workbooks.Open(myfile)
'        Set GuidelineFile = GuidelineFile.Activate
        
End Sub
