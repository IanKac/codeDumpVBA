Attribute VB_Name = "mdlFillOut"
'This module is filling out data based on dropdown selection'

Private module
Option Explicit

Public Sub setupDropdown()
''Declarations''
    Dim i As Long
    Dim rowStart As Long, rowLast As Long
    Dim IDCol As Long, nameCol As Long
    
    Dim colColl As Collection, colNameColl As Collection
    Dim dropDownColl As Collection
    
    Dim dropDownRng As Range
    Dim db As Worksheet
''

'Setup columns names collection'
        Set colNameColl = New Collection
        colNameColl.Add "ID"
        colNameColl.Add "Name"
        colNameColl.Add "City"
        colNameColl.Add "Street"
        colNameColl.Add "Building"
        colNameColl.Add "Local"
        colNameColl.Add "Phone"
        colNameColl.Add "NIP"
                
'Get columns'
        Set db = wshDB
        Set colColl = New Collection
        
        For i = 1 To colNameColl.Count
            On Error Resume Next
            If IsError(db.UsedRange.Find(colNameColl(i))) = True Then
'If no column name found inform user about it'
'                Call errorhandler(405, colNameColl(i))
            End If
            colColl.Add db.UsedRange.Find(colNameColl(i)), colNameColl(i)
        Next i
        
'Loop through rows and get data'
        Set dropDownColl = New Collection
        IDCol = colColl("ID").Column
        nameCol = colColl("Name").Column
        rowStart = colColl("ID").Row + 1
        rowLast = db.Cells(Rows.Count, IDCol).End(xlUp).Row
        For i = rowStart To rowLast
            dropDownColl.Add db.Cells(i, nameCol).Value
        Next i
        
'Put data into range'
        wshDropdown.UsedRange.Clear
        For i = 1 To dropDownColl.Count
            wshDropdown.Cells(i, 1).Value = dropDownColl(i)
        Next i
        Set dropDownRng = wshDropdown.Range(wshDropdown.Cells(1, 1), wshDropdown.Cells(dropDownColl.Count, 1))
        
'Put data into dropdown'
        With wshFillOut.Cells(1, 2).Validation
            .Delete
            .InCellDropdown = True
            .Add xlValidateList, , , Formula1:="='" & wshDropdown.Name & "'!" & dropDownRng.Address
        End With
        
End Sub

Public Sub fillOutData()
''Declarations''
    Dim idSelected As String
    
    Dim i As Long
    Dim IDRow As Long, IDCol As Long
    Dim colStart As Long, colLast As Long
    Dim rowFound As Long
    
    Dim dropDownColl As Collection, colFillColl As Collection
    Dim colDBColl As Collection, colNameColl As Collection
    
    Dim db As Worksheet, fill As Worksheet
''

'Setup columns names collection'
        Set colNameColl = New Collection
        colNameColl.Add "ID"
        colNameColl.Add "Name"
        colNameColl.Add "City"
        colNameColl.Add "Street"
        colNameColl.Add "Building"
        colNameColl.Add "Local"
        colNameColl.Add "Phone"
        colNameColl.Add "NIP"
                
'Get columns DB'
        Set db = wshDB
        Set colDBColl = New Collection
        
        For i = 1 To colNameColl.Count
            On Error Resume Next
            If IsError(db.UsedRange.Find(colNameColl(i))) = True Then
'If no column name found inform user about it'
'                Call errorhandler(405, colNameColl(i))
            End If
            colDBColl.Add db.UsedRange.Find(colNameColl(i)), colNameColl(i)
        Next i
        
'Get columns DB'
        Set fill = wshFillOut
        Set colFillColl = New Collection
        
        For i = 1 To colNameColl.Count
            On Error Resume Next
            If IsError(db.UsedRange.Find(colNameColl(i))) = True Then
'If no column name found inform user about it'
'                Call errorhandler(405, colNameColl(i))
            End If
            colFillColl.Add fill.UsedRange.Find(colNameColl(i)), colNameColl(i)
        Next i
        
        
'Find row with correct ID'
        idSelected = wshFillOut.Cells(1, 2).Value
        On Error Resume Next
        If IsError(db.UsedRange.Find(idSelected, lookat:=xlWhole)) = True Then
'                Call errorhandler(405, idselected))
                MsgBox "Error!"
        End If
        rowFound = db.UsedRange.Find(idSelected, lookat:=xlWhole).Row
        
'Loop through columns and get data'
        Set dropDownColl = New Collection
        IDRow = colDBColl("ID").Row
        colStart = colDBColl("ID").Column
        colLast = db.Cells(IDRow, Columns.Count).End(xlToLeft).Column
        For i = colStart To colLast
            dropDownColl.Add db.Cells(rowFound, i).Value
        Next i

'Fill out data'
        Application.EnableEvents = False
        
        For i = 1 To dropDownColl.Count
            fill.Cells(colFillColl(i).Row, colFillColl(i).Column + 1).Value = dropDownColl(i)
        Next i
        
        Application.EnableEvents = True
End Sub
