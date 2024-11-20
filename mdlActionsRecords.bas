Attribute VB_Name = "mdlActionsRecords"
Public module

Sub addRecord()
'This sub will add new record to database'
''Declartions''
    Dim i As Long, j As Long
    Dim maxID As Long
    Dim rowID As Long, colID As Long, rowLastID As Long
    Dim rngIDRecord As Range, colIDRecord As Long, rowLastIdTarget As Long
    Dim shtName As String
    
    Dim rngID As Range
    Dim TargetSht As Worksheet, dataFile As Worksheet
    
    Dim colColl As Collection
    Dim colCollRecord As Collection, dataColColl As Collection
    Dim newRecordColl As Collection, newRecord As Collection
''

'Disable events'
    Call eventHandler(False)
    
'Log'
    Call logger("addRecord", "Start")
    
'Create collection with columns names'
    Set colColl = New Collection
    colColl.Add "ID"
    colColl.Add "ID Pracownika"
    colColl.Add "Imie"
    colColl.Add "Nazwisko"
    colColl.Add "Data"
    colColl.Add "Start"
    colColl.Add "Koniec"
    
'Get columns from this file'
    Set TargetSht = ThisWorkbook.Worksheets("Record")
    Set colCollRecord = findColumns(colColl, "Record", ThisWorkbook)
    
'Get biggest ID from this file. How many records there are in record file'
    Set rngID = TargetSht.UsedRange.Find(colColl(1))
    rowLastIdTarget = TargetSht.Cells(Rows.Count, colCollRecord("ID")).End(xlUp).Row

'If row of ID is same as last record add 1'
    If rngID.Row = rowLastIdTarget Then
        rowLastIdTarget = rowLastIdTarget + 1
    End If
    
    maxID = TargetSht.Cells(rowLastIdTarget, colCollRecord("ID")).Value
    
'Call for file to open'
    Set openedFile = Nothing
    Call loadFile
    shtName = "TEST"
    Set dataFile = openedFile.Worksheets(shtName)
    
'Create collection with columns names'
    Set colColl = New Collection
    colColl.Add "ID"
    colColl.Add "ID Pracownika"
    colColl.Add "Imie"
    colColl.Add "Nazwisko"
    colColl.Add "Data"
    colColl.Add "Start"
    colColl.Add "Koniec"
    
'Get columns from file'
    Set dataColColl = findColumns(colColl, "TEST", openedFile)
    
'Find how many rows there are in new file'
    rowID = dataFile.UsedRange.Find(colColl(2)).Row
    colID = dataFile.UsedRange.Find(colColl(2)).Column
    rowLastID = dataFile.Cells(Rows.Count, colID).End(xlUp).Row
    
'Loop through data in file and add it to newRecord Collection'
    Set newRecordColl = New Collection
    For i = rowID + 1 To rowLastID
        Set newRecord = New Collection
        
'Create individual record'
        For j = 1 To dataColColl.Count
            newRecord.Add Item:=dataFile.Cells(i, dataColColl(j)).Value, Key:=colColl(j)
        Next j
        
'Add individual record to set of records'
        newRecordColl.Add Item:=newRecord, Key:="dataFileRow" & i
    Next i
    
'Loop throught data in collection and add it to file'
    For i = 1 To newRecordColl.Count

        'No ID from datafile. Add new automatically'
        If maxID > 0 Then
            rowLastIdTarget = rowLastIdTarget + 1
        End If
        TargetSht.Cells(rowLastIdTarget, colCollRecord(1)).Value = maxID + 1
        
        'Update max ID'
        maxID = TargetSht.Cells(rowLastIdTarget, colCollRecord(1)).Value
        
        'Get record from new records' collection'
        Set newRecord = newRecordColl(i)
        
        For j = 2 To dataColColl.Count
            TargetSht.Cells(rowLastIdTarget, colCollRecord(j)).Value = newRecord(j)
        Next j
    Next i
    
'Close opened file'
    openedFile.Close (vbNo)
    
    
'Log'
    Call logger("addRecord", "Finish")
    
'Enable events'
    Call eventHandler(True)

End Sub

Sub addDistinctRecord()
'This sub will add new record to database with omission of already present ones'
''Declartions''
    Dim i As Long, j As Long
    Dim maxID As Long
    Dim rowID As Long, colID As Long, rowLastID As Long
    Dim rngIDRecord As Range, colIDRecord As Long, rowLastIdTarget As Long
    Dim recordMatchRow As Long
    
    Dim shtName As String
    Dim recordTarget As String, recordInput As String
    
    Dim rngID As Range
    Dim TargetSht As Worksheet, dataFile As Worksheet
    
    Dim colColl As Collection
    Dim colCollRecord As Collection, dataColColl As Collection
    Dim newRecordColl As Collection, newRecord As Collection
    
    Dim recordMatch As Boolean
''

'Disable events'
    Call eventHandler(False)
    
'Log'
    Call logger("addRecord", "Start")
    
'Create collection with columns names'
    Set colColl = New Collection
    colColl.Add "ID"
    colColl.Add "ID Pracownika"
    colColl.Add "Imie"
    colColl.Add "Nazwisko"
    colColl.Add "Data"
    colColl.Add "Start"
    colColl.Add "Koniec"
    
'Get columns from this file'
    Set TargetSht = ThisWorkbook.Worksheets("Record")
    Set colCollRecord = findColumns(colColl, "Record", ThisWorkbook)
    
'Get biggest ID from this file. How many records there are in record file'
    Set rngID = TargetSht.UsedRange.Find(colColl(1))
    rowLastIdTarget = TargetSht.Cells(Rows.Count, colCollRecord("ID")).End(xlUp).Row

'If row of ID is same as last record add 1'
    If rngID.Row = rowLastIdTarget Then
        rowLastIdTarget = rowLastIdTarget + 1
    End If
    
    maxID = TargetSht.Cells(rowLastIdTarget, colCollRecord("ID")).Value
    
'Call for file to open'
    Set openedFile = Nothing
    Call loadFile
    shtName = "TEST"
    Set dataFile = openedFile.Worksheets(shtName)
    
'Create collection with columns names'
    Set colColl = New Collection
    colColl.Add "ID"
    colColl.Add "ID Pracownika"
    colColl.Add "Imie"
    colColl.Add "Nazwisko"
    colColl.Add "Data"
    colColl.Add "Start"
    colColl.Add "Koniec"
    
'Get columns from file'
    Set dataColColl = findColumns(colColl, "TEST", openedFile)
    
'Find how many rows there are in new file'
    rowID = dataFile.UsedRange.Find(colColl(2)).Row
    colID = dataFile.UsedRange.Find(colColl(2)).Column
    rowLastID = dataFile.Cells(Rows.Count, colID).End(xlUp).Row
    
'Loop through data in file and add it to newRecord Collection'
    Set newRecordColl = New Collection
    For i = rowID + 1 To rowLastID
        Set newRecord = New Collection
        
'Create individual record'
        For j = 1 To dataColColl.Count
            newRecord.Add Item:=dataFile.Cells(i, dataColColl(j)).Value, Key:=colColl(j)
        Next j
        
'Add individual record to set of records'
        newRecordColl.Add Item:=newRecord, Key:="dataFileRow" & i
    Next i
    
'Loop throught data in collection and add it to file'
    For i = 1 To newRecordColl.Count

'record to check'
        'Get record from new records' collection'
        Set newRecord = newRecordColl(i)
        recordInput = newRecord("ID Pracownika") _
        & " " _
        & newRecord("Data")
        
'Looking for match
        For j = rngID.Row To rowLastIdTarget
            recordTarget = _
                TargetSht.Cells(j, colCollRecord("ID Pracownika")).Value _
                & " " _
                & TargetSht.Cells(j, colCollRecord("Data")).Value
            If recordInput = recordTarget Then
                recordMatch = True
                recordMatchRow = j
                Debug.Print ("Match found for: '" & recordInput _
                        & "' = '" & recordTarget & _
                        "'. For row: " & recordMatchRow)
                Exit For
            End If
        Next j


        If recordMatch = True Then
            For j = 2 To dataColColl.Count
                TargetSht.Cells(recordMatchRow, colCollRecord(j)).Value = newRecord(j)
            Next j
        Else
            'No ID from datafile. Add new automatically'
            If maxID > 0 Then
                rowLastIdTarget = rowLastIdTarget + 1
            End If
            TargetSht.Cells(rowLastIdTarget, colCollRecord(1)).Value = maxID + 1
            
            'Update max ID'
            maxID = TargetSht.Cells(rowLastIdTarget, colCollRecord(1)).Value
            
            For j = 2 To dataColColl.Count
                TargetSht.Cells(rowLastIdTarget, colCollRecord(j)).Value = newRecord(j)
            Next j
        End If
    Next i
    
'Close opened file'
    openedFile.Close (vbNo)
    
    
'Log'
    Call logger("addRecord", "Finish")
    
'Enable events'
    Call eventHandler(True)

End Sub
Sub calculateSalary()
'This sub calculates salary for current dataset'
''Declarations''
''

'Disable events'
'Log'
'Loop through data in file and add it to collection'
'Log'
'Enable events'
End Sub
