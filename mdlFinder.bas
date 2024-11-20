Attribute VB_Name = "mdlFinder"
Private module
Option Explicit

Function findColumns(nameColl As Collection, shtName As String, dataFile As Workbook) As Collection
''Declarations''
    Dim i As Long
    
    Dim usedSht As Worksheet
    
    Dim colFoundColl As Collection
''
'Select file to find data in'
    Set usedSht = dataFile.Worksheets(shtName)

'Create new column collection'
    Set colFoundColl = New Collection
    
'Find columns in opened file'
    For i = 1 To nameColl.Count
        On Error Resume Next
        If IsError(usedSht.UsedRange.Find(nameColl(i), lookat:=xlWhole).Column = True) Then
           Call errorHandler(404)
        End If
        
       colFoundColl.Add Item:=usedSht.UsedRange.Find(nameColl(i), lookat:=xlWhole).Column, Key:=nameColl(i)
    Next i
    
'Return collection of found columns'
    Set findColumns = colFoundColl
    
End Function
