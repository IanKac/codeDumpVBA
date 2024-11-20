Attribute VB_Name = "mdlUsers"
Private module
Option Explicit
Public oUser As String

Public Sub setCurrentUser()
'This sub sets up current user for easier verification later'
''Declarations''
    Dim i As Long
    Dim rowID As Long, colID As Long
    Dim rowUser As Long
    Dim lastCol As Long
    
    Dim winName As String
    
    Dim rngID As Range
    
    
''

'Disable events'
        Call eventHandler(False)
        
'LOG'
        Call logger("setCurrentUser", "Start")
        
'Find Windows name of current user'
        winName = Environ("UserName")
    
'Move colum names as row names'
        Set rngID = wshUsers.UsedRange.Find("ID", lookat:=xlWhole)
        rowID = rngID.Row
        colID = rngID.Column
        lastCol = wshUsers.Cells(rowID, Columns.Count).End(xlToLeft).Column
        
        For i = 1 To lastCol
            wshCurUser.Cells(i, 1).Value = wshUsers.Cells(rowID, i).Value
        Next i
    
'Move data from wshUsers'
        
        On Error Resume Next
        If IsError(wshUsers.UsedRange.Find(winName, lookat:=xlWhole).Row) = True Then
            Call errorHandler(406)
        Else
            rowUser = wshUsers.UsedRange.Find(winName, lookat:=xlWhole).Row
        End If
        
        For i = 1 To lastCol
            wshCurUser.Cells(i, 2).Value = wshUsers.Cells(rowUser, i).Value
        Next i
        
        
'LOG'
        Call logger("setCurrentUser", "Finish")
        
'Enable events'
        Call eventHandler(True)

End Sub


Public Function verifyUser(rightToVerify As String) As Boolean
''Declarations''
    Dim oUser As clsUser
''
    Set oUser = New clsUser
    
    rightToVerify = UCase(Trim(rightToVerify))
    
    If rightToVerify = "ADMIN" Then
        If oUser.IsAdmin = True Then
            verifyUser = True
        End If
    ElseIf rightToVerify = "USER" Then
        If oUser.IsRegularUser = True Then
            verifyUser = True
        End If
    End If
    
End Function
