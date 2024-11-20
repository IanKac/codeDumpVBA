Attribute VB_Name = "mdlErrorHandling"
Option Explicit

Public Sub errorHandler(errorNmbr As Long)

'Log'
    Call logger("errorHandler-" & errorNmbr, "Start" & errorNmbr)
    
    If errorNmbr = 1 Then
'Joke'
    ElseIf errorNmbr = 69 Then
        Call MsgBox("To weŸ siê zdecyduj!" & vbLf & "Ca³uski.", _
                vbOKOnly, "Error 69")
'300s Wrong tab'
    ElseIf errorNmbr = 300 Then
        Call MsgBox("Wrong tab selected!" & vbLf & "Go to correct tab first and then repeat operation!", _
                vbCritical, "Error 300")
'400s navigation'
    ElseIf errorNmbr = 400 Then
    ElseIf errorNmbr = 404 Then
        Call MsgBox("No such place in this file!" & vbLf & "Create it or contact developer.", _
                vbCritical, "Error 404")
    ElseIf errorNmbr = 405 Then
        Call MsgBox("No file have been chosen." & vbLf & "Please retry operation.", _
                vbCritical, "Error 405")
    ElseIf errorNmbr = 406 Then
        Call MsgBox("No such user " & Environ("Username") & "." & _
                vbLf & "Create it or contact developer.", vbCritical, "Error 406")
    End If

'Log'
    Call logger("errorHandler-" & errorNmbr, "Finish" & errorNmbr)
    
'EnableEvents'
    Call eventHandler(True)
    End
    
End Sub
