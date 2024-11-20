Attribute VB_Name = "mdlNavigation"
Private module
Option Explicit
Public Sub jumpToWsh(buttonPressed As String)

    If buttonPressed = "MenuToRecord" Then
        wshRecord.Activate
'    ElseIf buttonPressed = "MenuToCost" Then
'        wshCost.Activate
'    ElseIf buttonPressed = "TimeToMenu" Then
'        wshLandingPage.Activate
'    ElseIf buttonPressed = "TimeToCost" Then
'        wshCost.Activate
'    ElseIf buttonPressed = "CostToTime" Then
'        wshRecord.Activate
'    ElseIf buttonPressed = "CostToMenu" Then
'        wshLandingPage.Activate
'    ElseIf buttonPressed = "CostToDB" Then
'        wshDBCost.Activate
'    ElseIf buttonPressed = "TimeToDB" Then
'        wshDBTime.Activate
'    ElseIf buttonPressed = "DBToCost" Then
'        wshCost.Activate
    ElseIf buttonPressed = "DBToTime" Then
        wshRecord.Activate
    Else
'Error'
        Call errorHandler(404)
    End If
    
End Sub

Private Sub MenuToRecord_Click()

    Call jumpToWsh("MenuToRecord")
    
End Sub

Private Sub MenuToCost_Click()
    
    jumpToWsh ("MenuToCost")
    
End Sub
Private Sub TimeToMenu_Click()
    
    jumpToWsh ("TimeToMenu")
    
End Sub
Private Sub TimeToCost_Click()
    
    jumpToWsh ("TimeToCost")
    
End Sub
Private Sub CostToTime_Click()
    
    jumpToWsh ("CostToTime")
    
End Sub
Private Sub CostToMenu_Click()
    
    jumpToWsh ("CostToMenu")
    
End Sub

Private Sub CostToDB_Click()
    
    Call jumpToWsh("CostToDB")
    
End Sub
Private Sub TimeToDB_Click()
    
    jumpToWsh ("TimeToDB")
    
End Sub
Private Sub DBToCost_Click()
    
    jumpToWsh ("DBToCost")
    
End Sub
Private Sub DBToTime_Click()
    
    jumpToWsh ("DBToTime")
    
End Sub

Private Sub hideTabs(wshCurrent As Worksheet)
''Declarations''
    Dim i As Long
''

    Call eventHandler(False)
'Find destination and make it visiable'
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Worksheets(i).name = wshCurrent.name Then
            Worksheets(i).Visible = xlSheetVisible
            Worksheets(i).Select
            Exit For
        End If
    Next i
    
'Hide rest of tabs'
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Worksheets(i).name <> wshCurrent.name Then
            Worksheets(i).Visible = xlSheetHidden
        End If
    Next i

    Call eventHandler(True)
    
End Sub
