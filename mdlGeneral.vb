' enevt handler
Public Sub eventHandler(turnOn As Boolean)
'Turn on or off applications events, screen updating and such'

    With Application
        If turnOn = True Then
            .ScreenUpdating = True
            .EnableEvents = True
            .DisplayAlerts = False
            .StatusBar = "Ready"
            .Cursor = xlDefault
        ElseIf turnOn = False Then
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayAlerts = True
            .StatusBar = "Busy"
            .Cursor = xlWait
        End If
    End With
    
End Sub

' Formating for buttons
Public Sub wshFormat(wsh As Worksheet)
'Format Borders, format buttons'
'Time formating. Banner color rgb (0,155,205), background RGB(224, 235, 235)''
    If wsh.Name = wshTime.Name Then
        With wsh
            .UsedRange.Cells.Borders.LineStyle = xlContinuous
            .UsedRange.Cells.Borders.Color = RGB(0, 0, 0)
            .Range("A:AAA").Interior.Color = RGB(224, 235, 235)
            .Rows("1:6").Interior.Color = RGB(0, 155, 205)
            .Rows("1:6").Borders.Color = RGB(0, 155, 205)
            With .Shapes("TimeToMenu")
                .Top = wshLandingPage.Cells(1, 1).Top
                .Left = wshLandingPage.Cells(1, 1).Left + 5
                .Width = wshLandingPage.Cells(1, 1).Width * 2
                .Height = wshLandingPage.Cells(1, 1).Height * 3
                .Fill.ForeColor.RGB = RGB(100, 0, 255)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("TimeToCost")
                .Top = wshLandingPage.Cells(1, 1).Top
                .Left = wshLandingPage.Cells(1, 1).Left + wshLandingPage.Cells(1, 1).Width * 2 + 10
                .Width = wshLandingPage.Cells(1, 1).Width * 2
                .Height = wshLandingPage.Cells(1, 1).Height * 3
                .Fill.ForeColor.RGB = RGB(0, 155, 0)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("AddRecordTimeButton")
                .Top = wshLandingPage.Cells(1, 1).Top + wshLandingPage.Cells(1, 1).Height * 3 + 5
                .Left = wshLandingPage.Cells(1, 1).Left + 5 + wshLandingPage.Cells(1, 1).Width
                .Width = wshLandingPage.Cells(1, 1).Width
                .Height = wshLandingPage.Cells(1, 1).Height * 1.5
                .Fill.ForeColor.RGB = RGB(0, 51, 153)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("DeleteRecordTimeButton")
                .Top = wshLandingPage.Cells(1, 1).Top + wshLandingPage.Cells(1, 1).Height * 3 + 5
                .Left = wshLandingPage.Cells(1, 1).Left + wshLandingPage.Cells(1, 1).Width * 2 + 10 + wshLandingPage.Cells(1, 1).Width
                .Width = wshLandingPage.Cells(1, 1).Width
                .Height = wshLandingPage.Cells(1, 1).Height * 1.5
                .Fill.ForeColor.RGB = RGB(0, 51, 153)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("CountValueByMonthButton")
                .Top = wshLandingPage.Cells(1, 1).Top + wshLandingPage.Cells(1, 1).Height * 3
                .Left = wshLandingPage.Cells(1, 1).Left + wshLandingPage.Cells(1, 1).Width * 15 + 10 + wshLandingPage.Cells(1, 1).Width
                .Width = wshLandingPage.Cells(1, 1).Width * 3
                .Height = wshLandingPage.Cells(1, 1).Height * 2
                .Fill.ForeColor.RGB = RGB(0, 51, 153)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("CountValueByClientButton")
                .Top = wshLandingPage.Cells(1, 1).Top + wshLandingPage.Cells(1, 1).Height * 3
                .Left = wshLandingPage.Cells(1, 1).Left + wshLandingPage.Cells(1, 1).Width * 15 _
                        + 10 + wshLandingPage.Cells(1, 1).Width + wshTime.Shapes("CountValueByMonthButton").Width * 1.1
                .Width = wshLandingPage.Cells(1, 1).Width * 3
                .Height = wshLandingPage.Cells(1, 1).Height * 2
                .Fill.ForeColor.RGB = RGB(0, 51, 153)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
        End With
            
'Cost formating. Banner color rgb (0,155,0), background RGB(224, 235, 235)'
    ElseIf wsh.Name = wshCost.Name Then
        With wsh
            .UsedRange.Cells.Borders.LineStyle = xlContinuous
            .UsedRange.Cells.Borders.Color = RGB(0, 0, 0)
            .Range("A:AAA").Interior.Color = RGB(224, 235, 235)
            .Rows("1:6").Interior.Color = RGB(0, 155, 0)
            .Rows("1:6").Borders.Color = RGB(0, 155, 0)
            With .Shapes("CostToMenu")
                .Top = wshLandingPage.Cells(1, 1).Top
                .Left = wshLandingPage.Cells(1, 1).Left + 5
                .Width = wshLandingPage.Cells(1, 1).Width * 2
                .Height = wshLandingPage.Cells(1, 1).Height * 3
                .Fill.ForeColor.RGB = RGB(100, 0, 255)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("CostToTime")
                .Top = wshLandingPage.Cells(1, 1).Top
                .Left = wshLandingPage.Cells(1, 1).Left + wshLandingPage.Cells(1, 1).Width * 2 + 10
                .Width = wshLandingPage.Cells(1, 1).Width * 2
                .Height = wshLandingPage.Cells(1, 1).Height * 3
                .Fill.ForeColor.RGB = RGB(0, 155, 205)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("AddRecordButton")
                .Top = wshLandingPage.Cells(1, 1).Top + wshLandingPage.Cells(1, 1).Height * 3 + 5
                .Left = wshLandingPage.Cells(1, 1).Left + 5 + wshLandingPage.Cells(1, 1).Width
                .Width = wshLandingPage.Cells(1, 1).Width
                .Height = wshLandingPage.Cells(1, 1).Height * 1.5
                .Fill.ForeColor.RGB = RGB(0, 51, 153)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("DeleteRecordButton")
                .Top = wshLandingPage.Cells(1, 1).Top + wshLandingPage.Cells(1, 1).Height * 3 + 5
                .Left = wshLandingPage.Cells(1, 1).Left + wshLandingPage.Cells(1, 1).Width * 2 + 10 + wshLandingPage.Cells(1, 1).Width
                .Width = wshLandingPage.Cells(1, 1).Width
                .Height = wshLandingPage.Cells(1, 1).Height * 1.5
                .Fill.ForeColor.RGB = RGB(0, 51, 153)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("CalculateCostButton")
                .Top = wshLandingPage.Cells(1, 1).Top + wshLandingPage.Cells(1, 1).Height * 3 + 5
                .Left = wshLandingPage.Cells(1, 1).Left + wshLandingPage.Cells(1, 1).Width * 4 + 10 + wshLandingPage.Cells(1, 1).Width
                .Width = wshLandingPage.Cells(1, 1).Width * 2
                .Height = wshLandingPage.Cells(1, 1).Height * 1.5
                .Fill.ForeColor.RGB = RGB(0, 51, 153)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
        End With
            
'Menu formating. Banner color rgb (100,0,255), background RGB(224, 235, 235)'
    ElseIf wsh.Name = wshLandingPage.Name Then
        With wsh
            .Columns(1).ColumnWidth = 8.43
            .Rows(1).RowHeight = 20
            .Cells.Borders.LineStyle = xlNone
            .Cells.Borders.Color = RGB(255, 255, 255)
            .Range("A:AAA").Interior.Color = RGB(224, 235, 235)
            .Range("A:AAA").Borders.Color = RGB(224, 235, 235)
            .Rows("1:6").Interior.Color = RGB(100, 0, 255)
            .Rows("1:6").Borders.Color = RGB(100, 0, 255)
            With .Shapes("MenuToTime")
                .Top = wshLandingPage.Cells(1, 1).Top
                .Left = wshLandingPage.Cells(1, 1).Left + 5
                .Width = wshLandingPage.Cells(1, 1).Width * 2
                .Height = wshLandingPage.Cells(1, 1).Height * 3
                .Fill.ForeColor.RGB = RGB(0, 155, 205)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("MenuToCost")
                .Top = wshLandingPage.Cells(1, 1).Top
                .Left = wshLandingPage.Cells(1, 1).Left + wshLandingPage.Cells(1, 1).Width * 2 + 10
                .Width = wshLandingPage.Cells(1, 1).Width * 2
                .Height = wshLandingPage.Cells(1, 1).Height * 3
                .Fill.ForeColor.RGB = RGB(0, 155, 0)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("DisablePass")
                .Top = wshLandingPage.Cells(1, 1).Top + wshLandingPage.Shapes("MenuToTime").Height * 1.5
                .Left = wshLandingPage.Cells(1, 1).Left + 5
                .Width = wshLandingPage.Cells(1, 1).Width * 2
                .Height = wshLandingPage.Cells(1, 1).Height * 3
                .Fill.ForeColor.RGB = RGB(255, 50, 0)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
            With .Shapes("EnablePass")
                .Top = wshLandingPage.Cells(1, 1).Top + wshLandingPage.Shapes("MenuToTime").Height * 1.5
                .Left = wshLandingPage.Cells(1, 1).Left + wshLandingPage.Cells(1, 1).Width * 2 + 10
                .Width = wshLandingPage.Cells(1, 1).Width * 2
                .Height = wshLandingPage.Cells(1, 1).Height * 3
                .Fill.ForeColor.RGB = RGB(255, 50, 0)
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
            End With
        End With
    End If
    
End Sub

' conversion to Boolean
Function ValToboolean(val As String) As Boolean

    val = Trim(UCase(val))
    
    Select Case val
        Case "TAK", "YES", "Y", "T", "TRUE"
            ValToboolean = True
        Case "NIE", "NO", "N", "FALSE"
            ValToboolean = False
        Case Else
            errorHandler (69)
    End Select

End Function

' Password managment
Public Sub DisablePasswords()
''Declarations''
    Dim pass As String
    
''
    
    If UCase(Environ$("username")) Like UCase("Jan*") _
            Or UCase(Environ$("username")) Like UCase("DELL") _
            Or UCase(Environ$("username")) Like UCase("JWedrych*") Then
        pass = UCase("zamekhauru")
    Else
        pass = UCase(Trim(InputBox("Provide password!")))
    End If
    
'Turn on/off button'
    wshLandingPage.Shapes("Disablepass").Visible = msoFalse
    wshLandingPage.Shapes("Enablepass").Visible = msoTrue
    
'Disable passwords'
    If pass = UCase(Trim("zamekhauru")) Then
        'Remove passwords from tabs'
        wshLandingPage.Unprotect pass
        wshTime.Unprotect pass
        wshCost.Unprotect pass
    Else
        MsgBox ("Wrong password!")
    End If

End Sub

Public Sub EnablePasswords()
    
'    If vbYes = MsgBox("Jeste? pewna/y, ?e chcesz zahas3owaa plik teraz?" & vbLf & "Plik has3uje sie ka?dorazowo po odpaleniu.", vbYesNo) Then
        wshLandingPage.Protect (UCase("zamekhauru"))
        wshTime.Protect (UCase("zamekhauru"))
        wshCost.Protect (UCase("zamekhauru"))
'Turn on/off button'
        wshLandingPage.Shapes("Disablepass").Visible = msoTrue
        wshLandingPage.Shapes("Enablepass").Visible = msoFalse

'    End If
    
End Sub

