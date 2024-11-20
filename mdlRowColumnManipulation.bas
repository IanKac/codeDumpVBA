Attribute VB_Name = "mdlRowColumnManipulation"
Private module
Option Explicit

Private Sub addRow(above As Boolean, i As Long)
    
    If above = True Then
        Rows(i).Insert xlShiftDown
        Rows(i).Interior.Color = xlNone
    Else
        Rows(i + 1).Insert xlShiftUp
        Rows(i + 1).Interior.Color = xlNone
    End If
    
End Sub

Private Sub deleteRow(i As Long)
    
    Rows(i).Delete
    
End Sub
Private Sub addColumn(addToTheLeft As Boolean, i As Long)

    If addToTheLeft = True Then
        Columns(i).EntireColumn.Insert
        Columns(i).Interior.Color = vbRed
    Else
        Columns(i + 1).EntireColumn.Insert
        Columns(i + 1).Interior.Color = vbRed
    End If

End Sub


Private Sub deleteColumn(i As Long)

    Columns(i).EntireColumn.Delete

End Sub
Sub addLeft()
    
    Call addColumn(True, Selection.Column)
    
End Sub

Sub addRight()
    
    Call addColumn(False, Selection.Column)
    
End Sub
Sub addAbove()

    Call addRow(True, Selection.Row)

End Sub

Sub addBelow()

    Call addRow(False, Selection.Row)
    
End Sub

Sub deleteRowBtn()
    
    Call deleteRow(Selection.Row)
    
End Sub

Sub deleteColumnBtn()

    Call deleteColumn(Selection.Column)
    
End Sub
