Attribute VB_Name = "mdlNumberToWords"
Option Explicit

Private Sub convertBulk(rowFrom As Long, rowTo As Long)
''Declarations''
    Dim i As Long
    
''
    
    For i = rowFrom To rowTo
        Cells(i, 2).Value = convertNumberToWords(Cells(i, 1).Value) & " " & convertCurrencyToString(Cells(i, 1).Value, "PLN")
        Cells(i, 2).Value = Replace(Cells(i, 2).Value, "   ", " ")
        Cells(i, 2).Value = Replace(Cells(i, 2).Value, "  ", " ")
    Next i

End Sub

Private Function convertNumberToWords(convertNbr As Long) As String
''Declarations''
    Dim millions As Long, thousands As Long, hundreds As Long
    Dim tens As Long, singles As Long
    Dim teens As Long
    Dim convertNbrTemp As Long
    
    Dim singlesColl As Collection, teensColl As Collection
    Dim tensColl As Collection, hundredsColl As Collection
    Dim thousandsColl As Collection, millionsColl As Collection
    Dim curColl As Collection
    
    Dim decimalEuro As Boolean
    
    Dim convertedString As String
    Dim millionsString As String, thousandsString As String, hundredsString As String
    Dim tensString As String, singlesString As String
    Dim teensString As String
''

'Set up collections'
'Singles'
        Set singlesColl = New Collection
        singlesColl.Add "jeden", "jeden"
        singlesColl.Add "dwa", "dwa"
        singlesColl.Add "trzy", "trzy"
        singlesColl.Add "cztery", "cztery"
        singlesColl.Add "pi��", "pi��"
        singlesColl.Add "sze��", "sze��"
        singlesColl.Add "siedem", "siedem"
        singlesColl.Add "osiem", "osiem"
        singlesColl.Add "dziewi��", "dziewi��"
        singlesColl.Add "zero", "zero"
        
'nastki'
        Set teensColl = New Collection
        teensColl.Add "jedena�cie", "jedena�cie"
        teensColl.Add "dwana�cie", "dwana�cie"
        teensColl.Add "trzyna�cie", "trzyna�cie"
        teensColl.Add "czterna�cie", "czterna�cie"
        teensColl.Add "pi�tna�cie", "pi�tna�cie"
        teensColl.Add "szesna�cie", "szesna�cie"
        teensColl.Add "siedemna�cie", "siedemna�cie"
        teensColl.Add "osiemna�cie", "osiemna�cie"
        teensColl.Add "dziewietna�cie", "dziewietna�cie"
        
'Tens'
        Set tensColl = New Collection
        tensColl.Add "dziesi��", "dziesie�"
        tensColl.Add "dwadzie�cia", "dwadzie�cia"
        tensColl.Add "trzydzie�ci", "trzydzie�ci"
        tensColl.Add "czterdzie�ci", "czterdzie�ci"
        tensColl.Add "pi��dziesi�t", "pi��dziesi�t"
        tensColl.Add "sze��dziesi�t", "sze��dziesi�t"
        tensColl.Add "siedemdziesi�t", "siedemdziesi�t"
        tensColl.Add "osiemdziesi�t", "osiemdziesi�t"
        tensColl.Add "dziewi��dziesi�t", "dziewi��dziesi�t"
        
'Hundreds'
        Set hundredsColl = New Collection
        hundredsColl.Add "sto", "sto"
        hundredsColl.Add "dwie�cie", "dwie�cie"
        hundredsColl.Add "trzysta", "trzysta"
        hundredsColl.Add "czterysta", "czterysta"
        hundredsColl.Add "pi��set", "pi��set"
        hundredsColl.Add "sze��set", "sze��set"
        hundredsColl.Add "siedemset", "siedemset"
        hundredsColl.Add "osiemset", "osiemset"
        hundredsColl.Add "dziewi��set", "dziewi��set"
        
'Thousands'
        Set thousandsColl = New Collection
        thousandsColl.Add "tysi�c", "tysi�c"
        thousandsColl.Add "tysi�ce", "tysi�ce"
        thousandsColl.Add "tysi�cy", "tysi�cy"
        
        
'Millions'
        Set millionsColl = New Collection
        millionsColl.Add "milion", "milion"
        millionsColl.Add "miliony", "miliony"
        millionsColl.Add "milion�w", "milion�w"

'Getting right decimal separator'
        If Application.International(xlDecimalSeparator) = "," Then
            decimalEuro = True
        Else
            decimalEuro = False
        End If

'Set up convertNbrTemp'
        convertNbrTemp = convertNbr
        
'Getting millions from value'
        millions = convertNbrTemp \ 1000000
        convertNbrTemp = convertNbrTemp - millions * 1000000
        If millions > 0 Then
            millionsString = Trim(convertNumberToWords(millions))
        End If
        
'Getting thousands from value'
        thousands = convertNbrTemp \ 1000
        convertNbrTemp = convertNbrTemp - thousands * 1000
        If thousands > 0 Then
            thousandsString = Trim(convertNumberToWords(thousands))
        End If
        
'Getting hundreds from value'
        hundreds = convertNbrTemp \ 100
        convertNbrTemp = convertNbrTemp - hundreds * 100
                
'If belowe 20 then special case'
        If convertNbrTemp < 20 And convertNbrTemp > 10 Then
            teens = convertNbrTemp - 10
        
'Getting tens from value'
        Else
            tens = convertNbrTemp \ 10
            convertNbrTemp = convertNbrTemp - tens * 10
            
'Getting singles from value'
            singles = convertNbrTemp
        End If
        
'Convert longs into strings based on collections'
        
'Millions'
        If millions = 0 Then
            millionsString = ""
        ElseIf millions = 1 Then
            millionsString = millionsString & " " & millionsColl(1)
        ElseIf millions >= 5 Then
            millionsString = millionsString & " " & millionsColl(3)
        Else
            millionsString = millionsString & " " & millionsColl(2)
        End If
    
'Thousands'
        If thousands = 0 Then
            thousandsString = ""
        ElseIf thousands = 1 Then
            thousandsString = thousandsString & " " & thousandsColl(1)
        ElseIf thousands >= 5 Then
            thousandsString = thousandsString & " " & thousandsColl(3)
        Else
            thousandsString = thousandsString & " " & thousandsColl(2)
        End If
        
'Hundreds'
        If hundreds = 0 Then
            hundredsString = ""
        Else
            hundredsString = hundredsColl(hundreds)
        End If
    
'Tens'
        If tens = 0 Then
            tensString = ""
        ElseIf tens > 0 Then
            tensString = tensColl(tens)
        End If
        
'Teens'
        If teens = 0 Then
            teensString = ""
        ElseIf teens > 0 Then
            teensString = teensColl(teens)
        End If
    
'Singles'
        If singles = 0 Then
            singlesString = ""
        ElseIf singles > 0 Then
            singlesString = singlesColl(singles)
        End If
                
        
'Create string'
        If Len(teensString) = 0 Then
            convertedString = Trim(millionsString) & " " & Trim(thousandsString) & _
                " " & Trim(hundredsString) & " " & Trim(tensString) & " " _
                & Trim(singlesString)
        Else
            convertedString = Trim(millionsString) & " " & Trim(thousandsString) & _
                " " & Trim(hundredsString) & " " & Trim(teensString)
        End If
    
'Remove unneeded space'
        convertedString = Trim(convertedString)
    
'If no string then input zero'
        If Len(convertedString) = 0 Then
            convertedString = singlesColl("zero")
        End If
        
'Return value'
        convertNumberToWords = convertedString
    
End Function
Private Function convertCurrencyToString(currencyValue As Long, currencyString As String)
''Declarations''
    Dim curValue As String
    
    Dim curColl As Collection

''

'Setting up currency collection'
        Set curColl = New Collection
        If currencyString = "PLN" Then
            curColl.Add "Z�oty"
            curColl.Add "Z�ote"
            curColl.Add "Z�otych"
        ElseIf currencyString = "USD" Then
            curColl.Add "Dollar"
            curColl.Add "Dollary"
            curColl.Add "Dollar�w"
        ElseIf currencyString = "EUR" Then
            curColl.Add "EURO"
        Else
'            Call errorhandler(407)
        End If
        
'split currency value'
        curValue = Right(currencyValue, 1)
        
'Determin correct currency version'
        If currencyString = "PLN" Then
            If currencyValue = 1 Then
                convertCurrencyToString = curColl(1)
            ElseIf (curValue = 2 Or curValue = 3 _
                    Or curValue = 4) And _
                    (currencyValue < 5 Or currencyValue > 20) Then
                convertCurrencyToString = curColl(2)
            Else
                convertCurrencyToString = curColl(3)
            End If
        ElseIf currencyString = "USD" Then
            If currencyValue = 1 Then
                convertCurrencyToString = curColl(1)
            ElseIf (curValue = 2 Or curValue = 3 _
                    Or curValue = 4) And _
                    (currencyValue < 5 Or currencyValue > 20) Then
                convertCurrencyToString = curColl(2)
            Else
                convertCurrencyToString = curColl(3)
            End If
        ElseIf currencyString = "EUR" Then
            convertCurrencyToString = curColl(1)
        End If

End Function
Sub convertValBtn()
''Declarations''
    Dim i As Long, j As Long
''

'Provide data'
'        i = CLng(InputBox("First row to change"))
'        j = CLng(InputBox("Last row to change"))
        i = 1
        j = 100
'Data check'
        If i < 2 Then
            i = 2
        End If
        
        If j < i Then
            j = i
        End If
        
        If j > ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row Then
            j = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        End If
        
'Execute'
        Call convertBulk(i, j)
    
End Sub
