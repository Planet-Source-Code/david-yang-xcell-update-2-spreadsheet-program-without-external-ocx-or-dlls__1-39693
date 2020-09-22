VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEquasion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private MyData As CData

Public Function Initilize(Data As CData)
    Set MyData = Data
End Function

Public Function Evaluate(Formula As String)
    OldFormula = Formula
    'Turn the A1 and B2 into values
    Formula = DereferenceValue(Formula)
    'Evaluate all functions within the brackets
    Do
        Debug.Print Formula
        'Loop through all characters, when it finds
        'a ")" then it exits the loop, with the location
        'of the "("
        Bracket = False
        FirstPos = 1
        SecondPos = Len(Formula)
        For a = 1 To Len(Formula)
            temp = Mid(Formula, a, 1)
            If temp = ")" Then
                SecondPos = a
                Bracket = True
                Exit For
            ElseIf temp = "(" Then
                FirstPos = a
            End If
        Next a
        'Substitute the formula inside the bracket for the
        'value of the formula, if there is brackets, then evaluate
        'string needs the brackets removed, so slignt
        'changes in the MID$ function
        If Bracket = True Then
            Formula = Left(Formula, FirstPos - 1) & EvaluateString(Mid(Formula, FirstPos + 1, SecondPos - FirstPos - 1)) & Right(Formula, Len(Formula) - SecondPos)
        Else
            Formula = Left(Formula, FirstPos - 1) & EvaluateString(Mid(Formula, FirstPos, SecondPos - FirstPos + 1)) & Right(Formula, Len(Formula) - SecondPos)
        End If
    Loop Until NoFormula(Formula) = True
    Debug.Print OldFormula & " = " & Formula
    Evaluate = Formula
End Function

Private Function NoFormula(Formula As String)
    NoFormula = True
    For a = 1 To Len(Formula)
        If Mid(Formula, a, 1) = "(" Then
            NoFormula = False
            Exit For
        End If
    Next a
    For a = 2 To Len(Formula)
        temp = Mid(Formula, a, 1)
        If temp = "+" Or temp = "-" Or temp = "*" Or temp = "/" Then
            NoFormula = False
            Exit For
        End If
    Next a
End Function

Private Function FindBracket(Formula As String)
    For a = 1 To Len(Formula)
        temp = Mid(Formula, a, 1)
        If temp = "(" Then
            FindBracket = a
        ElseIf temp = ")" Then
            FindBracket = a
        End If
    Next a
End Function

Private Function DereferenceValue(Formula As String)
    On Error GoTo ErrorHandler
    
    Formula = UCase(Formula)
    'Finds all values like A5 and AC2 and turns it into values
    For a = 1 To Len(Formula)
        'Enter this section if it spots a character
        If Mid(Formula, a, 1) >= "A" And Mid(Formula, a, 1) <= "Z" Then
            'Loops through all characters after
            ' the initial character to search for a digit
            For b = a To Len(Formula)
                If Mid(Formula, b, 1) <= 9 And Mid(Formula, b, 1) >= 0 Then
                    'When a digit is found, find the
                    ' location of the end of the reference
                    'e.g.   34 + AZ56
                    '            ^ ^ ^
                    '(a)---------| | |-------ChangePos2 (c)
                    '(b)-----------|
                    ' set the changepos2 to the length at the
                    ' beginning to prevent errors if the
                    ' cell reference is at the end of the string
                    
                    ChangePos2 = Len(Formula)
                    For c = b To Len(Formula)
                        If Not (Mid(Formula, c, 1) <= 9 And Mid(Formula, c, 1) >= 0) Then
                            ChangePos2 = c - 1
                            Exit For
                        End If
                    Next c
                    Exit For
                End If
            Next b
            Texts = ABC2Number(Mid(Formula, a, b - a))
            Numbers = Mid(Formula, b, ChangePos2 - b + 1)
            If Left(MyData.CellValue(CInt(Texts), CInt(Numbers)), 1) = "=" Then
                tempValue = Evaluate(Right(MyData.CellValue(CInt(Texts), CInt(Numbers)), Len(MyData.CellValue(CInt(Texts), CInt(Numbers))) - 1))
            Else
                tempValue = CDbl(MyData.CellValue(CInt(Texts), CInt(Numbers)))
            End If
            Formula = Left(Formula, a - 1) & tempValue & Right(Formula, Len(Formula) - ChangePos2)
            a = b
        End If
    Next a
    DereferenceValue = Formula
    GoTo EndIt
ErrorHandler:
    'MsgBox "There was a error in referencing the cells:" & Chr$(13) & "Make sure that the following cell" & Chr(13) & "is nonblank and numerical:  " & Number2ABC(CInt(Texts)) & Numbers, vbCritical + vbOKOnly, "Dereferencing Error"
    'Formula = "Bad Cell Reference: " & Number2ABC(CInt(Texts)) & Numbers
EndIt:
End Function

Private Function EvaluateString(Formula As String)
    'This is DIFFERENT from Evaluate function
    'because it does not consider brackets
    'Basically, this algorithm searches through every character
    'in the string, if it finds a * or / then a flag is flagged
    'since we are using order of operations.
    'It stores the location of the operators - +,-,*,/
    'then it does the operation
    'this is done untill there are no signs left except for
    'a "-" at the front of the formula
    On Error GoTo ErrorTrap
    
    Dim NextGet As Boolean
    Dim OperatorLoc1 As Integer, OperatorLoc2 As Integer
    Dim OperatorLoc3 As Integer
    Dim CurrChar As String
    
    OldFormula = Formula
    
    Do While CheckFormulaSigns(Formula)
        Debug.Print "    " & Formula
        NextGet = False
        OperatorLoc1 = 0
        OperatorLoc2 = 0
        OperatorLoc3 = Len(Formula) + 1
        
        For a = 1 To Len(Formula)
            CurrChar = Mid(Formula, a, 1)
            If CurrChar = "+" Or CurrChar = "-" Or CurrChar = "*" Or CurrChar = "/" Then
                'check if it is like this 4*-2
                If NextGet = True Then
                    'If CurrChar <> "-" And Trim(Mid(Formula, OperatorLoc2 + 1, a - OperatorLoc2)) <> "" Then
                        OperatorLoc3 = a
                        Exit For
                    'End If
                Else
                    'If CurrChar <> "-" And Trim(Mid(Formula, OperatorLoc2 + 1, a - OperatorLoc2)) <> "" And First = False Then
                        OperatorLoc1 = OperatorLoc2
                        OperatorLoc2 = a
                        If CurrChar = "*" Or CurrChar = "/" Then
                            NextGet = True
                            'tells it to start to look for the last
                            'operator
                        End If
                        'First = True
                    'End If
                End If
            End If
        Next a
        
        'These are the values of the numbers before
        'and after the operators
        TempValue1 = Mid(Formula, OperatorLoc1 + 1, OperatorLoc2 - OperatorLoc1 - 1)
        TempValue2 = Mid(Formula, OperatorLoc2 + 1, OperatorLoc3 - OperatorLoc2 - 1)
        
        'Does the appropriate job, stores it in TotalValue
        Select Case Mid(Formula, OperatorLoc2, 1)
            Case "*"
                TotalValue = CDbl(TempValue1) * CDbl(TempValue2)
            Case "/"
                TotalValue = CDbl(TempValue1) / CDbl(TempValue2)
            Case "+"
                TotalValue = CDbl(TempValue1) + CDbl(TempValue2)
            Case "-"
                TotalValue = CDbl(TempValue1) - CDbl(TempValue2)
        End Select
        
        If Abs(CDbl(TempValue1) - CDbl(TempValue2)) = CDbl(TempValue1) - CDbl(TempValue2) Then
        End If
        Formula = Mid(Formula, 1, OperatorLoc1) & TotalValue & Mid(Formula, OperatorLoc3, Len(Formula) - OperatorLoc3 + 1)
    Loop
    
    Debug.Print "  "; OldFormula & " = " & Formula
    EvaluateString = Formula
    GoTo ExitThis
ErrorTrap:
    Debug.Print "#Num"
ExitThis:
End Function

Private Function CheckFormulaSigns(Formula As String)
    CheckFormulaSigns = False
    For a = 2 To Len(Formula)
        temp = Mid(Formula, a, 1)
        If temp = "+" Or temp = "-" Or temp = "*" Or temp = "/" Then
            CheckFormulaSigns = True
            Exit For
        End If
    Next a
End Function
