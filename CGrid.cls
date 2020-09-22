VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents MyGrid As PictureBox
Attribute MyGrid.VB_VarHelpID = -1
Private MyEquasion As CEquasion
Public MyData As CData
Private XPos As Integer, YPos As Integer
Private SelectedX As Integer, SelectedY As Integer
Private DragX As Integer, DragY As Integer
Private Drag As Boolean
Private LeftBound As Integer, RightBound As Integer
Private TopBound As Integer, BottomBound As Integer
Private EditMode As Boolean
Private CurrTextPos As Integer

'Public Sub & Function ==============================
'====================================================
Public Sub Initalize(DrawSurface As PictureBox)
    Set MyGrid = DrawSurface
    Set MyData = New CData
    Set MyEquasion = New CEquasion
    Call MyData.Initalize
    Call MyEquasion.Initilize(MyData)
    
    'current scroll position
    XPos = 0
    YPos = 0
    
    'current selected position
    SelectedX = 1
    SelectedY = 1
    
End Sub

Public Sub RedrawGrid()
    
    aaa = Timer
    MyGrid.Cls
    
    MyGrid.DrawWidth = 1
    MyGrid.ForeColor = &H8000000F
    MyGrid.DrawMode = 13
    
    'This states the view that the form shows
    LeftBound = XPos + 1
    RightBound = Int(MyGrid.Width / 800) + 1 + XPos
    TopBound = YPos + 1
    BottomBound = Int(MyGrid.Height / 300) + 1 + YPos
    
    'Draw backgrounds for each cell
    For x = LeftBound To RightBound
        For y = TopBound To BottomBound
            Call DrawBackground(CInt(x), CInt(y))
        Next y
    Next x

    'Fill data for each cell
    For x = LeftBound To RightBound
        For y = TopBound To BottomBound
            If MyData.CellValue(CInt(x), CInt(y)) <> "" Then
                Call DrawCell(CInt(x), CInt(y))
            Else
                'Call DrawBackground(CInt(x), CInt(y))
            End If
        Next y
    Next x
    
    MyGrid.ForeColor = &H8000000F
    'Draw Borders
    For x = LeftBound To RightBound
        MyGrid.Line ((x - XPos) * 800, 0)-((x - XPos) * 800, MyGrid.ScaleHeight)
    Next x
    
    For y = TopBound To BottomBound
        MyGrid.Line (0, (y - YPos) * 300)-(MyGrid.ScaleWidth, (y - YPos) * 300)
    Next y
    
    'Change font back
    Call ReturnNormalFont
    
    'Draw the side indicators
    DrawSideIndicator

    'Draw the top indicators
    DrawTopIndicator
    
'Misc processing
    
    'Draw Selected Box
    If Not (SelectedX - XPos < 1 Or SelectedX - XPos > MyGrid.Width Or SelectedY - YPos < 1 Or SelectedY - YPos > MyGrid.Height) Then
        Call DrawBox((SelectedX - XPos) * 800, (SelectedY - YPos) * 300, (SelectedX + 1 - XPos) * 800, (SelectedY + 1 - YPos) * 300)
    End If
    
    'Change the toolbars according to the cell properties
    Call ChangeToolbarStatus
    
    'Move the scrollbars to the correct positions
    'VScroll.Value = YPos
    'HScroll.Value = XPos
    
    'Put the correct values into the equasion box
    'txtCellValue.Text = data(SelectedX, SelectedY).CellValue
    
    'Draw the caret
    Call SetFont(SelectedX, SelectedY)
    If EditMode = True Then
        Call GetTextExtentPoint(MyGrid.hdc, MyData.CellValue(SelectedX, SelectedY), CurrTextPos, TextExtent)
        MyGrid.ForeColor = vbBlack
        MyGrid.DrawWidth = 1
        MyGrid.Line ((SelectedX - XPos) * 800 + TextExtent.cx * 15 + 90, (SelectedY - YPos) * 300 + 30)-((SelectedX - XPos) * 800 + TextExtent.cx * 15 + 90, (SelectedY - YPos) * 300 + TextExtent.cy * 15 + 30)
    End If
    aaa = Timer - aaa
    MyGrid.CurrentX = 100
    MyGrid.CurrentY = 100
    MyGrid.ForeColor = vbBlack
    MyGrid.Print aaa
    'DoEvents
End Sub

Public Sub VertScroll(Position As Integer)
    YPos = Position
    RedrawGrid
End Sub

Public Sub HorizScroll(Position As Integer)
    XPos = Position
    RedrawGrid
End Sub

Property Get CurrValue() As String
    CurrValue = MyData.CellValue(SelectedX, SelectedY)
End Property

Property Let CurrValue(SetValue As String)
    MyData.CellValue(SelectedX, SelectedY) = SetValue
    RedrawGrid
End Property

'Private Subs & Functions=============================
'=====================================================
Private Sub DrawCell(x As Integer, y As Integer)
    'Fill background again, to prevent the previous text
    Call DrawBackground(CInt(x), CInt(y))
    'Set the position to print to
    MyGrid.CurrentX = (x - XPos) * 800 + 100
    MyGrid.CurrentY = (y - YPos) * 300 + 50
    'Get the font right
    MyGrid.ForeColor = MyData.FontColor(x, y)
    Call SetFont(CInt(x), CInt(y))
    'All this just to print the cell's text!!
    If Left(MyData.CellValue(x, y), 1) = "=" And EditMode = False Then
        MyGrid.Print MyEquasion.Evaluate(Right(MyData.CellValue(x, y), Len(MyData.CellValue(x, y)) - 1))
    Else
        MyGrid.Print MyData.CellValue(x, y)
    End If
End Sub

Private Sub SetFont(x As Integer, y As Integer)
    With MyGrid.Font
        .Bold = MyData.Bold(x, y)
        .Italic = MyData.Italic(x, y)
        .Underline = MyData.Underlined(x, y)
        .Name = MyData.FontName(x, y)
        .Size = MyData.FontSize(x, y)
    End With
End Sub

Private Sub DrawBackground(x As Integer, y As Integer)
    MyGrid.ForeColor = MyData.CellColor(x, y)
    MyGrid.Line ((x - XPos) * 800 + 15, (y - YPos) * 300 + 15)-((x + 1 - XPos) * 800 - 15, (y + 1 - YPos) * 300 - 15), , BF
End Sub

Private Sub DrawSideIndicator()
    'Draw grey side indicator
    MyGrid.ForeColor = &H8000000F
    MyGrid.Line (0, 0)-(800, MyGrid.Height), , BF
    
    'Draw grey side indicator border
    MyGrid.ForeColor = &H8000000C
    MyGrid.Line (800, 0)-(800, MyGrid.ScaleHeight)
    
    For y = 1 To Int(MyGrid.Height / 300) + 1
        
        'Draw dark grey borders in side indicator to
        'seperate rows and columns
        MyGrid.ForeColor = &H8000000C
        MyGrid.Line (0, y * 300)-(800, y * 300)
        
        'Print the row numbers
        MyGrid.ForeColor = &H0
        MyGrid.CurrentX = 10
        MyGrid.CurrentY = (y) * 300
        MyGrid.Print y + YPos
    Next y
End Sub

Private Sub DrawTopIndicator()
    'Draw grey top indicator
    MyGrid.ForeColor = &H8000000F
    MyGrid.Line (0, 0)-(MyGrid.Width, 300), , BF
    
    'Draw grey top indicator border
    MyGrid.ForeColor = &H8000000C
    MyGrid.Line (0, 300)-(MyGrid.ScaleWidth, 300)
    
    For x = 1 To Int(MyGrid.Width / 800) + 1
        'Draw dark grey borders to seperate rows and columns
        MyGrid.ForeColor = &H8000000C
        MyGrid.Line (x * 800, 0)-(x * 800, 300)
        
        'Print the column numbers
        MyGrid.ForeColor = &H0
        MyGrid.CurrentX = x * 800
        MyGrid.CurrentY = 0

        MyGrid.Print Number2ABC(x + XPos)
    Next x
End Sub

Private Sub ReturnNormalFont()
    With MyGrid.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .Name = "Arial"
        .Size = 10
    End With
End Sub
Private Sub DrawBox(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional BkColour As Long = &H0)
    MyGrid.FillStyle = 1
    MyGrid.ForeColor = &H0
    MyGrid.DrawWidth = 2
    If BkColour <> &H0 Then
        MyGrid.ForeColor = &HFFFFFF
        MyGrid.FillStyle = 5
        MyGrid.FillColor = BkColour
        MyGrid.Line (X1, Y1)-(X2, Y2), , BF
    End If
    MyGrid.Line (X1, Y1)-(X2, Y2), , B
    'Still to do
End Sub

Private Sub CheckXY()
    'Check the position of cursor
    If SelectedX > 50 Then
        SelectedX = 50
    End If
    
    If SelectedX < 1 Then
        SelectedX = 1
    End If
    
    If SelectedY > 200 Then
        SelectedY = 200
    End If
    
    If SelectedY < 1 Then
        SelectedY = 1
    End If
    
    'Scroll MyGrid if neccessary
    If SelectedX - XPos <= 0 Then
        XPos = XPos - 1
        RedrawGrid
    End If
    
    If SelectedY - YPos <= 0 Then
        YPos = YPos - 1
        RedrawGrid
    End If
    
    If SelectedY - YPos > MyGrid.Height / 300 Then
        YPos = YPos + 1
        RedrawGrid
    End If
    
    If SelectedX - XPos > MyGrid.Width / 800 Then
        XPos = XPos + 1
        RedrawGrid
    End If
End Sub

Private Sub SetCursorPos(x As Single, y As Single)
    'This routine is used if the use clicks while they
    ' are editing text, and they click within the textbox
    Call SetFont(SelectedX, SelectedY)
    For CurrChar = 1 To Len(MyData.CellValue(SelectedX, SelectedY))
    Call GetTextExtentPoint(MyGrid.hdc, MyData.CellValue(SelectedX, SelectedY), CurrChar + 1, TextExtent)
        If TextExtent.cx * 15 + (SelectedX - XPos) * 800 > x Then
            CurrTextPos = CurrChar
            Exit For
        End If
    Next CurrChar
    RedrawGrid
End Sub

Private Sub FindClickedCell(x As Single, y As Single)
    'This routine finds which cell you clicked on
    DragX = Int(x / 800) * 800
    DragY = Int(y / 300) * 300
    If DragX >= 800 Then
        If DragY >= 300 Then
            If SelectedX = Int(x / 800) + XPos And SelectedY = Int(y / 300) + YPos Then
                Call SetCursorPos(x, y)
                'Bug here
            Else
                SelectedX = Int(x / 800) + XPos
                SelectedY = Int(y / 300) + YPos
            End If
        Else
            MsgBox "Clicked on a column" & DragX / 800
        End If
    Else
        MsgBox "clicked on a row" & DragY / 800
    End If
    EditMode = False
End Sub

Private Sub ChangeToolbarStatus()
    'Change the buttons to suit the cell being selected
    If MyData.Bold(SelectedX, SelectedY) Then
        'tbrMenu.Buttons(5).Value = tbrPressed
    Else
        'tbrMenu.Buttons(5).Value = tbrUnpressed
    End If
    
    If MyData.Italic(SelectedX, SelectedY) Then
        'tbrMenu.Buttons(6).Value = tbrPressed
    Else
        'tbrMenu.Buttons(6).Value = tbrUnpressed
    End If
    
    If MyData.Underlined(SelectedX, SelectedY) Then
        'tbrMenu.Buttons(7).Value = tbrPressed
    Else
        'tbrMenu.Buttons(7).Value = tbrUnpressed
    End If
    
    'cboFontSize.Text = data(SelectedX, SelectedY).FontSize
    'cboFonts.Text = data(SelectedX, SelectedY).FontName
End Sub

Private Sub Class_Initialize()
    Set MyData = New CData
End Sub

Private Sub Class_Terminate()
    Set MyData = Nothing
End Sub

Private Sub MyGrid_KeyPress(KeyAscii As Integer)
    'This checks keys pressed, and then does correct job
    CurrText = MyData.CellValue(SelectedX, SelectedY)
    If (KeyAscii >= 32 And KeyAscii <= 126) Then 'any letter that can be typed on keyboard
        
        If EditMode = True Then
            MyData.CellValue(SelectedX, SelectedY) = Left(CurrText, CurrTextPos) & Chr(KeyAscii) & Right(CurrText, Len(CurrText) - CurrTextPos)
            CurrTextPos = CurrTextPos + 1
        Else
            MyData.CellValue(SelectedX, SelectedY) = Chr(KeyAscii)
            EditMode = True
            CurrTextPos = 1
        End If
        
    ElseIf KeyAscii = 8 Then 'for backspace key
        If EditMode = True Then 'if editing text, remove a character
            If CurrTextPos <> 0 Then 'to prevent backspace beyond the first character
                MyData.CellValue(SelectedX, SelectedY) = Left(CurrText, CurrTextPos - 1) & Right(CurrText, Len(CurrText) - CurrTextPos)
                CurrTextPos = CurrTextPos - 1
            End If
        Else 'if not editing text, clear contents, and start editing text
            MyData.CellValue(SelectedX, SelectedY) = ""
            EditMode = True
            CurrTextPos = 0
        End If
        
    ElseIf KeyAscii = 13 Then 'for enter key
        SelectedY = SelectedY + 1
        EditMode = False
        
    Else
        Debug.Print "Unknown Key: " & KeyAscii
                    MyData.CellValue(SelectedX, SelectedY) = Left(CurrText, CurrTextPos) & Chr(KeyAscii) & Right(CurrText, Len(CurrText) - CurrTextPos)
            CurrTextPos = CurrTextPos + 1

    End If
    RedrawGrid
    
End Sub

Private Sub MyGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Checks if you press the left, right, up or down keys
    If KeyCode = 38 Then
        SelectedY = SelectedY - 1
        EditMode = False
    ElseIf KeyCode = 40 Then
        SelectedY = SelectedY + 1
        EditMode = False
    ElseIf KeyCode = 39 Then
        SelectedX = SelectedX + 1
        EditMode = False
    ElseIf KeyCode = 37 Then
        SelectedX = SelectedX - 1
        EditMode = False
    End If
    
    'Check selected x and y positions and reset them if out of bounds
    CheckXY
    
    RedrawGrid
End Sub

Private Sub MyGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Get the text size - i.e. the width and height of text in selected font
    Call GetTextExtentPoint(MyGrid.hdc, MyData.CellValue(SelectedX, SelectedY), Len(MyData.CellValue(SelectedX, SelectedY)), TextExtent)
    
    ' if the user click within the cell that is active
    ' and the text, then put the cursor there
    If EditMode = True And y > (SelectedY - YPos) * 300 And y < (SelectedY - YPos + 1) * 300 And x > (SelectedX - XPos) * 800 And x < (SelectedX - XPos) * 800 + TextExtent.cx * 15 Then
        Call SetCursorPos(x, y)
    Else
        Call FindClickedCell(x, y)
    End If
End Sub

Private Sub MyGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Drag = True Then
        RedrawGrid
        Call DrawBox(Int(x / 800 + 1) * 800, Int(y / 300 + 1) * 300, DragX - XPos, DragY - YPos, &HC0C0C0)
    End If
End Sub

Private Sub MyGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RedrawGrid
End Sub

Private Sub MyGrid_Resize()
    RedrawGrid
End Sub
