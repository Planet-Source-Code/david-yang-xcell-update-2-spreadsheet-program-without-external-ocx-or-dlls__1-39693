VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private CurrForm As frmMain
Attribute CurrForm.VB_VarHelpID = -1
Private CurrRow As Integer, CurrColumn As Integer

Sub SetForm(MyForm As Form)
    Set CurrForm = MyForm
End Sub

Sub SetCell(CellName As String)
    
    CellName = UCase(CellName)
    For a = 1 To Len(CellName)
        If Mid(CellName, a, 1) >= "A" And Mid(CellName, a, 1) <= "Z" Then
            CurrPos = a
        Else
            Exit For
        End If
    Next a
    
    TempCurrColumn = Left(CellName, CurrPos)
    CurrColumn = ABC2Number(CStr(TempCurrColumn))
    
    CurrRow = Right(CellName, Len(CellName) - CurrPos)
End Sub

'Property Gets
Property Get Value()
    Value = CurrForm.GetCurrValue(CurrRow, CurrColumn)
End Property

Property Get Color()
    Color = CurrForm.GetCurrColor(CurrRow, CurrColumn)
End Property

Property Get FontColor()
    FontColor = CurrForm.GetCurrFontColor(CurrRow, CurrColumn)
End Property

Property Get FontName()
    FontName = CurrForm.GetCurrFontName(CurrRow, CurrColumn)
End Property

Property Get FontUnderlined()
    FontUnderlined = CurrForm.GetCurrFontUnderlined(CurrRow, CurrColumn)
End Property

Property Get FontItalic()
    FontItalic = CurrForm.GetCurrFontItalic(CurrRow, CurrColumn)
End Property

Property Get FontBold()
    FontBold = CurrForm.GetCurrFontBold(CurrRow, CurrColumn)
End Property

'Property Lets
Property Let Value(SetValue)
    Call CurrForm.SetCurrValue(CurrRow, CurrColumn, SetValue)
End Property

Property Let Color(SetValue)
    Call CurrForm.SetCurrColor(CurrRow, CurrColumn, SetValue)
End Property

Property Let FontColor(SetValue)
    Call CurrForm.SetCurrFontColor(CurrRow, CurrColumn, SetValue)
End Property

Property Let FontName(SetValue)
    Call CurrForm.SetCurrFontName(CurrRow, CurrColumn, SetValue)
End Property

Property Let FontUnderlined(SetValue)
    Call CurrForm.SetCurrFontUnderlined(CurrRow, CurrColumn, SetValue)
End Property

Property Let FontItalic(SetValue)
    Call CurrForm.SetCurrFontItalic(CurrRow, CurrColumn, SetValue)
End Property

Property Let FontBold(SetValue)
    Call CurrForm.SetCurrFontBold(CurrRow, CurrColumn, SetValue)
End Property

Public Sub RedrawGrid()
    CurrForm.MyGrid.RedrawGrid
End Sub

Private Sub CurrForm_ChangeValue(Value As String, Cell As String)
    Debug.Print Value
End Sub

Private Sub Class_Terminate()
    Set CurrForm = Nothing
End Sub
