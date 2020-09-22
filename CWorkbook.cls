VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Book As frmMain
Private MyRange As CRange

Public Sub ShowBook()
    Book.Show
    Call MyRange.SetForm(Book)
End Sub

Public Sub ResizeForm(Top As Integer, Left As Integer, Height As Integer, Width As Integer)
    Book.Top = Top
    Book.Left = Left
    Book.Height = Height
    Book.Width = Width
End Sub

Property Get Range(Row As String) As CRange
    Call MyRange.SetCell(Row)
    Set Range = MyRange
End Property

Property Let Range(Row As String, NewValue As CRange)
    Call MyRange.SetCell(Row)
    Set NewValue = MyRange
End Property

Public Sub RedrawGrid()
    MyRange.RedrawGrid
End Sub

Private Sub Class_Initialize()
    Set Book = New frmMain
    Set MyRange = New CRange
End Sub

Private Sub Class_Terminate()
    Set Book = Nothing
    Set MyRange = Nothing
End Sub
