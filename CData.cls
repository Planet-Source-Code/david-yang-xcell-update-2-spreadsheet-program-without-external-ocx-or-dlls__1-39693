VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Data(1 To 50, 1 To 200) As MyData

Public Sub Initalize()
    'Set all values of cells
    For x = 1 To 50
        For y = 1 To 200
            With Data(x, y)
                .CellValue = ""
                .CellColor = &HFFFFFF
                .FontColor = &H0
                .Bold = False
                .Italic = False
                .Underlined = False
                .FontName = "Arial"
                .FontSize = 10
            End With
        Next y
    Next x
End Sub

Property Get CellValue(x As Integer, y As Integer) As String
    CellValue = Data(x, y).CellValue
End Property

Property Let CellValue(x As Integer, y As Integer, SetValue As String)
    Data(x, y).CellValue = SetValue
End Property

Property Get CellColor(x As Integer, y As Integer) As Long
    CellColor = Data(x, y).CellColor
End Property

Property Let CellColor(x As Integer, y As Integer, SetValue As Long)
    Data(x, y).CellColor = SetValue
End Property

Property Get Bold(x As Integer, y As Integer) As Boolean
    Bold = Data(x, y).Bold
End Property

Property Let Bold(x As Integer, y As Integer, SetValue As Boolean)
    Data(x, y).Bold = SetValue
End Property

Property Get Italic(x As Integer, y As Integer) As Boolean
    Italic = Data(x, y).Italic
End Property

Property Let Italic(x As Integer, y As Integer, SetValue As Boolean)
    Data(x, y).Italic = SetValue
End Property

Property Get Underlined(x As Integer, y As Integer) As Boolean
    Underlined = Data(x, y).Underlined
End Property

Property Let Underlined(x As Integer, y As Integer, SetValue As Boolean)
    Data(x, y).Underlined = SetValue
End Property

Property Get FontSize(x As Integer, y As Integer) As Integer
    FontSize = Data(x, y).FontSize
End Property

Property Let FontSize(x As Integer, y As Integer, SetValue As Integer)
    Data(x, y).FontSize = SetValue
End Property

Property Get FontName(x As Integer, y As Integer) As String
    FontName = Data(x, y).FontName
End Property

Property Let FontName(x As Integer, y As Integer, SetValue As String)
    Data(x, y).FontName = SetValue
End Property

Property Get FontColor(x As Integer, y As Integer) As Long
    FontColor = Data(x, y).FontColor
End Property

Property Let FontColor(x As Integer, y As Integer, SetValue As Long)
    Data(x, y).FontColor = SetValue
End Property

Private Sub Class_Terminate()
    Erase Data
End Sub
