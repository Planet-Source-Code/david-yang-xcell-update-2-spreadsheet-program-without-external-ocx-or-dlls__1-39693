Attribute VB_Name = "modDeclares"
Public Type MyData
    CellColor As Long
    CellValue As String
    FontColor As Long
    Bold As Boolean
    Italic As Boolean
    Underlined As Boolean
    FontName As String
    FontSize As Integer
End Type

'Getting the width of text, so carets can be drawn
Public Type Size
        cx As Long
        cy As Long
End Type

Public Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long

Public TextExtent As Size
