Attribute VB_Name = "modUtils"
Public Function ABC2Number(ABC As String)
    'Use the method that is used to convert bases to
    'convert the Column references to a integer
    ' e.g. A => 1, Z => 26, AA => 27
    ' if you dont get this read a maths book
    temp = 0
    For a = 1 To Len(ABC)
        temp = temp + (26 ^ (a - 1)) * (Asc(Mid(ABC, Len(ABC) - a + 1, 1)) - 64)
    Next a
    ABC2Number = temp
End Function

Public Function Number2ABC(Number As Integer)
    Do While Number >= 1
        Characters = Chr((Number - 1) Mod 26 + 65) & Characters
        Number = Int(Number / 27)
    Loop
    Number2ABC = Characters
End Function

