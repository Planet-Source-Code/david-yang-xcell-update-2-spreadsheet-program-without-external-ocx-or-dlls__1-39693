Attribute VB_Name = "modMain"
'Public MyApplication As New CWorkbook
Private mya As CEquasion
Private aaa As CData

Sub Main()
    'XCell main code================================
    '===============================================
    Set MyApplication = New CWorkbook
    MyApplication.ShowBook
    MyApplication.Range("A1").FontBold = True
    MyApplication.Range("A1").Value = "Welcome to XCell"
    MyApplication.RedrawGrid
    
    'To check Equasion Class =======================
    '===============================================
    'Set mya = New CEquasion
    'Set aaa = New CData
    'Call aaa.Initalize
    'aaa.CellValue(1, 1) = 12
    'aaa.CellValue(2, 2) = 45
    'Call mya.Initilize(aaa)
    'MsgBox mya.Evaluate("(1+1)+35/2*(A1+B2)/(2*2)")
End Sub
