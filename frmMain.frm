VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   ":: DS Xcell ::                                                      V 0.1 Alpha"
   ClientHeight    =   4215
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6660
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6660
      TabIndex        =   3
      Top             =   0
      Width           =   6660
      Begin VB.TextBox txtCellValue 
         Height          =   285
         Left            =   1000
         TabIndex        =   8
         Top             =   420
         Width           =   3015
      End
      Begin MSComctlLib.Toolbar tbrMenu 
         Height          =   390
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               ImageIndex      =   4
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               ImageIndex      =   5
               Style           =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underlined"
               ImageIndex      =   6
               Style           =   1
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   3000
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Exit"
            EndProperty
         EndProperty
         Begin VB.ComboBox cboFontSize 
            Height          =   315
            Left            =   4560
            TabIndex        =   6
            Text            =   "10"
            Top             =   50
            Width           =   615
         End
         Begin VB.ComboBox cboFonts 
            Height          =   315
            ItemData        =   "frmMain.frx":0000
            Left            =   2220
            List            =   "frmMain.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   5
            Text            =   "Arial"
            Top             =   50
            Width           =   2295
         End
      End
      Begin VB.Label lblEqualsPic 
         Caption         =   "Cell ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   375
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0358
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "Arial"
      FontSize        =   10
   End
   Begin VB.VScrollBar VScroll 
      Height          =   3015
      LargeChange     =   5
      Left            =   6120
      Max             =   180
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   120
      Max             =   90
      TabIndex        =   1
      Top             =   3840
      Width           =   6015
   End
   Begin VB.PictureBox Grid 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      FillColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   840
      Width           =   5895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuFont 
         Caption         =   "Cell Font"
      End
      Begin VB.Menu mnuCellColor 
         Caption         =   "Cell Color"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyGrid As New CGrid
Public Event ChangeValue(Value As String, Cell As String)

Private Sub cboFonts_Click()
    'TODO
    MyGrid.MyData.FontName(1, 1) = cboFonts.Text
    MyGrid.RedrawGrid
End Sub

Private Sub cboFontSize_Click()
    'TODO
    MyGrid.MyData.FontSize(1, 1) = cboFontSize.Text
    MyGrid.RedrawGrid
End Sub

Private Sub Form_Load()

    Call MyGrid.Initalize(Grid)
    
    'Add fonts to fill the fonts combo
    For a = 0 To Screen.FontCount - 1
        Call cboFonts.AddItem(Screen.Fonts(a)) ', B)
    Next a
    
    'Add font sizes
    For a = 8 To 15
        cboFontSize.AddItem (a)
    Next a
    
    For a = 15 To 30 Step 2
        cboFontSize.AddItem (a)
    Next a
End Sub

Private Sub Form_Resize()
    'Resize all controls
    Grid.Width = Me.Width - 345 - VScroll.Width
    Grid.Height = Me.Height - 780 - HScroll.Height - Grid.Top
    
    HScroll.Top = Me.Height - 970
    HScroll.Width = Me.Width - HScroll.Left - VScroll.Width - 300
    HScroll.Max = 50 - Int(Grid.ScaleWidth / 800) - 1
    
    VScroll.Left = Me.Width - 450
    VScroll.Height = Me.Height - VScroll.Top - HScroll.Height - 800
    VScroll.Max = 200 - Int(Grid.ScaleHeight / 300) - 2
   
End Sub

Private Sub HScroll_Change()
    Call MyGrid.HorizScroll(HScroll.Value)
End Sub

Private Sub mnuCellColor_Click()
    dlgMain.DialogTitle = "Select Color"
    dlgMain.ShowColor
    'TODO
    MyGrid.MyData.CellColor(1, 1) = dlgMain.Color
    MyGrid.RedrawGrid
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Bold"
            'MyGrid.GridData(SelectedX, SelectedY).Bold = Not (GridData(SelectedX, SelectedY).Bold)
        Case "Italic"
            'MyGrid.GridData(SelectedX, SelectedY).Italic = Not (GridData(SelectedX, SelectedY).Italic)
        Case "Underlined"
            'MyGrid.GridData(SelectedX, SelectedY).Underlined = Not (GridData(SelectedX, SelectedY).Underlined)
        Case "Exit"
            Unload Me
            Unload frmSplash
        Case "New"
            MyGrid.MyData.Initalize
            MyGrid.RedrawGrid
    End Select
    MyGrid.RedrawGrid
End Sub

Private Sub txtCellValue_Change()
    MyGrid.CurrValue = txtCellValue.Text
End Sub

Private Sub txtCellValue_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Grid.SetFocus
    End If
End Sub

Private Sub VScroll_Change()
    Call MyGrid.VertScroll(VScroll.Value)
End Sub

'Private Access Code
Public Function GetCurrValue(x As Integer, y As Integer)
    GetCurrValue = MyGrid.MyData.CellValue(x, y)
End Function

Public Function GetCurrColor(x As Integer, y As Integer)
    GetCurrColor = MyGrid.MyData.CellColor(x, y)
End Function

Public Function GetCurrFontName(x As Integer, y As Integer)
    GetCurrFontName = MyGrid.MyData.FontName(x, y)
End Function

Public Function GetCurrFontColor(x As Integer, y As Integer)
    GetCurrFontColor = MyGrid.MyData.CellColor(x, y)
End Function

Public Function GetCurrFontUnderlined(x As Integer, y As Integer)
    GetCurrFontUnderlined = MyGrid.MyData.Underlined(x, y)
End Function

Public Function GetCurrFontItalic(x As Integer, y As Integer)
    GetCurrFontItalic = MyGrid.MyData.Italic(x, y)
End Function

Public Function GetCurrFontBold(x As Integer, y As Integer)
    GetCurrFontBold = MyGrid.MyData.Bold(x, y)
End Function

'Private access for setting values
Public Sub SetCurrValue(x As Integer, y As Integer, SetValue)
    MyGrid.MyData.CellValue(x, y) = SetValue
End Sub

Public Sub SetCurrColor(x As Integer, y As Integer, SetValue)
    MyGrid.MyData.CellColor(x, y) = SetValue
End Sub

Public Sub SetCurrFontName(x As Integer, y As Integer, SetValue)
    MyGrid.MyData.FontName(x, y) = SetValue
End Sub

Public Sub SetCurrFontColor(x As Integer, y As Integer, SetValue)
    MyGrid.MyData.CellColor(x, y) = SetValue
End Sub

Public Sub SetCurrFontUnderlined(x As Integer, y As Integer, SetValue)
    MyGrid.MyData.Underlined(x, y) = SetValue
End Sub

Public Sub SetCurrFontItalic(x As Integer, y As Integer, SetValue)
    MyGrid.MyData.Italic(x, y) = SetValue
End Sub

Public Sub SetCurrFontBold(x As Integer, y As Integer, SetValue)
    MyGrid.MyData.Bold(x, y) = SetValue
End Sub
