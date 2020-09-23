VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "ucTabStrip by Raul338"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   324
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   60
      ScaleHeight     =   1035
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   60
      Width           =   6675
      Begin VB.CheckBox chkCommonEvents 
         Caption         =   "Common Evnets"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   585
         Width           =   1455
      End
      Begin VB.CheckBox chkDebugPrint 
         Caption         =   "Debug.Print Events"
         Height          =   195
         Left            =   2820
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Item"
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   0
         Width           =   1035
      End
      Begin VB.HScrollBar hscWidth 
         Height          =   195
         LargeChange     =   20
         Left            =   4920
         Max             =   200
         Min             =   20
         SmallChange     =   5
         TabIndex        =   11
         Top             =   300
         Value           =   20
         Width           =   1755
      End
      Begin VB.CheckBox chkMinWidth 
         Caption         =   "Min tab width"
         Height          =   195
         Left            =   4920
         TabIndex        =   10
         Top             =   60
         Width           =   1275
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   60
         TabIndex        =   9
         Top             =   600
         Width           =   1395
      End
      Begin VB.CheckBox chkMultiline 
         Caption         =   "Multiline"
         Height          =   195
         Left            =   2820
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cmbAlign 
         Height          =   315
         Left            =   4980
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdSetText 
         Caption         =   "Set Text"
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Top             =   720
         Width           =   1035
      End
      Begin VB.CommandButton cmdGetItemText 
         Caption         =   "Get Text"
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   420
         Width           =   1035
      End
      Begin VB.TextBox txtItem 
         Height          =   285
         Left            =   60
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Insert Item"
         Default         =   -1  'True
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label lblAlign 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tab Align"
         Height          =   195
         Left            =   4980
         TabIndex        =   7
         Top             =   540
         Width           =   675
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insert Item"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   735
      End
   End
   Begin ucTabStripDemo.ucTabStrip ucTabStrip1 
      Height          =   3495
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6165
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "Tab 0###€Tab 1###€Tab 2###€Tab 3###€Tab 4"
      ControlCount    =   49
      CN0             =   "Command2"
      CT0             =   0
      CN1             =   "Command1"
      CT1             =   0
      CN2             =   "Command10"
      CT2             =   0
      CN3             =   "Command9"
      CT3             =   0
      CN4             =   "Command8"
      CT4             =   0
      CN5             =   "Command7"
      CT5             =   0
      CN6             =   "Command6"
      CT6             =   0
      CN7             =   "Command5"
      CT7             =   0
      CN8             =   "Command4"
      CT8             =   0
      CN9             =   "Command3"
      CT9             =   0
      CN10            =   "Command12"
      CT10            =   1
      CN11            =   "Command11"
      CT11            =   1
      CN12            =   "Command20"
      CT12            =   1
      CN13            =   "Command19"
      CT13            =   1
      CN14            =   "Command18"
      CT14            =   1
      CN15            =   "Command17"
      CT15            =   1
      CN16            =   "Command16"
      CT16            =   1
      CN17            =   "Command15"
      CT17            =   1
      CN18            =   "Command14"
      CT18            =   1
      CN19            =   "Command13"
      CT19            =   1
      CN20            =   "Command30"
      CT20            =   2
      CN21            =   "Command29"
      CT21            =   2
      CN22            =   "Command28"
      CT22            =   2
      CN23            =   "Command27"
      CT23            =   2
      CN24            =   "Command26"
      CT24            =   2
      CN25            =   "Command25"
      CT25            =   2
      CN26            =   "Command24"
      CT26            =   2
      CN27            =   "Command23"
      CT27            =   2
      CN28            =   "Command22"
      CT28            =   2
      CN29            =   "Command21"
      CT29            =   2
      CN30            =   "Command40"
      CT30            =   3
      CN31            =   "Command39"
      CT31            =   3
      CN32            =   "Command38"
      CT32            =   3
      CN33            =   "Command37"
      CT33            =   3
      CN34            =   "Command36"
      CT34            =   3
      CN35            =   "Command35"
      CT35            =   3
      CN36            =   "Command34"
      CT36            =   3
      CN37            =   "Command33"
      CT37            =   3
      CN38            =   "Command32"
      CT38            =   3
      CN39            =   "Command31"
      CT39            =   3
      CN40            =   "Command41(7"
      CT40            =   4
      CN41            =   "Command41(6"
      CT41            =   4
      CN42            =   "Command41(5"
      CT42            =   4
      CN43            =   "Command41(4"
      CT43            =   4
      CN44            =   "Command41(3"
      CT44            =   4
      CN45            =   "Command41(2"
      CT45            =   4
      CN46            =   "Command41(1"
      CT46            =   4
      CN47            =   "Command41(0"
      CT47            =   4
      CN48            =   "Command41(9"
      CT48            =   4
      CN49            =   "Command41(8"
      CT49            =   4
      bSaved          =   -1  'True
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   495
         Index           =   9
         Left            =   3000
         TabIndex        =   65
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   495
         Index           =   8
         Left            =   1440
         TabIndex        =   64
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   495
         Index           =   7
         Left            =   3000
         TabIndex        =   63
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   495
         Index           =   6
         Left            =   1440
         TabIndex        =   62
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   495
         Index           =   5
         Left            =   3000
         TabIndex        =   61
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   495
         Index           =   4
         Left            =   1440
         TabIndex        =   60
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   495
         Index           =   3
         Left            =   3000
         TabIndex        =   59
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   495
         Index           =   2
         Left            =   1440
         TabIndex        =   58
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   57
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   495
         Index           =   0
         Left            =   1440
         TabIndex        =   56
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command40 
         Caption         =   "Command40"
         Height          =   495
         Left            =   3240
         TabIndex        =   55
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Command39"
         Height          =   495
         Left            =   1680
         TabIndex        =   54
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command38 
         Caption         =   "Command38"
         Height          =   495
         Left            =   3240
         TabIndex        =   53
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command37 
         Caption         =   "Command37"
         Height          =   495
         Left            =   1680
         TabIndex        =   52
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command36 
         Caption         =   "Command36"
         Height          =   495
         Left            =   3240
         TabIndex        =   51
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command35 
         Caption         =   "Command35"
         Height          =   495
         Left            =   1680
         TabIndex        =   50
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command34 
         Caption         =   "Command34"
         Height          =   495
         Left            =   3240
         TabIndex        =   49
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command33 
         Caption         =   "Command33"
         Height          =   495
         Left            =   1680
         TabIndex        =   48
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Command32"
         Height          =   495
         Left            =   3240
         TabIndex        =   47
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Command31"
         Height          =   495
         Left            =   1680
         TabIndex        =   46
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Command30"
         Height          =   495
         Left            =   1680
         TabIndex        =   45
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Command29"
         Height          =   495
         Left            =   240
         TabIndex        =   44
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Command28"
         Height          =   495
         Left            =   1680
         TabIndex        =   43
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Command27"
         Height          =   495
         Left            =   240
         TabIndex        =   42
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Command26"
         Height          =   495
         Left            =   1680
         TabIndex        =   41
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Command25"
         Height          =   495
         Left            =   240
         TabIndex        =   40
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Command24"
         Height          =   495
         Left            =   1680
         TabIndex        =   39
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Command23"
         Height          =   495
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Command22"
         Height          =   495
         Left            =   1680
         TabIndex        =   37
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Command21"
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Command20"
         Height          =   495
         Left            =   4680
         TabIndex        =   35
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Command19"
         Height          =   495
         Left            =   3000
         TabIndex        =   34
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Command18"
         Height          =   495
         Left            =   4680
         TabIndex        =   33
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Command17"
         Height          =   495
         Left            =   3000
         TabIndex        =   32
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   495
         Left            =   4680
         TabIndex        =   31
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Command15"
         Height          =   495
         Left            =   3000
         TabIndex        =   30
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Command14"
         Height          =   495
         Left            =   4680
         TabIndex        =   29
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Command13"
         Height          =   495
         Left            =   3000
         TabIndex        =   28
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         Height          =   495
         Left            =   4680
         TabIndex        =   27
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Command11"
         Height          =   495
         Left            =   3000
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Command10"
         Height          =   495
         Left            =   1800
         TabIndex        =   25
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Left            =   1800
         TabIndex        =   23
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   495
         Left            =   1800
         TabIndex        =   21
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   495
         Left            =   1800
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Index           =   0
         Left            =   1800
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ===========================
' Necesario para los temas de windows
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControls Lib "COMCTL32" () As Long
Private hMod As Long

Private Sub cmdDelete_Click()
    Call ucTabStrip1.RemoveTab(ucTabStrip1.SelectedItem)
End Sub

Private Sub Form_Initialize()
    hMod = LoadLibraryA("shell32.dll")
    Call InitCommonControls
End Sub

Private Sub Form_Terminate()
    Call FreeLibrary(hMod)
End Sub

' ===========================
' UI
Private Sub txtItem_Change()
    cmdSetText.Enabled = (txtItem.text <> vbNullString) And ucTabStrip1.Count > 0
    cmdGetItemText.Enabled = ucTabStrip1.Count > 0
End Sub

Private Sub Form_Load()
    'With ucTabStrip1
        'Dim i As Integer
        'For i = 0 To 5
        '    Call .AddTab(i, "Item " & i)
        'Next
    'End With

    chkMinWidth.value = vbUnchecked
    hscWidth.Enabled = False
    ucTabStrip1.Left = 5
    Call Picture1.Move(0, 0)
    ucTabStrip1.Top = Picture1.Height
    txtItem.text = vbNullString
    'chkDebugPrint.value = vbChecked

    With cmbAlign
        Call .AddItem("1 - Top")
        Call .AddItem("2 - Bottom")
        ' Call .AddItem("3 - Left") ' Not recomeneded
        ' Call .AddItem("4 - Rigth (!)") ' Doesn't work on XP and later :(
        .ListIndex = 0
    End With
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = vbMinimized Then Exit Sub
    ucTabStrip1.Height = ScaleHeight - ucTabStrip1.Top - 5
    ucTabStrip1.Width = ScaleWidth - ucTabStrip1.Left - 5
End Sub

Private Sub chkMinWidth_Click()
    hscWidth.Enabled = chkMinWidth.value = vbChecked
    If chkMinWidth.value = vbChecked Then
        Call ucTabStrip1.SetMinTabWidth(hscWidth.value)
    Else
        Call ucTabStrip1.SetMinTabWidth(-1)
    End If
End Sub

'============================
' Propiedades y Eventos
Private Sub hscWidth_Scroll()
    Call hscWidth_Change
End Sub
Private Sub hscWidth_Change()
    Call ucTabStrip1.SetMinTabWidth(hscWidth.value)
End Sub

Private Sub chkMultiline_Click()
    ucTabStrip1.Multiline = (chkMultiline.value = vbChecked)
End Sub

Private Sub cmbAlign_Click()
    ucTabStrip1.Align = cmbAlign.ListIndex + 1
    chkMultiline.value = vbChecked
End Sub

Private Sub cmdAdd_Click()
    If txtItem.text <> vbNullString Then
        Call ucTabStrip1.AddTab(ucTabStrip1.Count, txtItem.text)
        txtItem.text = vbNullString
        txtItem.SetFocus
    Else
        Call ucTabStrip1.AddTab(ucTabStrip1.Count, "Item " & ucTabStrip1.Count)
    End If
End Sub

Private Sub cmdGetItemText_Click()
    txtItem.text = ucTabStrip1.ItemText(ucTabStrip1.SelectedItem)
End Sub

Private Sub cmdSetText_Click()
    ucTabStrip1.ItemText(ucTabStrip1.SelectedItem) = txtItem.text
End Sub

Private Sub cmdClear_Click()
    ucTabStrip1.Clear
End Sub

Private Sub ucTabStrip1_ChangingTab(Cancel As Boolean)
    If chkDebugPrint.value = vbChecked Then Debug.Print ">ChangingTab"
End Sub

Private Sub ucTabStrip1_Click()
    If chkCommonEvents.value = vbChecked Then Debug.Print ">Click"
End Sub

Private Sub ucTabStrip1_DblClick()
    If chkDebugPrint.value = vbChecked Then Debug.Print ">DblClick"
End Sub

Private Sub ucTabStrip1_KeyDown(KeyCode As Integer, Shift As Integer)
    If chkCommonEvents.value = vbChecked Then Debug.Print ">KeyDown", KeyCode, Shift
End Sub

Private Sub ucTabStrip1_KeyPress(KeyAscii As Integer)
    If chkCommonEvents.value = vbChecked Then Debug.Print ">KeyPress", KeyAscii
End Sub

Private Sub ucTabStrip1_KeyUp(KeyCode As Integer, Shift As Integer)
    If chkCommonEvents.value = vbChecked Then Debug.Print ">KeyUp", KeyCode, Shift
End Sub

Private Sub ucTabStrip1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If chkCommonEvents.value = vbChecked Then Debug.Print ">MouseDown", Button, Shift, X, y
End Sub

Private Sub ucTabStrip1_MouseEnter()
    If chkDebugPrint.value = vbChecked Then Debug.Print ">MouseEnter"
End Sub

Private Sub ucTabStrip1_MouseLeave()
    If chkDebugPrint.value = vbChecked Then Debug.Print ">MouseLeave"
End Sub

Private Sub ucTabStrip1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If chkCommonEvents.value = vbChecked Then Debug.Print ">MouseMove", Button, Shift, X, y
End Sub

Private Sub ucTabStrip1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If chkCommonEvents.value = vbChecked Then Debug.Print ">MouseUp", Button, Shift, X, y
End Sub

Private Sub ucTabStrip1_TabClick(ByVal lTab As Long)
    If chkDebugPrint.value = vbChecked Then Debug.Print ">TabClick", lTab
End Sub
Private Sub ucTabStrip1_TabRightClick(ByVal lTab As Long)
    If chkDebugPrint.value = vbChecked Then Debug.Print ">TabRightClick", lTab
End Sub
Private Sub ucTabStrip1_GotFocus()
    If chkDebugPrint.value = vbChecked Then Debug.Print ">Got focus"
End Sub
Private Sub ucTabStrip1_LostFocus()
    If chkDebugPrint.value = vbChecked Then Debug.Print ">Lost focus"
End Sub
