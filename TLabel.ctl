VERSION 5.00
Begin VB.UserControl TLabel 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   900
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   ScaleHeight     =   255
   ScaleWidth      =   900
   ToolboxBitmap   =   "TLabel.ctx":0000
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "TLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents fnt As StdFont
Attribute fnt.VB_VarHelpID = -1

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0
    Caption = lblText.Caption
End Property
Public Property Let Caption(ByVal text As String)
    lblText.Caption = text
    Call UserControl_Show
    Call PropertyChanged("Caption")
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Font() As StdFont
    Set Font = fnt
End Property
Public Property Set Font(value As StdFont)
    Set lblText.Font = value
    Set fnt = value
    Call PropertyChanged("Font")
    Call UserControl_Show
End Property

Private Sub fnt_FontChanged(ByVal PropertyName As String)
    Set lblText.Font = fnt
    Call PropertyChanged("Font")
    Call UserControl_Show
End Sub

Private Sub UserControl_Initialize()
    Set fnt = New StdFont
    Set lblText.Font = fnt
End Sub

Private Sub UserControl_InitProperties()
    lblText.Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", Caption)
    Set fnt = PropBag.ReadProperty("Font", Ambient.Font)
    Set lblText.Font = fnt
    Call UserControl_Show
End Sub

Private Sub UserControl_Terminate()
    Set fnt = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lblText.Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Font", fnt, Ambient.Font)
End Sub

Private Sub UserControl_Show()
    Width = lblText.Width
    Height = lblText.Height
    lblText.Visible = Not Ambient.UserMode
    
    If Ambient.UserMode Then
        BackStyle = 0
        Cls
        MaskColor = BackColor
        CurrentX = 0: CurrentY = 0
        Set UserControl.Font = lblText.Font
        Print lblText.Caption
        Set MaskPicture = Image
    Else
        BackStyle = 1
        If Caption = vbNullString Then lblText.Caption = "[" & Ambient.DisplayName & "]"
        Width = lblText.Width
        Height = lblText.Height
    End If
End Sub

