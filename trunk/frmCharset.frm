VERSION 5.00
Begin VB.Form frmCharset 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "L517 - EDIT CHARACTER-SETS"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmCharset.frx":0000
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SAVE && CLOSE"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOAD DEFAULT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmCharset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public txt_changed As Boolean

Private Sub Form_Load()
    Dim s$
    Call setwindowpos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
    s$ = regGet("charset")
    txt.Text = s$
    If s$ = "" Then lblDefault_Click
End Sub

Private Sub Form_Resize()
    txt.Width = Me.Width - (txt.Left * 2) - 200
    txt.Height = Me.Height - (txt.Top * 2) - 100
    lblSave.Left = txt.Left + txt.Width - lblSave.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Do While Right(txt.Text, 2) = vbCrLf
        DoEvents
        txt.Text = Left(txt.Text, Len(txt.Text) - 2)
    Loop
    
    regSet "charset", txt.Text
End Sub

Private Sub lblDefault_Click()
    Dim s$
    
    s$ = frmMain.defaultCharset()
    txt.Text = s$
End Sub

Private Sub lblSave_Click()
    Unload Me
End Sub

Private Sub txt_Change()
    txt_changed = True
End Sub
