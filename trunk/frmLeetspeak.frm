VERSION 5.00
Begin VB.Form frmLeetspeak 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "L517 - EDIT LEETSPEAK"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   4005
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
   ScaleHeight     =   6570
   ScaleWidth      =   4005
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
      Height          =   6015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   3735
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
      Left            =   2160
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
Attribute VB_Name = "frmLeetspeak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public txt_changed As Boolean

Private Sub Form_Load()
    Dim s$
    Call setwindowpos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
    s$ = regGet("leetspeak")
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
    
    regSet "leetspeak", txt.Text
End Sub

Private Sub lblDefault_Click()
    Dim s$
    
    s$ = frmMain.defaultLeetSpeak$()
    txt.Text = s$
End Sub

Private Sub lblSave_Click()
    Unload Me
End Sub

Private Sub txt_Change()
    txt_changed = True
End Sub
