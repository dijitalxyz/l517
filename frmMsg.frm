VERSION 5.00
Begin VB.Form frmMsg 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "L517"
   ClientHeight    =   270
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   ScaleHeight     =   270
   ScaleWidth      =   2250
   Visible         =   0   'False
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ok
'frmMain was the only form i needed.
'unfortunately, when a form isn't showing in the taskbar,
'   messageboxes and inputboxes don't appear on top of other windows
'this is bad. so i created a new form that DOES appear in the taskbar,
'   but remains invisible when not displaying dialog boxes

'this is so the msgboxes and inputboxes appear ontop of other windows.

Option Explicit

Private Sub Form_Load()
    Me.Left = Screen.Width * 2
    Me.Top = Screen.Height * 2
End Sub

Public Function MsgBocks(sMsg$, Optional sDat As VbMsgBoxStyle = 0, Optional sTitle$ = "L517") As VbMsgBoxResult
    Me.Visible = True
    Call setwindowpos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
    MsgBocks = MsgBox(sMsg$, sDat, sTitle$)
    Call setwindowpos(Me.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
    Call setwindowpos(frmMain.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
    Me.Visible = False
End Function

Public Function InputBocks$(sMsg$, Optional sTitle$ = "L517", Optional sDefault$ = "")
    Me.Visible = True
    Call setwindowpos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
    InputBocks$ = InputBox(sMsg$, sTitle$, sDefault$)
    Call setwindowpos(Me.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
    Call setwindowpos(frmMain.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
    Me.Visible = False
End Function

