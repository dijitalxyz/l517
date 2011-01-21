VERSION 5.00
Begin VB.Form frmWeb 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "L517 - web-based generator"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7605
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
   ScaleHeight     =   2295
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   6480
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame frameURL 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "URL(s) [press 'enter' to add]"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7335
      Begin VB.ComboBox cboURLs 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label lblClear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CLEAR"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6465
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame frameCrawl 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "CRAWL"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   5895
      Begin VB.CheckBox chkDomain 
         BackColor       =   &H00000000&
         Caption         =   "stay within domain"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox chkCrawl 
         BackColor       =   &H00000000&
         Caption         =   "'crawl' through links found on each page"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtDepth 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1920
         TabIndex        =   3
         Text            =   "5"
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblDepth 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH DEPTH:"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   390
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   630
         Width           =   1455
      End
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6120
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 words"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   2040
      Width           =   1485
   End
   Begin VB.Label lblStat 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " inactive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   2040
      Width           =   6195
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public is_running As Boolean, total_count&
Dim surls$()

Private Sub stat(s$)
    lblStat.Caption = " " + s$
End Sub

Private Sub updatecount()
    Dim s$
    s$ = Format(total_count&, "###,###")
    If s$ = "" Then s$ = "0"
    lblCount.Caption = s$ + " words"
End Sub

Private Sub cboURLs_KeyPress(KeyAscii As Integer)
    Dim s$, i&, found As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        s$ = Trim(cboURLs.Text)
        If s$ <> "" Then
            found = False
            For i& = 0 To cboURLs.ListCount - 1
                If cboURLs.List(i&) = s$ Then
                    found = True
                    Exit For
                End If
            Next i&
            
            If found = False Then
                cboURLs.AddItem s$
            End If
        End If
        cboURLs.Text = ""
    End If
End Sub

Private Sub chkCrawl_Click()
    txtDepth.Enabled = CBool(chkCrawl.Value)
    lblDepth.Enabled = CBool(chkCrawl.Value)
    chkDomain.Enabled = CBool(chkCrawl.Value)
    
    regSet "web_crawl", chkCrawl.Value
End Sub

Private Sub chkDomain_Click()
    regSet "web_domain", chkDomain.Value
End Sub

Private Sub Form_Load()
    Dim i&, sarr$()
    
    ReDim surls$(0)
    
    is_running = False
    
    Call setwindowpos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
    
    If regGet("web_crawl") = "1" Then
        chkCrawl.Value = 1
        txtDepth.Enabled = CBool(chkCrawl.Value)
        lblDepth.Enabled = CBool(chkCrawl.Value)
    End If
    
    If regGet("web_domain") = "1" Then
        chkDomain.Value = 1
    End If
    
    txtDepth.Text = regGet("web_depth")
    If txtDepth.Text = "" Or IsNumeric(txtDepth.Text) = False Then
        txtDepth.Text = "5"
    End If
    
    stat "loading urls..."
    sarr$() = Split(regGet("web_urls"), Chr(9))
    For i& = 0 To UBound(sarr$())
        If Trim(sarr$(i&)) <> "" Then
            cboURLs.AddItem sarr$(i&)
        End If
        If i& Mod 500 = 0 Then
            DoEvents
            stat "loaded " + Format(cboURLs.ListCount, "###,###") + " urls"
        End If
        
    Next i&
    If cboURLs.ListCount > 0 Then
        stat "loaded " + Format(cboURLs.ListCount, "###,###") + " urls"
    Else
        stat "inactive"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim s$, i&, max&
    
    max& = 500
    If cboURLs.ListCount - 1 < max& Then
        max& = cboURLs.ListCount - 1
    End If
    
    s$ = ""
    For i& = 0 To max&
        s$ = s$ + cboURLs.List(i&) + Chr(9)
    Next i&
    regSet "web_urls", s$
    
    Unload Me
End Sub

Private Sub lblClear_Click()
    cboURLs.Clear
    stat "urls cleared"
End Sub

Private Sub lblStart_Click()
    lst.Clear
    If lblStart.Caption = "START" Then
        total_count& = 0
        If Trim(cboURLs.Text) <> "" Then
            cboURLs_KeyPress 13
        End If
        frameCrawl.Enabled = False
        frameURL.Enabled = False
        lblStart.Caption = "STOP"
        searchurls
    Else
        frameCrawl.Enabled = True
        frameURL.Enabled = True
        lblStart.Caption = "START"
        stat "inactive"
    End If
    
End Sub

Private Sub searchurls()
    Dim url$, depth%, sttl$, tempsite$, count&, s$
    
    If cboURLs.ListCount = 0 Then
        If lblStart.Caption <> "START" Then
            lblStart_Click
        End If
        Exit Sub
    End If
    
    If lblStart.Caption = "START" Then
        stat "inactive"
        Exit Sub
    End If
    
    url$ = cboURLs.List(0)
    cboURLs.RemoveItem 0
    
    If Left(url$, 1) = "(" Then
        depth% = CInt(Mid(url$, 2, InStr(url$, ")") - 2))
        url$ = Right(url$, Len(url$) - InStr(url$, ")"))
    Else
        depth% = 0
    End If
    
    If Left(url$, 4) <> "http" Then
        url$ = "http://" + url$
    End If
    tempsite$ = url$
    If Len(tempsite$) > 32 Then tempsite$ = Left(tempsite$, 32)
    
    sttl$ = Trim(Format(cboURLs.ListCount, "###,###"))
    If sttl$ = "" Then sttl$ = "0"
    
    stat "" + sttl$ + " remain; downloading " + tempsite$ + "..."
    s$ = webgetsource$(url$)
    stat sttl$ + " remain; parsing " + tempsite$ + "..."
    total_count& = total_count& + ParseWebData&(s$)
    updatecount
    
    If chkCrawl.Value = vbChecked And depth% < CLng(txtDepth.Text) Then
        stat sttl$ + " remain; crawling " + tempsite$ + "..."
        CrawlURLs url$, s$, depth% + 1
    End If
    
    searchurls
End Sub

Private Sub CrawlURLs(url$, source$, depth%)
    Dim i&, j&, s$, domain$, sarr$(), donotadd As Boolean
    
    domain$ = LCase(url$)
    domain$ = Replace(domain$, "http://", "")
    domain$ = Replace(domain$, "https://", "")
    
    j& = InStr(domain$, "/")
    If j& = 0 Then j& = Len(domain$) + 1
    domain$ = Left(domain$, j& - 1)
    
    sarr$() = Split(domain$, ".")
    If UBound(sarr$) = 0 Then Exit Sub
    
    domain$ = sarr$(UBound(sarr$()) - 1) + "." + sarr$(UBound(sarr$()))
    
    i& = 0
    Do
        DoEvents
        i& = InStr(i& + 1, source$, "href=" + Chr(34))
        j& = InStr(i& + Len("href=X") + 1, source$, Chr(34))
        If i& <> 0 And j& <> 0 Then
            i& = i& + Len("href=X")
            s$ = Mid(source$, i&, j& - i&)
            If Left(s$, 1) = "/" Then
                s$ = "http://" + domain$ + s$
            End If
            
            donotadd = False
            
            If Left(s$, 1) = "#" Then
                donotadd = True
            End If
            
            If chkDomain.Value = 1 And InStr(LCase(s$), domain$) = 0 Then
                donotadd = True
            End If
            
            If Left(LCase(s$), 4) <> "http" Then
                donotadd = True
            End If
            
            If InStr(s$, "favicon.ico") <> 0 Or InStr(s$, ".css") <> 0 Or InStr(s$, "imgur.com") <> 0 Then
                donotadd = True
            End If
            
            Select Case Right(LCase(s$), 4)
            Case ".jpg", "jpeg", ".bmp", ".png", ".gif", ".pdf", ".doc", "docx"
                donotadd = True
            End Select
            
            If donotadd = False Then
                AddUrl s$, depth% + 1
            End If
        End If
    Loop Until i& = 0
End Sub

Private Function UrlLookup(url$) As Boolean
    Dim i&
    
    stat "checking for url duplicate..."
    For i& = 0 To UBound(surls$())
        DoEvents
        If InStr(surls$(i&), "," + url$ + ",") <> 0 Then
            stat "duplicate found"
            UrlLookup = True
            Exit Function
        End If
    Next i&
    stat "no dupe found"
    UrlLookup = False
End Function

Private Function AddUrl(url$, depth%) As Boolean
    Dim i&, added As Boolean
    added = False
    
    If UrlLookup(url$) = False Then
        added = True
        i& = UBound(surls$()) - 1
        If i& < 0 Then
            ReDim surls$(0)
            surls$(0) = "," + url$ + ","
        Else
            If Len(surls$(i&)) > 10000 Then
                ReDim Preserve surls$(i& + 1)
                surls$(i& + 1) = "," + url$ + ","
            Else
                surls$(i&) = surls$(i&) + url$ + ","
            End If
        End If
    End If
    
    If added = True And cboURLs.ListCount < 32765 Then
        cboURLs.AddItem "(" + CStr(depth%) + ")" + url$
    End If
    stat "pending urls: " + Format(cboURLs.ListCount, "###,###")
End Function

Private Sub txtDepth_Change()
    regSet "web_depth", txtDepth.Text
End Sub

Private Sub txtDepth_LostFocus()
    If IsNumeric(txtDepth.Text) = False Then
        txtDepth.Text = "5"
    End If
    regSet "web_depth", txtDepth.Text
End Sub
