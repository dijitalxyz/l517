Attribute VB_Name = "modL517"

Option Explicit

Global SALPH$(25), SCOUNT%(25)

'window declares
Public Declare Function setcapture Lib "user32" Alias "SetCapture" (ByVal hWnd As Long) As Long
Public Declare Function releasecapture Lib "user32" Alias "ReleaseCapture" () As Long
Public Declare Function getcursorpos Lib "user32" Alias "GetCursorPos" (lpPoint As pointapi) As Long
Public Declare Function setforegroundwindow Lib "user32" Alias "SetForegroundWindow" (ByVal hWnd As Long) As Long
Public Declare Function setwindowpos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Type pointapi
    x As Long
    Y As Long
End Type

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const SM_CXVSCROLL = 2
Public Const SM_CYHSCROLL = 3

'end window declares

'directory declares
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'end directory declares

'websource declares
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Const IF_FROM_CACHE = &H1000000
Private Const IF_MAKE_PERSISTENT = &H2000000
Private Const IF_NO_CACHE_WRITE = &H4000000
Private Const BUFFER_LEN = 256
'end web declares

'registry declares
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const ERROR_SUCCESS = 0&
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const REG_SZ = 1 ' Unicode nul terminated String
Private Const REG_DWORD = 4 ' 32-bit number
'end reg declares

'file dialog declares
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Const OFN_FILEMUSTEXIST = &H1000
'end dialog declares

'directory dialog declares
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
    
Private Const MAX_PATH = 260

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_NEWDIALOGSTYLE = &H40

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_VALIDATEFAILED As Long = 3

Private Const WM_USER = &H400

Private Const BFFM_SETSTATUSTEXT As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTION As Long = (WM_USER + 102)
   
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40

Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

'end directory dialog declares

'mp3 info type
Public Type TagInfo
    Tag As String * 3
    Songname As String * 30
    artist As String * 30
    album As String * 30
    year As String * 4
    comment As String * 30
    genre As String * 1
End Type
'end mp3info type

'main sub
Public Sub Main()
    'this code runs first, before loading the main form.
    'the main form requires MSCOMCTL.OCX (for ListView, imageList, and StatusBar)
    'in order to avoid the error, the file is included as a resource (1.0MB)
    'this sub checks if it's installed and, if not, installs it
    
    Dim b_installed As Boolean, s$
    
    'b_installed = ExtractAndRegister("OCX", 101&, "MSCOMCTL.OCX")
    
    If b_installed = False Then
        s$ = "MSCOMCTL.OCX" + vbCrLf
        'MsgBox "the following component was unable to be installed: " + s$ + vbCrLf + "if the program errors on load, try right-clicking L517.exe and selecting 'Run As Administrator'.", vbCritical + vbOKOnly, "L517"
    End If
    
    Load frmMain
End Sub
Private Function ExtractAndRegister(ResType$, ResID&, FileName$) As Boolean
    'try installing file in 3 directories!
    ' - System
    ' - Windows+\SysWOW64 (for 64bit machines)
    ' - App.Path (working directory)
    
    Dim i&, sd$, f$, b_installed!
    
    b_installed = False
    f$ = FileName$
    
    For i& = 1 To 3
        Select Case i&
        Case 1
            sd$ = WindowsDirectory$() + "\SysWOW64\"
        Case 2
            sd$ = SystemDirectory$() + "\"
        Case 3
            sd$ = App.Path + IIf(Right(App.Path, 1) = "\", "", "\")
        End Select
        
        If DirExists(sd$) = True Then
            If Len(Dir(sd$ + f$)) <> 0 Then
                'file already installed
                b_installed = True
                Exit For
            Else
                If ExtractResource(ResType$, ResID&, sd$ + f$) = True Then
                    Shell ("Regsvr32 " & (sd$ + f$) & " /s")
                    b_installed = True
                    Exit For
                End If
            End If
        End If
    Next i&
    ExtractAndRegister = b_installed
End Function
'end main sub

'resource/system functions
Private Function ExtractResource(ResType$, ResID&, OutputPath$) As Boolean
    'returns true if extraction was successful, otherwise false
    On Error GoTo Err_Occurred
    Dim ff%, OCX() As Byte
    
    ExtractResource = False
    
    OCX = LoadResData(ResID&, ResType$)
    ff% = FreeFile
    Open OutputPath$ For Binary As #ff%
        Put #ff%, , OCX()
    Close #ff%
    ExtractResource = True
    
    Exit Function
Err_Occurred:
    Err.Clear
    Close #ff%
    ExtractResource = False
End Function
Private Function SystemDirectory$()
    'returns system directory
    Dim rStr$, rLen&
    
    rStr$ = String(255, 0)
    rLen& = GetSystemDirectory(rStr$, Len(rStr$))
    If rLen& < Len(rStr$) Then
        rStr$ = Left(rStr$, rLen&)
        If Right(rStr$, 1) = "\" Then
            SystemDirectory = Left(rStr$, Len(rStr$) - 1)
        Else
            SystemDirectory = rStr$
        End If
    Else
        SystemDirectory = ""
    End If
End Function
Private Function WindowsDirectory$()
    'returns windows directory [uses environ variables]
    Dim rStr$, rLen&
    
    rStr$ = String(255, 0)
    rLen& = GetWindowsDirectory(rStr$, Len(rStr$))
    If rLen& < Len(rStr$) Then
        rStr$ = Left(rStr$, rLen&)
        If Right(rStr$, 1) = "\" Then
            WindowsDirectory = Left(rStr$, Len(rStr$) - 1)
        Else
            WindowsDirectory = rStr$
        End If
    Else
        WindowsDirectory = ""
    End If
End Function
'end resource/system funcs

'program sub and functions (specific to this program)
Private Function DirExists(sdir$) As Boolean
    'returns true if directory exists, false if not
    If Dir$(sdir$, 16) <> "" Then
        DirExists = True
    Else
        DirExists = False
    End If
End Function
Public Function FilterCheck(ByRef txt$) As Boolean
    'checks if string passes filter, returns true if it does, otherwise false
    'ByRef means the variable being passed is the one that's changed.
    'notice that everytime this function is called, a special 'temporary' variable made
    'when calling this function, EXPECT the passed argument to return the fixed, filtered word!!!
    
    Dim iMin%, imax%, tleft$, tright$
    
    If txt$ = "" Then
        FilterCheck = False
        Exit Function
    End If
    
    iMin% = CInt(frmMain.mnuFilterLenMin.Tag)
    imax% = CInt(frmMain.mnuFilterLenMax.Tag)
    tleft$ = frmMain.mnuFilterTextLeft.Tag
    tright$ = frmMain.mnuFilterTextRight.Tag
    
    If iMin% > 0 And Len(txt$) < iMin% Then
        FilterCheck = False
        Exit Function
    End If
    
    If imax% > 0 And Len(txt$) > imax% Then
        FilterCheck = False
        Exit Function
    End If
    
    FilterCheck = True
    
    With frmMain
        If .mnuCaseUpper.Checked = True Then
            txt$ = UCase(txt$)
        ElseIf .mnuCaseLower.Checked = True Then
            txt$ = LCase(txt$)
        End If
    End With
    
    If InStr(txt$, tleft$) <> 0 And tleft$ <> "" Then
        txt$ = Left(txt$, InStr(txt$, tleft$) - 1)
    End If
    If InStr(txt$, tright$) <> 0 And tright$ <> "" Then
        txt$ = Right(txt$, Len(txt$) - InStr(txt$, tright$) - Len(tright$) + 1)
    End If
    
    If frmMain.mnuFilterHex.Checked = True Then
        txt$ = FixHex$(txt$)
    End If
    
    If IsAlphaNum(txt$) = False Then
        FilterCheck = False
    End If
End Function
Private Function FixHex$(txt$)
    'converts ascii special characters (non-alphanum) to hex
    'includes percent symbol %
    'good for generating passwords to be used in web-based attack
    Dim i%, s$, a%
    
    s$ = ""
    For i% = 1 To Len(txt$)
        a% = Asc(Mid(txt$, i%, 1))
        If a% >= 32 And a% <= 36 Or a% >= 38 And a% <= 47 Or a% >= 58 And a% <= 64 Or _
           a% >= 91 And a% <= 96 Or a% >= 123 And a% <= 126 Then
            s$ = s$ + "%" + CStr(Hex(a%))
        Else
            s$ = s$ + Chr(a%)
        End If
    Next i%
    
    FixHex$ = s$
End Function
Public Sub stat(txt$)
    'updates status bar, blank means 'inactive'
    Dim s$
    
    Select Case frmMain.LANGUAGE$
    Case "english"
        s$ = txt$
        If s$ = "" Then s$ = "inactive"
    Case "french"
        If s$ = "" Then
            s$ = "inactifs"
        Else
            s$ = Replace(s$, "clipboard", "texte copié")
            s$ = Replace(s$, "dragged text", "traîné texte")
            s$ = Replace(s$, "reading", "lecture")
            s$ = Replace(s$, "list", "liste")
            s$ = Replace(s$, "adding", "ajoutant")
            s$ = Replace(s$, "items", "articles")
            s$ = Replace(s$, "postfixes", "postfixes")
            s$ = Replace(s$, "prefixes", "préfixes")
            s$ = Replace(s$, "queueing", "ajoutant")
            s$ = Replace(s$, "saving", "l'épargne")
            s$ = Replace(s$, "saved", "sauvé")
            s$ = Replace(s$, " to ", " à ")
            s$ = Replace(s$, "loading", "chargement")
            s$ = Replace(s$, "updating", "la mise à jour")
            s$ = Replace(s$, "complete", "complète")
            s$ = Replace(s$, "canceled", "annulée")
            s$ = Replace(s$, "removed", "supprimé")
            s$ = Replace(s$, "analyzing", "analyser")
            s$ = Replace(s$, "sorting", "tri")
            s$ = Replace(s$, "generating", "génératrices")
            s$ = Replace(s$, "finalizing", "finalisation")
            s$ = Replace(s$, "subfolders", "dossiers")
            s$ = Replace(s$, "folders", "dossiers")
            s$ = Replace(s$, "parsing", "lecture")
            s$ = Replace(s$, "grabbing", "S'y")
            s$ = Replace(s$, "phone", "Téléphone")
            s$ = Replace(s$, "checking", "vérifier")
            s$ = Replace(s$, "data-write speed", "Vitesse d'écriture data")
            s$ = Replace(s$, "file", "fichier")
            s$ = Replace(s$, "site", "site")
            s$ = Replace(s$, "please wait", "S'il vous plaît patienter")
            s$ = Replace(s$, "clearing list", "Effacer la liste")
            s$ = Replace(s$, "list cleared", "liste autorisé")
            s$ = Replace(s$, "dupekilling", "supprimer les doublons")
            s$ = Replace(s$, "inactive", "inactifs")
            s$ = Replace(s$, "not found", "pas trouver")
            s$ = Replace(s$, "found", "trouvé")
            s$ = Replace(s$, "filtering", "filtrage")
            s$ = Replace(s$, "removing", "enlever")
        End If
    Case "german"
        If s$ = "" Then
            s$ = "inaktiv"
        Else
            s$ = Replace(s$, "clipboard", "kopierten Text")
            s$ = Replace(s$, "dragged text", "Text gezogen")
            s$ = Replace(s$, "reading", "Lesung")
            s$ = Replace(s$, "list", "Liste")
            s$ = Replace(s$, "adding", "Hinzufügen")
            s$ = Replace(s$, "items", "Artikel")
            s$ = Replace(s$, "postfixes", "Postfixen")
            s$ = Replace(s$, "prefixes", "Präfixe")
            s$ = Replace(s$, "queueing", "Hinzufügen")
            s$ = Replace(s$, "saving", "Speichern")
            s$ = Replace(s$, "saved", "gespeichert")
            s$ = Replace(s$, " to ", " nach ")
            s$ = Replace(s$, "loading", "laden")
            s$ = Replace(s$, "updating", "Aktualisierung")
            s$ = Replace(s$, "complete", "abgeschlossen")
            s$ = Replace(s$, "canceled", "abgebrochen")
            s$ = Replace(s$, "removed", "entfernt")
            s$ = Replace(s$, "analyzing", "Analyse")
            s$ = Replace(s$, "sorting", "Sortierung")
            s$ = Replace(s$, "generating", "Erzeugung")
            s$ = Replace(s$, "finalizing", "Abschließen")
            s$ = Replace(s$, "dates", "Termine")
            s$ = Replace(s$, "searching", "Benutzer")
            s$ = Replace(s$, "subfolders", "Unterordner")
            s$ = Replace(s$, "folders", "Ordner")
            s$ = Replace(s$, "parsing", "Parsen")
            s$ = Replace(s$, "grabbing", "Anreise")
            s$ = Replace(s$, "phone", "Telefon")
            s$ = Replace(s$, "checking", "Kontrolle")
            s$ = Replace(s$, "data-write speed", "data Schreibgeschwindigkeit")
            s$ = Replace(s$, "file", "Datei")
            s$ = Replace(s$, "site", "website")
            s$ = Replace(s$, "please wait", "Bitte warten")
            s$ = Replace(s$, "clearing list", "Entfernen aller Elemente")
            s$ = Replace(s$, "list cleared", "Alle Artikel entfernen")
            s$ = Replace(s$, "dupekilling", "Löschen von Duplikaten")
            s$ = Replace(s$, "inactive", "inaktiv")
            s$ = Replace(s$, "not found", "nicht gefunden")
            s$ = Replace(s$, "found", "gefunden")
            s$ = Replace(s$, "filtering", "Filterung")
            s$ = Replace(s$, "removing", "Entfernen")
        End If
    Case "spanish"
        If s$ = "" Then
            s$ = "inactivo"
        Else
            s$ = Replace(s$, "clipboard", "copia de texto")
            s$ = Replace(s$, "dragged text", "arrastrado texto")
            s$ = Replace(s$, "reading", "lectura")
            s$ = Replace(s$, "list", "lista")
            s$ = Replace(s$, "adding", "añadir")
            s$ = Replace(s$, "items", "artículos")
            s$ = Replace(s$, "postfixes", "sufijos")
            s$ = Replace(s$, "prefixes", "prefijos")
            s$ = Replace(s$, "queueing", "añadir")
            s$ = Replace(s$, "saving", "ahorro")
            s$ = Replace(s$, "saved", "salvado")
            s$ = Replace(s$, " to ", " para ")
            s$ = Replace(s$, "loading", "añadir")
            s$ = Replace(s$, "updating", "actualización")
            s$ = Replace(s$, "complete", "completa")
            s$ = Replace(s$, "canceled", "cancelado")
            s$ = Replace(s$, "removed", "eliminado")
            s$ = Replace(s$, "analyzing", "análisis")
            s$ = Replace(s$, "sorting", "Clasificación")
            s$ = Replace(s$, "generating", "generación")
            s$ = Replace(s$, "finalizing", "finalizar")
            s$ = Replace(s$, "dates", "fechas")
            s$ = Replace(s$, "searching", "búsqueda")
            s$ = Replace(s$, "subfolders", "subcarpetas")
            s$ = Replace(s$, "folders", "carpetas")
            s$ = Replace(s$, "parsing", "análisis")
            s$ = Replace(s$, "grabbing", "agarrar")
            s$ = Replace(s$, "phone", "teléfono")
            s$ = Replace(s$, "checking", "control")
            s$ = Replace(s$, "data-write speed", "velocidad de escritura")
            s$ = Replace(s$, "file", "archivo")
            s$ = Replace(s$, "site", "sitio")
            s$ = Replace(s$, "please wait", "por favor, espere")
            s$ = Replace(s$, "clearing list", "lista de facilitación")
            s$ = Replace(s$, "list cleared", "lista despejado")
            s$ = Replace(s$, "dupekilling", "eliminación de duplicados")
            s$ = Replace(s$, "inactive", "inactivo")
            s$ = Replace(s$, "not found", "que no se encuentra")
            s$ = Replace(s$, "found", "encontrado")
            s$ = Replace(s$, "filtering", "Filtro")
            s$ = Replace(s$, "removing", "la eliminación de")
        End If
    End Select
    
    frmMain.lblStat.Caption = " " + s$
End Sub
Private Function DoubleToDouble#(dbln#)
    'converts crazy exponential double values to normal
    'sometimes, when converting double-to-string, we get the "E-02" at the end.
    'this gets rid of that!
    Dim s$, exp%
    
    s$ = CStr(dbln#)
    If InStr(s$, "E") <> 0 Then
        exp% = CInt(Right(s$, 2))
        s$ = Left(s$, 3)
        DoubleToDouble# = CDbl(s$) * (10 ^ -(exp%))
    Else
        DoubleToDouble# = dbln#
    End If
End Function
Public Sub prog(perc#)
    'updates progress bar and progress text in status bar
    Dim s$
    
    If perc# <= 0 Then
        frmMain.lblCancel.Visible = False
        frmMain.picP.Visible = False
        frmMain.lblProg.Caption = ""
        Exit Sub
    End If
    
    frmMain.lblCancel.Visible = True
    
    perc# = DoubleToDouble(perc#)
    
    If perc# > 1 Then perc# = 1
    
    frmMain.picP.Width = frmMain.lst.Width * perc#
    If frmMain.picP.Visible = False Then
        frmMain.picP.Visible = True
    End If
    
    frmMain.lblCancel.Left = frmMain.picP.Left + frmMain.picP.Width
    perc# = perc# * 100
    
    s$ = CStr(perc#)
    If InStr(s$, ".") <> 0 Then
        s$ = Left(s$, InStr(s$, ".") + 1)
    End If
    frmMain.lblProg.Caption = s$ + "%"
    
End Sub
Public Function ParseWebData&(sDat$)
    'used to extract out non-HTML and non-code words from a site
    'sdat is the webpage source
    Dim i&, j&, stxt$, schar$, itemp&, count&
    count& = 0
    
    'remove SCRIPT tags (usually contains code)
    i& = InStr(sDat$, "<script")
    Do While i& > 0
        DoEvents
        j& = InStr(i& + 1, sDat$, "</script>")
        If i& > 0 And j& > 0 Then
            sDat$ = Left(sDat$, i& - 1) + Right(sDat$, Len(sDat$) - j& - Len("</script>") + 1)
        End If
        
        i& = InStr(i&, sDat$, "<script")
    Loop
    
    'remove STYLE tags (more useless code)
    i& = InStr(sDat$, "<style")
    Do While i& > 0
        DoEvents
        j& = InStr(i& + 1, sDat$, "</style>")
        If i& > 0 And j& > 0 Then
            sDat$ = Left(sDat$, i& - 1) + Right(sDat$, Len(sDat$) - j& - Len("</style>") + 1)
        End If
        
        i& = InStr(i&, sDat$, "<style")
    Loop
    
    stxt$ = ""
    For i& = 1 To Len(sDat$)
        schar$ = Mid(sDat$, i&, 1)
        Select Case schar$
        Case "<"
            'skip html tags ( dangerous if source non-tags contain < )
            itemp& = InStr(i&, sDat$, ">")
            If itemp& = 0 Then Exit For
            'If itemp& < 350 Then 'itemp& < InStr(i& + 1, sdat$, vbCrLf)
                i& = itemp&
            'End If
        Case Else
            If IsAlphaNum(schar$) = True Then
                stxt$ = stxt$ + schar$
            ElseIf stxt$ <> "" Then
                If FilterCheck(stxt$) = True Then
                    frmMain.lst.ListItems().add , , stxt$
                    count& = count& + 1
                End If
                stxt$ = ""
            End If
        End Select
    Next i&
    
    If stxt$ <> "" And FilterCheck(stxt$) = True Then
        frmMain.lst.ListItems().add , , stxt$
        count& = count& + 1
    End If
    
    UpdateCaption
    ParseWebData& = count&
End Function
Public Function CaseFirst$(txt$)
    'sets first letter of string [and every letter after space] to upper case, the rest lower
    Dim i&, ch$, bnext As Boolean, sall$
    bnext = True
    sall$ = ""
    For i& = 1 To Len(txt$)
        ch$ = Mid$(txt$, i&, 1)
        sall$ = sall$ + IIf(bnext, UCase(ch$), LCase(ch$))
        If ch$ = " " Then
            bnext = True
        Else
            bnext = False
        End If
    Next i&
    CaseFirst$ = sall$
End Function
Public Function CaseLeet0$(sword$)
    Dim lword%(), i%, s$
    
    ReDim lword%(Len(sword$))
    For i% = 0 To UBound(lword%())
        lword%(i%) = 0
    Next i%
    lword%(0) = 1
    
    Do
        DoEvents
        s$ = ""
        For i% = 0 To UBound(lword%())
            s$ = s$ + LetterToLeet$(Mid(sword$, i% + 1, 1), lword%(i%))
        Next i%
        
    Loop Until i% = -1
End Function
Public Function CaseLeet$(sword$)
    'new method of converting to leetspeak
    'will generate every possible mutation of the word (DOESN'T)
    Dim s$, i%, j%, lword%(), sreturn$, a%
    
    'setup intiial conditions
    ReDim lword%(Len(sword$))
    For i% = 0 To UBound(lword%())
        lword%(i%) = 0
    Next i%
    lword%(0) = 1
    
    Do
        DoEvents
        If frmMain.lblCancel.Visible = False Then
            CaseLeet$ = sreturn$
            Exit Function
        End If
        s$ = ""
        For i% = 0 To UBound(lword%())
            s$ = s$ + LetterToLeet$(Mid(sword$, i% + 1, 1), lword%(i%))
        Next i%
        sreturn$ = sreturn$ + s$ + IIf(frmMain.mnuFileTypeUnix.Checked, Chr(10), vbCrLf)
        For i% = 0 To UBound(lword%()) - 1
            a% = Asc(LCase(Mid(sword$, i% + 1, 1)))
            If a% >= Asc("a") And a% <= Asc("z") Then
                If lword%(i%) < SCOUNT%(a% - 97) - 1 Then
                    lword%(i%) = lword%(i%) + 1
                    For j% = 0 To i% - 1
                        lword%(j%) = 0
                    Next j%
                    Exit For
                Else
                    If i% = UBound(lword%) - 1 Then
                        i% = -1
                        Exit For
                    End If
                End If
            Else
                If i% = UBound(lword%) - 1 Then
                    i% = -1
                    Exit For
                End If
            End If
        Next i%
    Loop Until i% = -1
    
    If Right(sreturn$, 2) = vbCrLf Then sreturn$ = Left(sreturn$, Len(sreturn$) - 2)
    If Right(sreturn$, 1) = Chr(10) Then sreturn$ = Left(sreturn$, Len(sreturn$) - 1)
    
    CaseLeet$ = sreturn$
End Function
Private Function LetterToLeet$(letter$, Index%)
    Dim sa$()
    
    If letter$ = "" Then
        LetterToLeet$ = ""
        Exit Function
    End If
    
    If Asc(LCase(letter$)) >= Asc("a") And Asc(LCase(letter$)) <= Asc("z") Then
        sa$() = Split(SALPH$(Asc(LCase(letter$)) - 97), ",")
        If UBound(sa$()) < Index% Then
            LetterToLeet$ = letter$
        Else
            LetterToLeet$ = sa$(Index%)
        End If
    Else
        LetterToLeet$ = letter$
    End If
    
    Erase sa$()
End Function
Public Function CountString&(sBig$, SCOUNT$)
    Dim c&, i&
    
    c& = 0
    i& = 0
    Do
        DoEvents
        i& = InStr(i& + 1, sBig$, SCOUNT$)
        If i& <> 0 Then c& = c& + 1
    Loop Until i& = 0
    CountString& = c&
End Function
Public Function CaseLeet2$(txt$)
    'old method of converting to 'leet case'
    Dim i&, ch$, sall$
    
    sall$ = ""
    
    For i& = 1 To Len(txt$)
        ch$ = Mid(txt$, i&, 1)
        Select Case LCase(ch$)
        Case "a"
            sall$ = sall$ + "4"
        Case "e"
            sall$ = sall$ + "3"
        Case "g"
            sall$ = sall$ + "9"
        Case "h"
            sall$ = sall$ + "#"
        Case "i"
            sall$ = sall$ + "1"
        Case "l"
            sall$ = sall$ + "1"
        Case "o"
            sall$ = sall$ + "0"
        Case "s"
            sall$ = sall$ + "5"
        Case "t"
            sall$ = sall$ + "7"
        Case "z"
            sall$ = sall$ + "2"
        Case Else
            sall$ = sall$ + ch$
        End Select
    Next i&
    CaseLeet2$ = sall$
End Function
Public Function CaseEveryOther$(txt$)
    'changes case of every other letter from upper, lower, upper, etc
    Dim i&, sall$
    
    sall$ = ""
    For i& = 1 To Len(txt$)
        If i Mod 2 = 0 Then
            sall$ = sall$ + LCase(Mid(txt$, i&, 1))
        Else
            sall$ = sall$ + UCase(Mid(txt$, i&, 1))
        End If
    Next i&
    
    CaseEveryOther$ = sall$
End Function
Private Function IsAlphaNum(stxt$) As Boolean
    'checks if character is made up of only letters and numbers
    'returns true if it is, false if it is not alphanumeric
    Dim x%
    
    If stxt$ = "" Then
        IsAlphaNum = False
        Exit Function
    End If
    
    IsAlphaNum = False
    x% = Asc(stxt$) 'Asc(Mid(stxt$, 1, 1))
    
    If x% >= Asc("a") And x% <= Asc("z") Or x% >= Asc("A") And x% <= Asc("Z") Or x% >= Asc("0") And x% <= Asc("9") Then
        IsAlphaNum = True
    Else
        If frmMain.mnuFilterForeign.Checked = True Then
            
            If x% >= 192 And x% <= 255 Then
                IsAlphaNum = True
            Else
                Select Case x%
                Case 161, 164, 158, 159, 154, 131, 142, 138, 128, 163, 165
                    IsAlphaNum = True
                End Select
            End If
        End If
    End If
End Function
Public Sub ParseTextBlock(stxt$)
    'parses a block of text for words, adds words to list as it finds them.
    'good for small chunks of text
    Dim i&, sword$, schar$
    
    For i& = 1 To Len(stxt$)
        schar$ = Mid(stxt$, i&, 1)
        If IsAlphaNum(schar$) = True Then
            sword$ = sword$ + schar$
        Else
            sword$ = Trim(sword$)
            If sword$ <> "" And FilterCheck(sword$) = True Then
                frmMain.lst.ListItems().add , , sword$
            End If
            sword$ = ""
        End If
    Next i&
    
    sword$ = Trim(sword$)
    If sword$ <> "" And FilterCheck(sword$) = True Then
        frmMain.lst.ListItems().add , , sword$
    End If
End Sub
Public Sub SaveList(sfile$)
    'saves frmMain.lst (main program list) to filename [sfile$]
    'handles split files and all the different cases.
    
    Dim ff%, i&, s$, b_first!, b_leet!, b_everyother!, lcount#, isplit&, icount&, stemp$
    Dim lastitem$, skipped$
    
    icount& = 1
    isplit& = 0
    s$ = regGet("split")
    If IsNumeric(s$) = True Then isplit& = CLng(s$)
    
    stat "saving '" + GetFileName(sfile$) + "'"
    
    b_first! = frmMain.mnuCaseFirst.Checked
    b_leet! = frmMain.mnuCaseLeet.Checked
    b_everyother! = frmMain.mnuCaseEveryother.Checked
    
    lcount# = frmMain.lst.ListItems().count
    
    prog 0.001
    ff% = FreeFile
    Open sfile$ For Binary Access Write As #ff%
        For i& = 1 To lcount#
            If i& Mod 250 = 0 Then
                DoEvents
                If frmMain.lblCancel.Visible = False Then
                    prog 0
                    stat "saving canceled"
                    Exit Sub
                End If
                prog i& / lcount#
            End If
            
            If lastitem$ <> frmMain.lst.ListItems(i&).Text Then
                If isplit& > 0 Then
                    If icount& >= isplit& Then
                        DoEvents
                        icount& = 0
                        'close old file
                        Close #ff%
                        'find next file name in sequence
                        sfile$ = NextFile$(sfile$)
                        'open new file
                        ff% = FreeFile
                        Open sfile$ For Binary Access Write As #ff%
                        DoEvents
                    End If
                End If
                
                'save the item
                Put #ff%, , CStr(frmMain.lst.ListItems(i&).Text + IIf(frmMain.mnuFileTypeWin.Checked, vbCrLf, Chr(10)))
                
                icount& = icount& + 1
                
                s$ = frmMain.lst.ListItems(i&).Text
                If b_first = True And s$ <> CaseFirst$(s$) Then
                    'if first letter uppercase is true and it's not gonna be a dupe..
                    Put #ff%, , CStr(CaseFirst$(frmMain.lst.ListItems(i&).Text) + IIf(frmMain.mnuFileTypeWin.Checked, vbCrLf, Chr(10)))
                    icount& = icount& + 1
                End If
                
                If b_everyother = True And s$ <> CaseEveryOther(s$) Then
                    If b_first = True And CaseFirst(s$) <> CaseEveryOther$(s$) Or b_first = False Then
                        'if every other letter upper is true and it's nto gonna be a dupe..
                        Put #ff%, , CStr(CaseEveryOther$(frmMain.lst.ListItems(i&).Text) + IIf(frmMain.mnuFileTypeWin.Checked, vbCrLf, Chr(10)))
                        icount& = icount& + 1
                    End If
                End If
                
                If b_leet = True Then
                    stemp$ = CaseLeet$(frmMain.lst.ListItems(i&).Text)
                    If b_first = True And CaseFirst(s$) <> stemp$ Or b_first = False Then
                        If b_everyother = True And CaseEveryOther(s$) <> stemp$ Or b_everyother = False Then
                            If s$ <> stemp$ Then
                                Put #ff%, , CStr(CStr(stemp$) + IIf(frmMain.mnuFileTypeWin.Checked, vbCrLf, Chr(10)))
                                icount& = icount& + 1
                            End If
                        End If
                    End If
                End If
                lastitem$ = frmMain.lst.ListItems(i&).Text
            End If
        Next i&
    Close #ff%
    
    skipped$ = Format(lcount# - icount&, "###,###")
    If skipped$ = "" Then skipped$ = "0"
    stat "saved '" + GetFileName(sfile$) + "'; " + skipped$ + " skipped"
    
    prog 0
End Sub
Public Function NextFile$(s$)
    'give it a filename, it will increment the filename by one.
    'give it c:\asdf.txt,   it will return c:\asdf1.txt
    'give it c:\asdf32.txt, it will return c:\asdf33.txt
    'useful when splitting files, helps keep the naming conventions
    Dim ext$, i&, sf$, stemp$, sfall$, b As Boolean
    
    ext$ = Right(s$, 4)
    If InStr(ext$, ".") <> 0 Then
        sf$ = Left(s$, InStrRev(s$, ".") - 1)
        ext$ = Right(ext$, Len(ext$) - InStr(ext$, ".") + 1)
        sfall$ = ""
        b = False
        For i& = Len(sf$) To 1 Step -1
            If IsNumeric(Mid(sf$, i&, 1)) = False Or b = True Then
                sfall$ = Mid(sf$, i&, 1) + sfall$
                b = True
            End If
        Next i&
        sf$ = sfall$
        i& = 0
        Do
            DoEvents
            i& = i& + 1
            stemp$ = sf$ + CStr(i&) + ext$
        Loop Until Len(Dir(stemp$)) = 0
    Else
        sfall$ = ""
        b = False
        For i& = Len(s$) To 1 Step -1
            If IsNumeric(Mid(s$, i&, 1)) = False Or b = True Then
                sfall$ = Mid(s$, i&, 1) + sfall$
                b = True
            End If
        Next i&
        s$ = sfall$
        
        i& = 0
        Do
            DoEvents
            i& = i& + 1
            stemp$ = s$ + CStr(i&)
        Loop Until Len(Dir(stemp$)) = 0
    End If
    
    NextFile$ = stemp$
End Function
Private Function IsWindowsFile(sfile$) As Boolean
    'checks if file contains vbCrLf (chr(13) + chr(10))
    'if only chr(10) is given for line breaks, it's a unix file
    Dim ff%, s$, flen&
    
    ff% = FreeFile
    Open sfile$ For Binary Access Read As #ff%
        'open file for binary
        flen& = FileLen(sfile$)
        If flen& > 256 Then flen& = 128
        'maximum length of chunk we're getting is 128 bytes
        s$ = Space(256)
        
        Get #ff%, 256, s$
        
        If InStr(s$, Chr(13) + Chr(10)) <> 0 Then 'contains chr(13) + chr(10) [aka carriage-return + linefeed]
            IsWindowsFile = True
            
        ElseIf InStr(s$, Chr(10)) <> 0 And InStr(s$, Chr(13) + Chr(10)) = 0 Then 'just line-feed, no carriage return
            IsWindowsFile = False
        Else 'can't find chr(10) by itself OR chr(13) + chr(10), default to unix to be safe.
            IsWindowsFile = False
        End If
    Close #ff%
End Function
Public Function BytesToString$(ByVal num#)
    'converts number of bytes in a file to a string
    'instead of 235809725 bytes, it will return 2.53 gigabytes
    ' ... or whatever that equals
    Dim eb#, pb#, tb#, gb#, mb#, kb#, s$
    kb# = 1024
    If num# < kb# Then
        s$ = CStr(num#) + " bytes"
        If InStr(s$, ".") <> 0 Then s$ = Left(s$, InStr(s$, ".") + 2)
        BytesToString$ = s$
        Exit Function
    End If
    
    mb# = kb# * kb#
    If num# < mb# Then
        num# = num# / kb#
        s$ = CStr(num#)
        If InStr(s$, ".") <> 0 Then s$ = Left(s$, InStr(s$, ".") + 2)
        BytesToString$ = s$ + " kilobytes"
        Exit Function
    End If
    
    gb# = mb# * kb#
    If num# < gb# Then
        num# = num# / mb#
        s$ = CStr(num#)
        If InStr(s$, ".") <> 0 Then s$ = Left(s$, InStr(s$, ".") + 2)
        BytesToString$ = s$ + " megabytes"
        Exit Function
    End If
    
    tb# = gb# * kb#
    If num# < tb# Then
        num# = num# / gb#
        s$ = CStr(num#)
        If InStr(s$, ".") <> 0 Then s$ = Left(s$, InStr(s$, ".") + 2)
        BytesToString$ = s$ + " gigabytes"
        Exit Function
    End If
    
    pb# = tb# * kb#
    If num# < pb# Then
        num# = num# / tb#
        s$ = CStr(num#)
        If InStr(s$, ".") <> 0 Then s$ = Left(s$, InStr(s$, ".") + 2)
        BytesToString$ = s$ + " terabytes"
        Exit Function
    End If
    
    eb# = pb# * kb#
    If num# < eb# Then
        num# = num# / pb#
        s$ = CStr(num#)
        If InStr(s$, ".") <> 0 Then s$ = Left(s$, InStr(s$, ".") + 2)
        BytesToString$ = s$ + " exabytes"
        Exit Function
    End If
    
    num# = num# / (eb#)
    s$ = CStr(num#)
    If InStr(s$, ".") <> 0 Then s$ = Left(s$, InStr(s$, ".") + 2)
    BytesToString$ = s$ + " zettabytes"
End Function
Public Function CalcETA$(size#, lag#)
    'returns the guesstimated time it will take to save a file
    'size is the filesize we are writing
    'lag is the time it takes for this computer to write 1MB of data
    Dim megs#, sec#
    
    megs# = size# / 1048576
    sec# = megs# * lag#
    sec# = sec# * 1.5
    
    CalcETA$ = SecToString$(sec#)
End Function
Private Function SecToString$(sec#)
    'converts number of seconds to a string
    '3665 seconds would return:
    '1 hours, 1 minutes
    Dim s%, m%, h%, d%, w%, mon%, x As Double
    If sec# < 60 Then
        s% = CInt(sec#)
        SecToString$ = CStr(s%) + " seconds"
    ElseIf sec# < 3600 Then
        m% = CInt(sec# / 60)
        s% = sec# Mod 60 ' Abs(sec# - (m% * 60))
        SecToString$ = CStr(m%) + " minutes, " + CStr(s%) + " seconds"
    ElseIf sec# < 86400 Then
        h% = CInt(sec# / 3600)
        m% = sec Mod 3600 ' Abs(sec# - (h% * 3600) / 60)
        SecToString$ = CStr(h%) + " hours, " + CStr(m%) + " minutes"
    ElseIf sec# < 604800 Then
        d% = CInt(sec# / 86400)
        h% = CInt(CDbl(sec# Mod 86400) / 3600)
        SecToString$ = CStr(d%) + " days, " + CStr(h%) + " hours"
    ElseIf sec# < 18144000 Then
        mon% = CInt(sec# / 604800)
        d% = CInt(CDbl(sec# Mod 604800) / 86400)
        SecToString$ = CStr(mon%) + " months, " + CStr(d%) + " days"
    Else
        SecToString$ = "too large to calculate."
    End If
End Function
Public Sub Pause(interval#)
    'waits for defined interval
    'uses do-loop with doevents
    Dim timah#
    
    timah# = Timer
    Do While Timer - timah# < interval#
        DoEvents
    Loop
End Sub
Public Function LoadList(sfile$, Optional ShowListAtEnd As Boolean = True) As Boolean
    'sfile$ is the path to the file with words in it
    'showlistatend decides if the list is set to visible after running (good when loading multiple files)
    'LoadList() returns true if it loads the entire list, false if it's canceled during load
    
    Dim s$, i&, sprog$, bthere As Boolean, m As TagInfo, count&, ff%, stemp$, char$
    Dim current&, j&, k&, chunk%, bdone As Boolean, ass%, spath$, timah#, filec&
    
    If Len(Dir(sfile$)) = 0 Then Exit Function
    
    LoadList = True
    
    Select Case Right(LCase(sfile$), 4)
    Case ".doc", "docx" '".doc", "docx"
        'msword doc files
        
        'make sure MSWORD is installed...
        sprog$ = Environ("ProgramFiles")
        If Right(sprog$, 1) <> "\" Then sprog$ = sprog$ + "\"
        bthere = False
        For i& = 12 To 8 Step -1
            If Len(Dir(sprog$ + "\Microsoft Office\Office" + CStr(i&) + "\MSWORD.OLB")) <> 0 Then
                bthere = True
                Exit For
            End If
        Next i&
        
        'if word IS installed
        If bthere = True Then
            Dim werd As Word.Application
            Dim werdcon As Word.Range
            
            Set werd = New Word.Application
            werd.Documents.Open sfile$, , True, False
            
            Set werdcon = werd.ActiveDocument.Content
            
            ParseTextBlock CStr(werdcon)
            DoEvents
        End If
    Case ".mp3"
        'mp3
        'strip artist/title/tag information
        With m
            Open sfile$ For Binary Access Read As #1
                Get #1, FileLen(sfile$) - 127, .Tag
                If Not .Tag = "TAG" Then
                    'no tag
                    Close #1
                    Exit Function
                End If
                Get #1, , .Songname
                Get #1, , .artist
                Get #1, , .album
                Get #1, , .year
                Get #1, , .comment
                Get #1, , .genre
            Close #1
            ParseTextBlock RTrim(.Tag) + " " + RTrim(.Songname) + " " + RTrim(.artist) + " " + RTrim(.album) + " " + RTrim(.year) + " " + Trim(.comment) + " " + RTrim(.genre)
        End With
    Case ".jpg", ".jpeg"
        'PHOTOS
        'strip tag information
        s$ = Space(FileLen(sfile$))
        Open sfile$ For Binary Access Read As #1
            Get #1, , s$
        Close #1
        i& = InStr(s$, "<x:xmpmeta")
        If i& <> 0 Then
            s$ = Mid(s$, i&, InStr(i& + 1, s$, "</x:xmpmeta"))
        End If
        ParseWebData s$
    Case ".pdf"
        'PDF document
        'need to use pdftotext.exe to convert, dump to a temp file, read the file for words [and then delete it]
        spath$ = App.Path + IIf(Right(App.Path, 1) = "\", "", "\")
        If ExtractResource("EXE", 101, spath$ + "pdftotext.exe") = True Then
            Do While Len(Dir(spath$ + "temp.txt")) <> 0
                Kill spath$ + "temp.txt"
                DoEvents
            Loop
            
            Shell spath$ + "pdftotext.exe " + Chr(34) + sfile$ + Chr(34) + " " + Chr(34) + spath$ + "temp.txt" + Chr(34), vbHide
            DoEvents
            sfile$ = spath$ + "temp.txt"
            
            timah# = Timer
            Do While Timer - timah# < 3 And Len(Dir(sfile$)) = 0
                DoEvents
            Loop
            DoEvents
            If Len(Dir(sfile$)) = 0 Then
                'pdf wasn't decoded correctly, must be a corrupt file.
                LoadList = True
                Kill spath$ + "pdftotext.exe"
                Exit Function
            End If
            
            timah# = Timer
            Do While FileLen(sfile$) = 0 And Timer - timah# < 2
                DoEvents
            Loop
            
            filec& = -1
            Do While Timer - timah# < 3 And filec& <> FileLen(sfile$)
                filec = FileLen(sfile$)
                Pause 0.3
            Loop
            
            count& = 0
            i& = 1
            
            stat "initializing..."
            DoEvents
            
            ff% = FreeFile
            Open sfile$ For Input As #ff%
                While Not EOF(ff%)
                    If count& Mod 500 = 0 Then DoEvents
                    Input #ff%, s$
                    count& = count& + 1
                Wend
            Close #ff%
            stat "loading '" + GetFileName$(sfile$) + "'"
            DoEvents
            
            prog 0.01
            frmMain.lst.Visible = False
            
            ff% = FreeFile
            Open sfile$ For Input As #ff%
                While Not EOF(ff%)
                    If i& Mod 250 = 0 Then
                        DoEvents
                        
                        If frmMain.lblCancel.Visible = False Then
                            'cancel was clicked
                            prog 0
                            frmMain.lst.Visible = True
                            stat "loading cancelled"
                            UpdateCaption
                            LoadList = False
                            Exit Function
                        End If
                        
                        UpdateCaption
                        prog i& / FileLen(sfile$)
                    End If
                    
                    Input #ff%, s$
                    
                    i& = i& + Len(s$) + 2
                    
                    If InStr(s$, " ") = 0 And InStr(s$, Chr(0)) = 0 Then
                        If FilterCheck(s$) = True Then
                            frmMain.lst.ListItems().add , , s$
                        End If
                    Else
                        ParseTextBlock s$
                    End If
                Wend
            Close #ff%
            
            prog 0
            stat ""
            frmMain.lst.Visible = True
            
            On Error Resume Next
            'remove _temp.txt file
            Do While Len(Dir(sfile$)) <> 0
                Kill sfile$
                DoEvents
            Loop
            
            'remove pdftotext.exe file
            Do While Len(Dir(spath$ + "pdftotext.exe")) <> 0
                Kill spath$ + "pdftotext.exe"
                DoEvents
            Loop
        End If
        
        LoadList = True
    'Case ".rtf", ".doc", "docx"
    '    'rich text file, can also handle word documents
    '    On Error Resume Next 'just in case
    '    frmMain.rtf1.LoadFile sfile$
    '    ParseTextBlock frmMain.rtf1.Text
    '    On Error GoTo 0
    Case ".avi", ".mpg", "mpeg", ".mov", ".mkv", "divx", ".mp4", ".exe", "ipsw", _
         ".msi", ".zip", ".rar", "r.gz", ".tar", "rmvb", ".flv", "rent", ".dat", _
         ".sub", ".mdf", ".iso", ".uif", ".dll", _
         ".r00", ".r01", ".r02", ".r03", ".r04", ".r05", ".r06", ".r07", ".r08", ".r09", ".r10", ".r11", ".r12", ".r13", ".r14", ".r15", ".r16", ".r17", ".r18", ".r19", ".r20", ".r21", ".r22", ".r23", ".r24", ".r25", ".r26", ".r27", ".r28", ".r29", ".r30", ".r31", ".r32", ".r33", ".r34", ".r35", ".r36", ".r37", ".r38", ".r39", ".r40", ".r41", ".r42", ".r43", ".r44", ".r45", ".r46", ".r47", ".r48", ".r49", ".r50", ".r51", ".r52", ".r53", ".r54", ".r55", ".r56", ".r57", ".r58", ".r59", ".r60", ".r61", ".r62", ".r63", ".r64", ".r65", ".r66", ".r67", ".r68", ".r69", ".r70", ".r71", ".r72", ".r73", ".r74", ".r75", ".r76", ".r77", ".r78", ".r79", ".r80", ".r81", ".r82", ".r83", ".r84", ".r85", ".r86", ".r87", ".r88", ".r89", ".r90", ".r91", ".r92", ".r93", ".r94", ".r95", ".r96", ".r97", ".r98", ".r99"
         
        'binary, needs special rules!
        'load as unix file...
        chunk% = 512
        count& = FileLen(sfile$)
        If count& = 0 Then
            LoadList = True
            Exit Function
        End If
        
        If count& < chunk% Then
            chunk% = count&
        End If
        
        ReDim bbytes(chunk%) As Byte
        current& = 1
        
        prog 0.001
        
        frmMain.lst.Visible = False
        
        stat "loading unix file '" + GetFileName$(sfile$) + "'"
        DoEvents
        
        bdone = False
        ff% = FreeFile
        Open sfile$ For Binary As #ff%
            Do While current& <> count&
                If (current& / 512) Mod 250 Then
                    DoEvents
                    If frmMain.lblCancel.Visible = False Then
                        'cancel was clicked
                        prog 0
                        frmMain.lst.Visible = True
                        stat "loading cancelled"
                        UpdateCaption
                        LoadList = False
                        Exit Function
                    End If
                    
                    UpdateCaption
                    prog current& / count&
                End If
                
                If current& + chunk% > count& Then
                    chunk% = count& - current&
                    ReDim bbytes(chunk%) As Byte
                End If
                
                Get #ff, current&, bbytes
                
                current& = current& + chunk%
                
                For i& = 1 To chunk%
                    char$ = Chr(Format$(bbytes(i&)))
                    'ass% = Asc%(char$)
                    'If ass% >= 33 And ass% <= 126 Then
                    If IsAlphaNum(char$) = True Then
                        stemp$ = stemp$ + char$
                    Else
                        If stemp$ <> "" And FilterCheck(stemp$) = True Then
                            frmMain.lst.ListItems().add , , stemp$
                        End If
                        stemp$ = ""
                    End If
                Next i&
            Loop
            
            If stemp$ <> "" And FilterCheck(stemp$) = True Then
                frmMain.lst.ListItems().add , , stemp$
            End If
        Close #ff%
        
        Erase bbytes()
        
        LoadList = True
        
        stat ""
        prog 0
        
        DoEvents
        
        If ShowListAtEnd = True Then
            frmMain.lst.Visible = True
        End If
    Case Else
        '.txt, .lst, .ppt, other
        
        frmMain.B_CHANGE = True
        If IsWindowsFile(sfile$) = True Then
            'windows file, uses vbCrLf (carriage return + line-feed for line breaks)
            count& = 0
            i& = 0
            
            stat "initializing..."
            DoEvents
            
            ff% = FreeFile
            Open sfile$ For Input As #ff%
                While Not EOF(ff%)
                    If count& Mod 500 = 0 Then DoEvents
                    Input #ff%, s$
                    count& = count& + 1
                Wend
            Close #ff%
            
            stat "loading '" + GetFileName$(sfile$) + "'"
            DoEvents
            
            prog 0.01
            frmMain.lst.Visible = False
            
            ff% = FreeFile
            Open sfile$ For Input As #ff%
                While Not EOF(ff%)
                    i& = i& + 1
                    If i& Mod 250 = 0 Then
                        DoEvents
                        
                        If frmMain.lblCancel.Visible = False Then
                            'cancel was clicked
                            prog 0
                            frmMain.lst.Visible = True
                            stat "loading cancelled"
                            UpdateCaption
                            LoadList = False
                            Exit Function
                        End If
                        
                        UpdateCaption
                        prog i& / count&
                    End If
                    
                    Input #ff%, s$
                    If InStr(s$, " ") = 0 And InStr(s$, Chr(0)) = 0 Then
                        If FilterCheck(s$) = True Then
                            frmMain.lst.ListItems().add , , CStr(s$)
                        End If
                    Else
                        ParseTextBlock s$
                    End If
                Wend
            Close #ff%
            'end of windows-file-handler
        Else
            'unix file, only uses chr(10) (carriage return) for line breaks
            
            chunk% = 512
            count& = FileLen(sfile$)
            If count& = 0 Then
                LoadList = True
                Exit Function
            End If
            
            If count& < chunk% Then
                chunk% = count&
            End If
            
            ReDim bbytes(chunk%) As Byte
            current& = 1
            
            prog 0.001
            
            frmMain.lst.Visible = False
            
            stat "loading unix file '" + GetFileName$(sfile$) + "'"
            DoEvents
            
            bdone = False
            ff% = FreeFile
            Open sfile$ For Binary As #ff%
                Do While current& <> count&
                    If (current& / 512) Mod 250 Then
                        DoEvents
                        If frmMain.lblCancel.Visible = False Then
                            'cancel was clicked
                            prog 0
                            frmMain.lst.Visible = True
                            stat "loading cancelled"
                            UpdateCaption
                            LoadList = False
                            Exit Function
                        End If
                        
                        UpdateCaption
                        prog current& / count&
                    End If
                    
                    If current& + chunk% > count& Then
                        chunk% = count& - current&
                        ReDim bbytes(chunk%) As Byte
                    End If
                    
                    Get #ff, current&, bbytes
                    
                    current& = current& + chunk%
                    
                    For i& = 1 To chunk%
                        char$ = Chr(Format$(bbytes(i&)))
                        If char$ <> Chr(10) And char$ <> " " Then
                            stemp$ = stemp$ + char$
                        Else
                            If stemp$ <> "" And FilterCheck(stemp$) = True Then
                                frmMain.lst.ListItems().add , , stemp$
                            End If
                            stemp$ = ""
                        End If
                    Next i&
                Loop
                
                If stemp$ <> "" And FilterCheck(stemp$) = True Then
                    frmMain.lst.ListItems().add , , stemp$
                End If
            Close #ff%
            
            Erase bbytes()
        End If
        
        
        LoadList = True
        
        stat ""
        prog 0
        
        DoEvents
        
        If ShowListAtEnd = True Then
            frmMain.lst.Visible = True
        End If
    End Select
    
    UpdateCaption
End Function
Public Sub UpdateCaption()
    'updates the programs titlebar with the # of items in the list
    'only puts L517 if the list is empty
    Dim s$
    s$ = Format(frmMain.lst.ListItems().count, "###,###")
    If s$ <> "" Then
        s$ = " : " + s$ + ""
    End If
    frmMain.lblTitle.Caption = "L517" + s$
End Sub
Public Function getdirectory$(sfile$)
    'receives path to a file.
    'returns everything except the filename [full path to file]
    Dim i&, stemp$
    
    stemp$ = sfile$
    i& = InStrRev(stemp$, "\")
    If i& <> 0 Then
        stemp$ = Left(stemp$, i&)
    End If
    
    getdirectory$ = stemp$
End Function
Public Function GetFileName$(sf$)
    'opposite of getdirectory$()
    'receives path to file, returns only the filename [no directories]
    
    Dim sarr$()
    
    If sf$ = "" Or InStr(sf$, "\") = 0 Then
        GetFileName$ = ""
        Exit Function
    End If
    
    sarr$() = Split(sf$, "\")
    GetFileName$ = sarr$(UBound(sarr$()))
    
    Erase sarr$()
End Function
'end program specific subs/funcs

'begin registry functions
Public Function regGet(sSetting As String) As String
    'reads value from registry, specific to this program
    regGet$ = getstring(HKEY_CURRENT_USER&, "Software\VB and VBA Program Settings\L517", sSetting$)
End Function
Public Sub regSet(sSetting As String, sVal As String)
    'sets value to registry, specific to this program
    Call savestring(HKEY_CURRENT_USER&, "Software\VB and VBA Program Settings\L517", sSetting$, sVal$)
End Sub
Public Function getstring(Hkey As Long, strPath As String, strValue As String)
    'retrieves value from registry
    On Error Resume Next
    Dim keyhand&, datatype&, lresult&, strBuf$, lDataBufSize&, intZeroPos%, lValueType&, r&
    r = RegOpenKey(Hkey, strPath, keyhand)
    lresult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lresult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        
        If lresult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            
            If intZeroPos > 0 Then
                getstring = Left$(strBuf, intZeroPos - 1)
            Else
                getstring = strBuf
            End If
        End If
    End If
    On Error GoTo 0
End Function

Public Sub savestring(Hkey&, strPath$, strValue$, strData$)
    'saves value to registry
    Dim keyhand&, r&
    r = RegCreateKey(Hkey, strPath$, keyhand)
    r = RegSetValueEx(keyhand, strValue$, 0, REG_SZ, ByVal strData$, Len(strData$))
    r = RegCloseKey(keyhand)
End Sub
'end reg functions

'begin dialog functions
Public Function getopen$()
    'display open dialog, specific to this program
    Dim spath$, sfile$
    
    spath$ = regGet("last_path")
    If spath$ = "" Or Len(Dir(spath$)) = 0 Then
        spath$ = App.Path
    End If
    
    sfile$ = OpenDialog(frmMain, "all files (*.*)|*.*|text files (*.txt)|*.txt|", "L517 | load", spath$)
    
    Do While Right(sfile$, 1) = Chr(0)
        DoEvents
        sfile$ = Left(sfile$, Len(sfile$) - 1)
    Loop
    
    If sfile$ <> "" Then
        regSet "last_path", getdirectory(sfile$)
    End If
    
    getopen$ = sfile$
End Function
Public Function getsave$()
    'display save dialog, specific to this program
    Dim spath$, sfile$
    
    spath$ = regGet("last_path")
    If spath$ = "" Or Len(Dir(spath$)) = 0 Then
        spath$ = App.Path
    End If
    sfile$ = SaveDialog(frmMain, "text files(*.txt)|*.txt|all files (*.*)|*.*", "L517 | save", spath$, OFN_FILEMUSTEXIST)
    
    Do While Right(sfile$, 1) = Chr(0)
        DoEvents
        sfile$ = Left(sfile$, Len(sfile$) - 1)
    Loop
    
    If sfile$ <> "" Then
        regSet "last_path", getdirectory(sfile$)
    End If
    
    getsave$ = sfile$
End Function
Public Function OpenDialog(Form1 As Form, Filter As String, title As String, InitDir As String) As String
    'display open dialog
    Dim ofn As OPENFILENAME, a As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next
    
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = title
    ofn.flags = OFN_FILEMUSTEXIST
    a = GetOpenFileName(ofn)
    
    If (a) Then
        OpenDialog = Trim$(ofn.lpstrFile)
    Else
        OpenDialog = ""
    End If
End Function
Public Function SaveDialog(Form1 As Form, Filter As String, title As String, InitDir As String, F_FLAGS As Long) As String
    'display save dialog
    Dim ofn As OPENFILENAME, a As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next
    
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = title
    ofn.flags = F_FLAGS
    a = GetSaveFileName(ofn)
    
    If (a) Then
        SaveDialog = Trim$(ofn.lpstrFile)
    Else
        SaveDialog = ""
    End If
End Function
'end file dialog functions

'start directory functions
Public Function GetFolder(ByVal title As String, ByVal start As String, ByVal newfolder As Boolean) As String
    'opens Locate Directory dialog window, returns path
    Dim BI As BROWSEINFO, pidl As Long, lpSelPath As Long
    Dim spath As String * MAX_PATH
    
    'fill in the info it needs
    With BI
        .hOwner = GetForegroundWindow
        .pidlRoot = 0
        .lpszTitle = title
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
        .ulFlags = BIF_RETURNONLYFSDIRS
        If newfolder = True Then .ulFlags = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE
        lpSelPath = LocalAlloc(LPTR, Len(start) + 1)
        CopyMemory ByVal lpSelPath, ByVal start, Len(start) + 1
        .lParam = lpSelPath
    End With
    
    'get the idlist long from the returned folder
    pidl = SHBrowseForFolder(BI)
    
    'do then if they clicked ok
    If pidl Then
        If SHGetPathFromIDList(pidl, spath) Then
            'next line is the returned folder
            GetFolder = Left$(spath, InStr(spath, vbNullChar) - 1)
            
        End If
        Call CoTaskMemFree(pidl)
    Else
        'user clicked cancel
    End If
    
    Call LocalFree(lpSelPath)
End Function
Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    'this seems to happen before the box comes up and when a folder is clicked on within it
    Dim spath As String, bFlag As Long
                                       
    spath = Space$(MAX_PATH)
        
    Select Case uMsg
        Case BFFM_INITIALIZED
            'browse has been initialized, set the start folder
            Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal lpData)
        Case BFFM_SELCHANGED
            If SHGetPathFromIDList(lParam, spath) Then
                spath = Left(spath, InStr(1, spath, Chr(0)) - 1)
            End If
    End Select
End Function
Public Function FARPROC(pfn As Long) As Long
    'used for showing the directory dialog
    FARPROC = pfn
End Function
'end directory dialog functions

'web functions
Public Function webgetsource$(site$)
    On Error GoTo error_happened
    Dim objHttp As Object, strURL As String, strText As String
    
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    strURL = site$
    If Left(LCase(strURL$), Len("http")) <> "http" Then
        strURL$ = "http://" + strURL$
    End If
    
    objHttp.Open "GET", strURL, False
    objHttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.2.13) Gecko/20101203 Firefox/3.6.13"
    objHttp.Send ("")
    
    strText = objHttp.responseText

    Set objHttp = Nothing
    
    webgetsource$ = strText$
    Exit Function
error_happened:
    webgetsource$ = ""
End Function
'end web functions
