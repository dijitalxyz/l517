VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5085
   ClientLeft      =   2970
   ClientTop       =   1995
   ClientWidth     =   3225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5085
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComctlLib.ListView lst 
      Height          =   3855
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "passwords"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.PictureBox picWelcome 
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Left            =   345
      ScaleHeight     =   2235
      ScaleWidth      =   2475
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label lblDismiss 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(click anywhere to hide)"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label lblSite 
         BackStyle       =   0  'Transparent
         Caption         =   "http://code.google.com/p/l517"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   30
         MousePointer    =   2  'Cross
         TabIndex        =   19
         Top             =   1425
         Width           =   2475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWelcome 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " welcome to L517,  to access the readme, press F1. visit this program's website at:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   255
         TabIndex        =   18
         Top             =   0
         Width           =   2100
      End
   End
   Begin VB.ListBox lstFile 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   2040
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstDir 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   1320
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.DirListBox dir1 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picP 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3435
      Left            =   120
      ScaleHeight     =   3405
      ScaleWidth      =   2985
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblProg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2580
      OLEDropMode     =   1  'Manual
      TabIndex        =   16
      Top             =   4830
      Width           =   645
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
      TabIndex        =   15
      Top             =   4830
      Width           =   3195
   End
   Begin VB.Label lblMin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   -15
      OLEDropMode     =   1  'Manual
      TabIndex        =   14
      Top             =   0
      Width           =   300
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2925
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Top             =   0
      Width           =   300
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HELP"
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
      Height          =   255
      Left            =   2520
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblGen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GENERATE"
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
      Height          =   255
      Left            =   1320
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EDIT"
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
      Height          =   255
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FILE"
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
      Height          =   255
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L517"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
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
      TabIndex        =   12
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CANCEL"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblAppend2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "APPEND"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2145
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblCase2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CASE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   4500
      Width           =   735
   End
   Begin VB.Label lblFilter2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FILTER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1050
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   4500
      Width           =   855
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "filter"
      Visible         =   0   'False
      Begin VB.Menu mnuFilterLenMin 
         Caption         =   "minimum length: [none]"
      End
      Begin VB.Menu mnuFilterLenMax 
         Caption         =   "maximum length: [none]"
      End
      Begin VB.Menu mnuFilterDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilterTextRight 
         Caption         =   "text to the right of [string]"
      End
      Begin VB.Menu mnuFilterTextLeft 
         Caption         =   "text to the left of [string]"
      End
      Begin VB.Menu mnuFilterDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilterHex 
         Caption         =   "convert ! @ # $  ... to hex"
      End
      Begin VB.Menu mnuFilterForeign 
         Caption         =   "include foreign characters"
      End
   End
   Begin VB.Menu mnuCase 
      Caption         =   "case"
      Visible         =   0   'False
      Begin VB.Menu mnuCaseConvertto 
         Caption         =   "     [convert list to:]"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCaseLower 
         Caption         =   "&lowercase"
      End
      Begin VB.Menu mnuCaseUpper 
         Caption         =   "&UPPERCASE"
      End
      Begin VB.Menu mnuCaseDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCaseCopies 
         Caption         =   "   [also make copies with:]"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCaseFirst 
         Caption         =   "&First Letter Uppercase"
      End
      Begin VB.Menu mnuCaseEveryother 
         Caption         =   "&eVeRy OtHeR uPpEr"
      End
      Begin VB.Menu mnuCaseDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCaseLeet 
         Caption         =   "&1337c4s3 (mutations!)"
      End
      Begin VB.Menu mnuCaseLeetEdit 
         Caption         =   "&edit ""leetspeak"" ..."
      End
   End
   Begin VB.Menu mnuAppend 
      Caption         =   "append"
      Visible         =   0   'False
      Begin VB.Menu mnuAppendPre 
         Caption         =   "add to &left of each item"
         Begin VB.Menu mnuAppendPreCustom 
            Caption         =   "&custom file..."
         End
         Begin VB.Menu mnuAppendPreDefault 
            Caption         =   "&default prefixes"
         End
         Begin VB.Menu mnuAppendPreDash1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAppendPreNum 
            Caption         =   "&numeric"
         End
         Begin VB.Menu mnuAppendPreAlpha 
            Caption         =   "&alpha"
         End
         Begin VB.Menu mnuAppendPreDash2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAppendPreString 
            Caption         =   "&string..."
         End
      End
      Begin VB.Menu mnuAppendPost 
         Caption         =   "add to the &right of each item"
         Begin VB.Menu mnuAppendPostCustom 
            Caption         =   "&custom file..."
         End
         Begin VB.Menu mnuAppendPostDefault 
            Caption         =   "&default postfixes"
         End
         Begin VB.Menu mnuAppendPostDash1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAppendPostNum 
            Caption         =   "&numeric"
         End
         Begin VB.Menu mnuAppendPostAlpha 
            Caption         =   "&alpha"
         End
         Begin VB.Menu mnuAppendPostDash2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAppendPostString 
            Caption         =   "&string..."
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileNew 
         Caption         =   "&new"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&save as..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSplit 
         Caption         =   "sp&lit into files"
         Begin VB.Menu mnuFileSplitNever 
            Caption         =   "never"
         End
         Begin VB.Menu mnuFileSplit50000 
            Caption         =   "every 50,000 words"
         End
         Begin VB.Menu mnuFileSplit100000 
            Caption         =   "every 100,000 words"
         End
         Begin VB.Menu mnuFileSplit1000000 
            Caption         =   "every 1,000,000 words"
         End
         Begin VB.Menu mnuFileSplitDash1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileSplitCustom 
            Caption         =   "custom [every # words]"
         End
      End
      Begin VB.Menu mnuFileType 
         Caption         =   "save &type"
         Begin VB.Menu mnuFileTypeWin 
            Caption         =   "&windows"
         End
         Begin VB.Menu mnuFileTypeUnix 
            Caption         =   "&unix"
         End
      End
      Begin VB.Menu mnuFileProfiles 
         Caption         =   "length &presets"
         Begin VB.Menu mnuFilesProfilesWPA 
            Caption         =   "&wpa passwords (length 8-63)"
         End
         Begin VB.Menu mnuFileProfilesWeb 
            Caption         =   "web &passwords (length 4-12)"
         End
      End
      Begin VB.Menu mnuFileDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnuListDupekill 
         Caption         =   "&remove duplicates"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuListDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListRemove 
         Caption         =   "remove &item"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuListRemoveItemsWith 
         Caption         =   "remove item&s..."
      End
      Begin VB.Menu mnuListClear 
         Caption         =   "&clear list"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuListDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListPaste 
         Caption         =   "&paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuListDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListFind 
         Caption         =   "&find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuListFindNext 
         Caption         =   "find &next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuGen 
      Caption         =   "Generate"
      Visible         =   0   'False
      Begin VB.Menu mnuGenWeb 
         Caption         =   "words from &website"
      End
      Begin VB.Menu mnuGenFiles 
         Caption         =   "words from &folder(s)"
      End
      Begin VB.Menu mnuGenDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenString 
         Caption         =   "&string from charset..."
         Begin VB.Menu mnuGenStringHelp 
            Caption         =   "what does this do?"
         End
         Begin VB.Menu mnuGenStringDash0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGenStringEdit 
            Caption         =   "edit list of charsets"
         End
         Begin VB.Menu mnuGenStringDash1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGenStringX 
            Caption         =   "x"
            Index           =   0
         End
      End
      Begin VB.Menu mnuGenDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenDateS 
         Caption         =   "&dates"
         Begin VB.Menu mnuGenDateSep 
            Caption         =   "&separator: [none]"
         End
         Begin VB.Menu mnuGenDateDash1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGenDate 
            Caption         =   "mm/dd/yy"
            Index           =   0
         End
         Begin VB.Menu mnuGenDate 
            Caption         =   "mm/dd/yyyy"
            Index           =   1
         End
         Begin VB.Menu mnuGenDate 
            Caption         =   "dd/mm/yy"
            Index           =   2
         End
         Begin VB.Menu mnuGenDate 
            Caption         =   "dd/mm/yyyy"
            Index           =   3
         End
         Begin VB.Menu mnuGenDate 
            Caption         =   "mmm/dd/yy"
            Index           =   4
         End
         Begin VB.Menu mnuGenDate 
            Caption         =   "mmm/dd/yyyy"
            Index           =   5
         End
         Begin VB.Menu mnuGenDate 
            Caption         =   "dd/mmm/yy"
            Index           =   6
         End
         Begin VB.Menu mnuGenDate 
            Caption         =   "dd/mmm/yyyy"
            Index           =   7
         End
      End
      Begin VB.Menu mnuGenPhone 
         Caption         =   "&phone numbers"
         Begin VB.Menu mnuGenPhoneSep 
            Caption         =   "separator: -"
         End
         Begin VB.Menu mnuGenPhoneDash1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGenPhoneArea 
            Caption         =   "[areacode]-[prefix]-####"
         End
         Begin VB.Menu mnuGenPhoneNoarea 
            Caption         =   "[prefix]-####"
         End
         Begin VB.Menu mnuGenPhoneDash2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGenPhoneCustom 
            Caption         =   "custom..."
         End
      End
      Begin VB.Menu mnuGenDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenAnal 
         Caption         =   "anal&yze patterns"
         Begin VB.Menu mnuGenAnalANAL 
            Caption         =   "&analyze!"
         End
         Begin VB.Menu mnuGenAnalDash1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGenAnalMinPre 
            Caption         =   "&min. length of p&refix: [0]"
         End
         Begin VB.Menu mnuGenAnalMinPost 
            Caption         =   "min. length of p&ostfix: [0]"
         End
         Begin VB.Menu mnuGenAnalDash3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGenAnalCount 
            Caption         =   "minimum &count: [0]"
         End
         Begin VB.Menu mnuGenAnalCase 
            Caption         =   "&ignore case"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuGenAnalDash2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGenAnalHelp 
            Caption         =   "&what does this do?"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&about"
      End
      Begin VB.Menu mnuHelpItems 
         Caption         =   "&items not loading?"
      End
      Begin VB.Menu mnuHelpDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpEnglish 
         Caption         =   "english language"
      End
      Begin VB.Menu mnuHelpSpanish 
         Caption         =   "idioma español (spanish)"
      End
      Begin VB.Menu mnuHelpFrench 
         Caption         =   "en français (french)"
      End
      Begin VB.Menu mnuHelpGerman 
         Caption         =   "deutsch sprache (german)"
      End
      Begin VB.Menu mnuHelpDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpGetlists 
         Caption         =   "&get some wordlists!"
      End
      Begin VB.Menu mnuHelpSite 
         Caption         =   "&project homepage"
      End
      Begin VB.Menu mnuHelpDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&readme"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TODO

'spidering in web-search?
' \_ possibly? might get out of hand...


'version history

'0.96: added leetspeak/charset editing -- removed charset.lst and leetspeak.txt


'0.95: lots of random fixes

'0.94: fixed string generator bug

'0.93: removed mscomctl.ocx from the resource file - no more AV alerts

'0.8 : language support for:
'       -english
'       -french
'       -german
'       -spanish

'0.7 : added leetspeak mutations     1/20/10

'0.6 : added paste to menu,          1/19/10
'      various bug fixes,

'0.5 : fixed case bugs.              1/18/10

'0.4 : fixed RICHTX32.OCX error,     1/18/10
'      removed RTF control,
'      using built-in Word API's for doc's

'0.3 : phonenumber list generation,  1/17/10
'      charset support (with predictions),
'      every other and leetspeak cases,
'      split files,

'0.2 : analyzer option,              1/10/10
'      help documentation

'0.1 : first release, base options.  1/06/10

Option Explicit

Public B_CHANGE As Boolean, B_MOUSEDOWN As Boolean, C_MOVER As Boolean, C_FORMX As Single, C_FORMY As Single, LANGUAGE$, RESUME_STRING$

Public Function defaultCharset$()
    Dim s$
         s$ = "numeric                          = [0123456789]" + vbCrLf
    s$ = s$ + "numeric-space                    = [0123456789 ]" + vbCrLf
    s$ = s$ + "" + vbCrLf
    s$ = s$ + "ualpha                           = [ABCDEFGHIJKLMNOPQRSTUVWXYZ]" + vbCrLf
    s$ = s$ + "ualpha-space                     = [ABCDEFGHIJKLMNOPQRSTUVWXYZ ]" + vbCrLf
    s$ = s$ + "ualpha-numeric                   = [ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789]" + vbCrLf
    s$ = s$ + "ualpha-numeric-space             = [ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ]" + vbCrLf
    s$ = s$ + "ualpha-numeric-symbol            = [ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_+=]" + vbCrLf
    s$ = s$ + "ualpha-numeric-symbol-space      = [ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_+= ]" + vbCrLf
    s$ = s$ + "ualpha-numeric-all               = [ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_+=~`[]{}|\:;""'<>,.?/]" + vbCrLf
    s$ = s$ + "ualpha-numeric-all-space         = [ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_+=~`[]{}|\:;""'<>,.?/ ]" + vbCrLf
    s$ = s$ + "" + vbCrLf
    s$ = s$ + "lalpha                           = [abcdefghijklmnopqrstuvwxyz]" + vbCrLf
    s$ = s$ + "lalpha-space                     = [abcdefghijklmnopqrstuvwxyz ]" + vbCrLf
    s$ = s$ + "lalpha-numeric                   = [abcdefghijklmnopqrstuvwxyz0123456789]" + vbCrLf
    s$ = s$ + "lalpha-numeric-space             = [abcdefghijklmnopqrstuvwxyz0123456789 ]" + vbCrLf
    s$ = s$ + "lalpha-numeric-symbol            = [abcdefghijklmnopqrstuvwxyzäöüß0123456789!@#$%^&*()-_+=""]" + vbCrLf
    s$ = s$ + "lalpha-numeric-symbol-space      = [abcdefghijklmnopqrstuvwxyzäöüß0123456789!@#$%^&*()-_+="" ]" + vbCrLf
    s$ = s$ + "lalpha-numeric-all               = [abcdefghijklmnopqrstuvwxyzäöüß0123456789!@#$%^&*()-_+=~`[]{}|\:;""'<>,.?/]" + vbCrLf
    s$ = s$ + "lalpha-numeric-all-space         = [abcdefghijklmnopqrstuvwxyzäöüß0123456789!@#$%^&*()-_+=~`[]{}|\:;""'<>,.?/ ]" + vbCrLf
    s$ = s$ + "" + vbCrLf
    s$ = s$ + "mixalpha                         = [abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ]" + vbCrLf
    s$ = s$ + "mixalpha-space                   = [abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ ]" + vbCrLf
    s$ = s$ + "mixalpha-numeric                 = [abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789]" + vbCrLf
    s$ = s$ + "mixalpha-numeric-space           = [abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ]" + vbCrLf
    s$ = s$ + "mixalpha-numeric-symbol          = [abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_+=]" + vbCrLf
    s$ = s$ + "mixalpha-numeric-symbol-space    = [abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_+= ]" + vbCrLf
    s$ = s$ + "mixalpha-numeric-all             = [abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_+=~`[]{}|\:;""'<>,.?/]" + vbCrLf
    s$ = s$ + "mixalpha-numeric-all-space       = [abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_+=~`[]{}|\:;""'<>,.?/ ]"
    defaultCharset$ = s$
End Function

Public Sub Change_Language()
    'does not change status bar information,
    '                message / input boxes,
    '                printed text (readme)
    
    Select Case LANGUAGE$
    Case "english"
        'labels
        lblFile.Width = 615
        lblEdit.Width = 735
        lblGen.Width = 1200
        lblHelp.Width = 735
        
        lblEdit.Left = lblFile.Left + lblFile.Width - 15
        lblGen.Left = lblEdit.Left + lblEdit.Width - 15
        lblHelp.Left = lblGen.Left + lblGen.Width - 15
        
        lblFile.Caption = "FILE"
        lblEdit.Caption = "EDIT"
        lblGen.Caption = "GENERATE"
        lblHelp.Caption = "HELP"
        lblCase2.Caption = "CASE"
        lblFilter2.Caption = "FILTER"
        lblAppend2.Caption = "APPEND"
        
        '[file]
        mnuFileNew.Caption = "&new"
        mnuFileOpen.Caption = "&open"
        mnuFileSave.Caption = "&save as..."
        mnuFileSplit.Caption = "s&plit into files"
        mnuFileSplitNever.Caption = "&never"
        mnuFileSplit50000.Caption = "every &50,000"
        mnuFileSplit100000.Caption = "every &100,000"
        mnuFileSplit1000000.Caption = "every 1,&000,000"
        mnuFileSplitCustom.Caption = "custom [every " + IIf(mnuFileSplitCustom.Tag = "", "#", mnuFileSplitCustom.Tag) + " words]"
        mnuFileType.Caption = "save &type"
        mnuFileProfiles.Caption = "length p&resets"
        mnuFileProfilesWeb.Caption = "&web passwords (length 4-12)"
        mnuFilesProfilesWPA.Caption = "w&pa passwords (length (8-63)"
        mnuFileExit.Caption = "&quit"
        
        '[edit]
        mnuListDupekill.Caption = "remove &duplicates"
        mnuListRemove.Caption = "&remove item"
        mnuListRemoveItemsWith.Caption = "remove item&s..."
        mnuListClear.Caption = "&clear"
        mnuListPaste.Caption = "&paste"
        mnuListFind.Caption = "&find"
        mnuListFindNext.Caption = "find &next"
        
        '[generate]
        mnuGenWeb.Caption = "words from &site"
        mnuGenFiles.Caption = "words from &folder(s)"
        mnuGenString.Caption = "string from &charset..."
        mnuGenStringHelp.Caption = "&what does this do?"
        mnuGenDateS.Caption = "&dates"
        mnuGenDateSep.Caption = "&separator: [" + IIf(mnuGenDateSep.Tag = "", "none", mnuGenDateSep.Tag) + "]"
        mnuGenPhone.Caption = "&phone numbers"
        mnuGenPhoneSep.Caption = "&separator: [" + IIf(mnuGenPhoneSep.Tag = "", "none", mnuGenDateSep.Tag) + "]"
        mnuGenAnal.Caption = "&analyze patterns..."
        mnuGenAnalANAL.Caption = "&analyze!"
        mnuGenAnalMinPre.Caption = "min. length of &prefix: [" + mnuGenAnalMinPre.Tag + "]"
        mnuGenAnalMinPost.Caption = "min. length of p&ostfix: [" + mnuGenAnalMinPost.Tag + "]"
        mnuGenAnalCount.Caption = "minimum &count: [" + regGet("pattern_count") + "]"
        mnuGenAnalCase.Caption = "&ignore case"
        mnuGenAnalHelp.Caption = "&what does this do?"
        
        '[help]
        mnuHelpAbout.Caption = "&about"
        mnuHelpItems.Caption = "&items not loading?"
        Me.mnuHelpGetlists.Caption = "&get some wordlists!"
        mnuHelpSite.Caption = "&project homepage"
        mnuHelpHelp.Caption = "&readme"
        
        '[case]
        Me.mnuCaseConvertto.Caption = "   [convert list to:]"
        mnuCaseLower.Caption = "lowercase"
        mnuCaseUpper.Caption = "UPPERCASE"
        Me.mnuCaseCopies.Caption = "   [make copies with:]"
        Me.mnuCaseFirst.Caption = "First Letter Upper"
        Me.mnuCaseEveryother.Caption = "EvErY oThEr LeTtEr"
        mnuCaseLeet.Caption = "13375P34K (mutations!)"
        
        '[filter]
        mnuFilterLenMin.Caption = "&minimum length: " + IIf(mnuFilterLenMin.Tag = "0", "[none]", mnuFilterLenMin.Tag)
        mnuFilterLenMax.Caption = "ma&ximum length: " + IIf(mnuFilterLenMax.Tag = "0", "[none]", mnuFilterLenMax.Tag)
        mnuFilterTextRight.Caption = "text to the &right of [" + IIf(mnuFilterTextRight.Tag = "", "string", mnuFilterTextRight.Tag) + "]"
        mnuFilterTextLeft.Caption = "text to the &left of [" + IIf(mnuFilterTextLeft.Tag = "", "string", mnuFilterTextLeft.Tag) + "]"
        mnuFilterHex.Caption = "&convert ! @ # $ ... to hex"
        mnuFilterForeign.Caption = "include &foreign characters"
        
        '[append]
        mnuAppendPre.Caption = "add to &left of string"
        mnuAppendPreCustom.Caption = "&custom file..."
        mnuAppendPreDefault.Caption = "&default prefixes"
        mnuAppendPreNum.Caption = "&numeric"
        mnuAppendPreAlpha.Caption = "&alphabet"
        mnuAppendPreString.Caption = "&string..."
        
        mnuAppendPost.Caption = "add to &right of string"
        mnuAppendPostCustom.Caption = "&custom file..."
        mnuAppendPostDefault.Caption = "&default postfixes"
        mnuAppendPostNum.Caption = "&numeric"
        mnuAppendPostAlpha.Caption = "&alphabet"
        mnuAppendPostString.Caption = "&string..."
    Case "french"
        'labels
        lblFile.Width = 900
        lblEdit.Width = 735
        lblGen.Width = 900
        lblHelp.Width = 735
        lblEdit.Left = lblFile.Left + lblFile.Width - 15
        lblGen.Left = lblEdit.Left + lblEdit.Width - 15
        lblHelp.Left = lblGen.Left + lblGen.Width - 15
        
        lblFile.Caption = "DOSSIER"
        lblEdit.Caption = "MODIF"
        lblGen.Caption = "GÉNÉRER"
        lblHelp.Caption = "AIDER"
        lblCase2.Caption = "CAS"
        lblFilter2.Caption = "FILTRE"
        lblAppend2.Caption = "JOINDRE"
        
        '[file]
        mnuFileNew.Caption = "&nouveau"
        mnuFileOpen.Caption = "&ouvert"
        mnuFileSave.Caption = "&enregistrer sous..."
        mnuFileSplit.Caption = "&fractionné en plusieurs fichiers"
        mnuFileSplitNever.Caption = "&jamais"
        mnuFileSplit50000.Caption = "toutes les &50,000"
        mnuFileSplit100000.Caption = "toutes les&100,000"
        mnuFileSplit1000000.Caption = "toutes les1,&000,000"
        mnuFileSplitCustom.Caption = "&personnalisé [toutes les" + IIf(mnuFileSplitCustom.Tag = "", "#", mnuFileSplitCustom.Tag) + " mots]"
        mnuFileType.Caption = "enregistrer le &type"
        mnuFileProfiles.Caption = "p&resets"
        mnuFileProfilesWeb.Caption = "&web passwords"
        mnuFilesProfilesWPA.Caption = "w&pa passwords"
        mnuFileExit.Caption = "&sortie"
        
        '[edit]
        mnuListDupekill.Caption = "supprimer les &doublons"
        mnuListRemove.Caption = "&supprimer l'élément"
        mnuListRemoveItemsWith.Caption = "supprimer des élément&s..."
        mnuListClear.Caption = "&détruire"
        mnuListPaste.Caption = "&coller"
        mnuListFind.Caption = "&rechercher"
        mnuListFindNext.Caption = "rechercher &suivant"
        
        '[generate]
        mnuGenWeb.Caption = "mots de &site web"
        mnuGenFiles.Caption = "mots de &répertoire(s)"
        mnuGenString.Caption = "mots de &charset..."
        mnuGenStringHelp.Caption = "&qu'est ce que cela fait?"
        mnuGenDateS.Caption = "&dates"
        mnuGenDateSep.Caption = "&séparateur: [" + IIf(mnuGenDateSep.Tag = "", "aucun", mnuGenDateSep.Tag) + "]"
        mnuGenPhone.Caption = "&numéro de téléphone"
        mnuGenPhoneSep.Caption = "&séparateur: [" + IIf(mnuGenPhoneSep.Tag = "", "aucun", mnuGenPhoneSep.Tag) + "]"
        mnuGenAnal.Caption = "&d'analyser les habitudes..."
        mnuGenAnalANAL.Caption = "&analyser!"
        mnuGenAnalMinPre.Caption = "min. longueur du &préfixe: [" + mnuGenAnalMinPre.Tag + "]"
        mnuGenAnalMinPost.Caption = "min. longueur de p&ostfix: [" + mnuGenAnalMinPost.Tag + "]"
        mnuGenAnalCount.Caption = "minimum &nombre: [" + regGet("pattern_count") + "]"
        mnuGenAnalCase.Caption = "&ignorer case"
        mnuGenAnalHelp.Caption = "&qu'est ce que cela fait?"
        
        '[help]
        mnuHelpAbout.Caption = "&à propos de"
        mnuHelpItems.Caption = "&éléments qui ne présentent pas?"
        Me.mnuHelpGetlists.Caption = "&obtenir des listes de mots!"
        mnuHelpSite.Caption = "&L517 homepage"
        mnuHelpHelp.Caption = "&lisez-moi"
        
        '[case]
        Me.mnuCaseConvertto.Caption = "   [convertir les éléments de:]"
        mnuCaseLower.Caption = "minuscules"
        mnuCaseUpper.Caption = "MAJUSCULES"
        Me.mnuCaseCopies.Caption = "   [faire des copies avec:]"
        Me.mnuCaseFirst.Caption = "Première Lettre Haute"
        Me.mnuCaseEveryother.Caption = "ChAqUe LeTtRe AuTrEs"
        mnuCaseLeet.Caption = "13375P34K (mutations!)"
        
        '[filter]
        mnuFilterLenMin.Caption = "&minimum longueur: " + IIf(mnuFilterLenMin.Tag = "0", "[aucun]", mnuFilterLenMin.Tag)
        mnuFilterLenMax.Caption = "ma&ximale longueur: " + IIf(mnuFilterLenMax.Tag = "0", "[aucun]", mnuFilterLenMax.Tag)
        mnuFilterTextRight.Caption = "texte à &droite de [" + IIf(mnuFilterTextRight.Tag = "", "string", mnuFilterTextRight.Tag)
        mnuFilterTextLeft.Caption = "texte à &gauche [" + IIf(mnuFilterTextLeft.Tag = "", "string", mnuFilterTextLeft.Tag)
        mnuFilterHex.Caption = "&convertir ! @ # $ ... à hex"
        mnuFilterForeign.Caption = "&inclure des caractères étrangers ("
        
        '[append]
        mnuAppendPre.Caption = "jouter à &gauche des éléments"
        mnuAppendPreCustom.Caption = "&personnalisés liste..."
        mnuAppendPreDefault.Caption = "préfixes par &défaut"
        mnuAppendPreNum.Caption = "&numérique"
        mnuAppendPreAlpha.Caption = "&alphabet"
        mnuAppendPreString.Caption = "&mot..."
        
        mnuAppendPost.Caption = "ajouter à &droite des éléments"
        mnuAppendPostCustom.Caption = "&personnalisés liste..."
        mnuAppendPostDefault.Caption = "postfixes par &défaut"
        mnuAppendPostNum.Caption = "&numérique"
        mnuAppendPostAlpha.Caption = "&alphabet"
        mnuAppendPostString.Caption = "&mot..."
    Case "german"
        'labels
        lblFile.Width = 750
        lblEdit.Width = 800
        lblGen.Width = 1000
        lblHelp.Width = 735
        lblEdit.Left = lblFile.Left + lblFile.Width - 15
        lblGen.Left = lblEdit.Left + lblEdit.Width - 15
        lblHelp.Left = lblGen.Left + lblGen.Width - 15
        
        lblFile.Caption = "DATEI"
        lblEdit.Caption = "CUTTEN"
        lblGen.Caption = "ERZEUGEN"
        lblHelp.Caption = "HILFE"
        lblCase2.Caption = "CASE"
        lblFilter2.Caption = "FILTER"
        lblAppend2.Caption = "APPEND"
        
        '[file]
        mnuFileNew.Caption = "&neu"
        mnuFileOpen.Caption = "&öffnen"
        mnuFileSave.Caption = "&ziel speichern unter..."
        mnuFileSplit.Caption = "&Aufspaltung in separate Dateien"
        mnuFileSplitNever.Caption = "&nei"
        mnuFileSplit50000.Caption = "alle &50,000"
        mnuFileSplit100000.Caption = "alle &100,000"
        mnuFileSplit1000000.Caption = "alle 1,&000,000"
        mnuFileSplitCustom.Caption = "&brauch [alle " + IIf(mnuFileSplitCustom.Tag = "", "#", mnuFileSplitCustom.Tag) + " worte]"
        mnuFileType.Caption = "&dateityp"
        mnuFileProfiles.Caption = "p&resets"
        mnuFileProfilesWeb.Caption = "&web passwörter"
        mnuFilesProfilesWPA.Caption = "w&pa passwörter"
        mnuFileExit.Caption = "&ausfahrt"
        
        '[edit]
        mnuListDupekill.Caption = "&entfernen von duplikaten"
        mnuListRemove.Caption = "&Artikel entfernen"
        mnuListRemoveItemsWith.Caption = "&Elemente entfernen..."
        mnuListClear.Caption = "&klar"
        mnuListPaste.Caption = "&Einfügen"
        mnuListFind.Caption = "&finden"
        mnuListFindNext.Caption = "&Weitersuchen"
        
        '[generate]
        mnuGenWeb.Caption = "Worte von der Website"
        mnuGenFiles.Caption = "Worte aus dem Ordner(n)"
        mnuGenString.Caption = "Wörter aus &charset..."
        mnuGenStringHelp.Caption = "&Was bedeutet das?"
        mnuGenDateS.Caption = "&Datum"
        mnuGenDateSep.Caption = "&separator: [" + IIf(mnuGenDateSep.Tag = "", "keine", mnuGenDateSep.Tag) + "]"
        mnuGenPhone.Caption = "&Telefonnummern"
        mnuGenPhoneSep.Caption = "&separator: [" + IIf(mnuGenPhoneSep.Tag = "", "keine", mnuGenPhoneSep.Tag) + "]"
        mnuGenAnal.Caption = "&Analyse von Mustern..."
        mnuGenAnalANAL.Caption = "&analyse!"
        mnuGenAnalMinPre.Caption = "min. Länge des &Präfix: [" + mnuGenAnalMinPre.Tag + "]"
        mnuGenAnalMinPost.Caption = "min. Länge des p&ostfix: [" + mnuGenAnalMinPost.Tag + "]"
        mnuGenAnalCount.Caption = "Mindest-&count: [" + regGet("pattern_count") + "]"
        mnuGenAnalCase.Caption = "&ignorieren case"
        mnuGenAnalHelp.Caption = "&Was bedeutet das?"
        
        '[help]
        mnuHelpAbout.Caption = "&über"
        mnuHelpItems.Caption = "&Elemente werden nicht angezeigt?"
        Me.mnuHelpGetlists.Caption = "&download wordlists!"
        mnuHelpSite.Caption = "&Projekt-Homepage"
        mnuHelpHelp.Caption = "&readme"
        
        '[case]
        Me.mnuCaseConvertto.Caption = "   [konvertieren Elemente:]"
        mnuCaseLower.Caption = "kleinbuchstaben"
        mnuCaseUpper.Caption = "GROß"
        Me.mnuCaseCopies.Caption = "   [Kopien mit:]"
        Me.mnuCaseFirst.Caption = "Anfangsbuchstaben Oberen"
        Me.mnuCaseEveryother.Caption = "AlLe AnDeReN bUcHsTaBeN"
        mnuCaseLeet.Caption = "13375P34K (mutationen!)"
        
        '[filter]
        mnuFilterLenMin.Caption = "&Mindestlänge: " + IIf(mnuFilterLenMin.Tag = "0", "[keine]", mnuFilterLenMin.Tag)
        mnuFilterLenMax.Caption = "ma&ximale Länge: " + IIf(mnuFilterLenMax.Tag = "0", "[keine]", mnuFilterLenMax.Tag)
        mnuFilterTextRight.Caption = "Text auf der &rechten Seite [" + IIf(mnuFilterTextRight.Tag = "", "string", mnuFilterTextRight.Tag)
        mnuFilterTextLeft.Caption = "Text auf der &linken Seite [" + IIf(mnuFilterTextLeft.Tag = "", "string", mnuFilterTextLeft.Tag)
        mnuFilterHex.Caption = "&konvertieren ! @ # $ ... to hex"
        mnuFilterForeign.Caption = "gehören &fremden Buchstaben"
        
        '[append]
        mnuAppendPre.Caption = "in den &links-String"
        mnuAppendPreCustom.Caption = "&benutzerdefinierte Liste..."
        mnuAppendPreDefault.Caption = "&Default Präfixe"
        mnuAppendPreNum.Caption = "&Zahlen"
        mnuAppendPreAlpha.Caption = "&Alphabet"
        mnuAppendPreString.Caption = "&wort..."
        
        mnuAppendPost.Caption = "in den Rechts-String"
        mnuAppendPostCustom.Caption = "&benutzerdefinierte Liste..."
        mnuAppendPostDefault.Caption = "&Default Postfixe"
        mnuAppendPostNum.Caption = "&Zahlen"
        mnuAppendPostAlpha.Caption = "&alphabet"
        mnuAppendPostString.Caption = "&wort..."
    Case "spanish"
        'labels
        lblFile.Width = 850
        lblEdit.Width = 800
        lblGen.Width = 900
        lblHelp.Width = 735
        lblEdit.Left = lblFile.Left + lblFile.Width - 15
        lblGen.Left = lblEdit.Left + lblEdit.Width - 15
        lblHelp.Left = lblGen.Left + lblGen.Width - 15
        
        lblFile.Caption = "ARCHIVO"
        lblEdit.Caption = "EDITAR"
        lblGen.Caption = "GENERAR"
        lblHelp.Caption = "AYUDA"
        lblCase2.Caption = "CASO"
        lblFilter2.Caption = "FILTRO"
        lblAppend2.Caption = "ANEXAR"
        
        '[file]
        mnuFileNew.Caption = "&nuevo"
        mnuFileOpen.Caption = "&abrir"
        mnuFileSave.Caption = "&salve como..."
        mnuFileSplit.Caption = "&dividido en archivos"
        mnuFileSplitNever.Caption = "&nunca"
        mnuFileSplit50000.Caption = "cada &50,000"
        mnuFileSplit100000.Caption = "cada &100,000"
        mnuFileSplit1000000.Caption = "cada 1,&000,000"
        mnuFileSplitCustom.Caption = "personalizada [cada " + IIf(mnuFileSplitCustom.Tag = "", "#", mnuFileSplitCustom.Tag) + " palabras]"
        mnuFileType.Caption = "salve &tipo"
        mnuFileProfiles.Caption = "p&resets"
        mnuFileProfilesWeb.Caption = "&web contraseñas"
        mnuFilesProfilesWPA.Caption = "w&pa contraseñas"
        mnuFileExit.Caption = "&salida"
        
        '[edit]
        mnuListDupekill.Caption = "eliminar &duplicados"
        mnuListRemove.Caption = "&eliminar el elemento"
        mnuListRemoveItemsWith.Caption = "eliminar los elemento&s..."
        mnuListClear.Caption = "&borrar la lista"
        mnuListPaste.Caption = "&pegar"
        mnuListFind.Caption = "&encontrar"
        mnuListFindNext.Caption = "&Buscar siguiente"
        
        '[generate]
        mnuGenWeb.Caption = "palabras del &sitio web"
        mnuGenFiles.Caption = "palabras de la &carpeta(s)"
        mnuGenString.Caption = "palabras de &charset..."
        mnuGenStringHelp.Caption = "¿&qué hace esto?"
        mnuGenDateS.Caption = "&fechas"
        mnuGenDateSep.Caption = "&separador: [" + IIf(mnuGenDateSep.Tag = "", "ninguno", mnuGenDateSep.Tag) + "]"
        mnuGenPhone.Caption = "&números de teléfono"
        mnuGenPhoneSep.Caption = "&separador: [" + IIf(mnuGenPhoneSep.Tag = "", "ninguno", mnuGenPhoneSep.Tag) + "]"
        mnuGenAnal.Caption = "&analizar los patrones..."
        mnuGenAnalANAL.Caption = "&analizar!"
        mnuGenAnalMinPre.Caption = "longitud mínima de prefijo: [" + mnuGenAnalMinPre.Tag + "]"
        mnuGenAnalMinPost.Caption = "longitud mínima de p&ostfix: [" + mnuGenAnalMinPost.Tag + "]"
        mnuGenAnalCount.Caption = "&conteo mínimo: [" + regGet("pattern_count") + "]"
        mnuGenAnalCase.Caption = "&Ignorar mayúsculas"
        mnuGenAnalHelp.Caption = "¿&qué hace esto?"
        
        '[help]
        mnuHelpAbout.Caption = "&sobre"
        mnuHelpItems.Caption = "&son elementos que no se muestra?"
        Me.mnuHelpGetlists.Caption = "&algunas listas de palabras descargar!"
        mnuHelpSite.Caption = "&página de inicio de L517"
        mnuHelpHelp.Caption = "&Léame"
        
        '[case]
        Me.mnuCaseConvertto.Caption = "   [convertir a la lista:]"
        mnuCaseLower.Caption = "minusculas"
        mnuCaseUpper.Caption = "MAYUSCULAS"
        Me.mnuCaseCopies.Caption = "   [hacer copias con:]"
        Me.mnuCaseFirst.Caption = "Primera Letra De La Parte Superior"
        Me.mnuCaseEveryother.Caption = "ToDaS lAs LeTrAs De OtRoS"
        mnuCaseLeet.Caption = "13375P34K (mutaciones!)"
        
        '[filter]
        mnuFilterLenMin.Caption = "longitud &mínima: " + IIf(mnuFilterLenMin.Tag = "0", "[ninguno]", mnuFilterLenMin.Tag)
        mnuFilterLenMax.Caption = "longitud má&ximo: " + IIf(mnuFilterLenMax.Tag = "0", "[ninguno]", mnuFilterLenMax.Tag)
        mnuFilterTextRight.Caption = "texto a la &derecha de la [" + IIf(mnuFilterTextRight.Tag = "", "string", mnuFilterTextRight.Tag)
        mnuFilterTextLeft.Caption = "texto a la &izquierda de la [" + IIf(mnuFilterTextLeft.Tag = "", "string", mnuFilterTextLeft.Tag)
        mnuFilterHex.Caption = "! @ # $ convertir a &hexadecimal"
        mnuFilterForeign.Caption = "incluir letras &extranjeras"
        
        '[append]
        mnuAppendPre.Caption = "añadir a la &izquierda de los elementos"
        mnuAppendPreCustom.Caption = "&lista personalizada..."
        mnuAppendPreDefault.Caption = "&prefijos por defecto"
        mnuAppendPreNum.Caption = "&numeric"
        mnuAppendPreAlpha.Caption = "&alphabet"
        mnuAppendPreString.Caption = "&string..."
        
        mnuAppendPost.Caption = "añadir a la derecha de los elementos"
        mnuAppendPostCustom.Caption = "&lista personalizada..."
        mnuAppendPostDefault.Caption = "&sufijos por defecto"
        mnuAppendPostNum.Caption = "&números"
        mnuAppendPostAlpha.Caption = "&letras (alfabeto)"
        mnuAppendPostString.Caption = "&palabra..."
    Case Else
        'unsupported!
    End Select
End Sub

'b_change=true if list is changed
'mousedown,mover,formx,formy=used for moving the main form

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'handle hotkeys
    
    If Shift = 2 Then
        'control [ctrl] is held
        Select Case KeyCode
        Case Asc("F")
            mnuListFind_Click
        Case Asc("N")
            mnuFileNew_Click
        Case Asc("O")
            mnuFileOpen_Click
        Case Asc("A")
            mnuFileOpen_Click
        Case Asc("S")
            mnuFileSave_Click
        Case Asc("Q")
            mnuFileExit_Click
        Case Asc("D")
            mnuListDupekill_Click
        Case Asc("V")
            stat "reading clipboard..."
            ParseTextBlock Clipboard.GetText
            DoEvents
            UpdateCaption
            stat ""
        End Select
    ElseIf Shift = 1 Then
        'shift is held
        If KeyCode = vbKeyEscape Then
            mnuListClear_Click
        End If
    ElseIf Shift = 4 Then
        'alt is held
        Select Case KeyCode
        Case 115
            'alt+f4
            mnuFileExit_Click
        End Select
    End If
    
    Select Case KeyCode
    Case 114
        mnuListFindNext_Click
    Case 112
        mnuHelpHelp_Click
    Case 46
        mnuListRemove_Click
    End Select
End Sub

Public Function defaultLeetSpeak$()
    Dim s$
    
        s$ = "a=a,4,@" + vbCrLf
    s$ = s$ + "b=b,8" + vbCrLf
    s$ = s$ + "c=c,(" + vbCrLf
    s$ = s$ + "d=d" + vbCrLf
    s$ = s$ + "e=e,3" + vbCrLf
    s$ = s$ + "f=f" + vbCrLf
    s$ = s$ + "g=g,9" + vbCrLf
    s$ = s$ + "h=h,#" + vbCrLf
    s$ = s$ + "i=i,1,|,!" + vbCrLf
    s$ = s$ + "j=j" + vbCrLf
    s$ = s$ + "k=k" + vbCrLf
    s$ = s$ + "l=l,7,1,|,!" + vbCrLf
    s$ = s$ + "m=m" + vbCrLf
    s$ = s$ + "n=n" + vbCrLf
    s$ = s$ + "o=o,0" + vbCrLf
    s$ = s$ + "p=p" + vbCrLf
    s$ = s$ + "q=q" + vbCrLf
    s$ = s$ + "r=r" + vbCrLf
    s$ = s$ + "s=s,5,$" + vbCrLf
    s$ = s$ + "t=t,7,+" + vbCrLf
    s$ = s$ + "u=u" + vbCrLf
    s$ = s$ + "v=v" + vbCrLf
    s$ = s$ + "w=w" + vbCrLf
    s$ = s$ + "x=x" + vbCrLf
    s$ = s$ + "y=y" + vbCrLf
    s$ = s$ + "z=z,2" + vbCrLf
    
    defaultLeetSpeak$ = s$
End Function

Public Sub loadLeetspeak()
    Dim sarr$(), i&, ch%, s$
    s$ = regGet("leetspeak")
    sarr$ = Split(s$, vbCrLf)
    For i& = 0 To UBound(sarr$())
        s$ = sarr$(i&)
        If InStr(s$, "=") <> 0 Then
            ch% = Asc(LCase(Left(s$, 1)))
            If ch% >= Asc("a") And ch% <= Asc("z") Then
                'it's a letter
                SALPH$(ch% - Asc("a")) = Right(s$, Len(s$) - InStr(s$, "="))
                SCOUNT%(ch% - Asc("a")) = CountString&(SALPH$(ch% - Asc("a")), ",") + 1
            End If
        End If
    Next i&
End Sub

Private Sub Form_Load()
    Dim s$, i&, j&, f$, ff%, sarr$(), ch%
    
    B_CHANGE = False
    
    'load leetspeak dictionary
    s$ = regGet("leetspeak")
    If s$ = "" Then
        s$ = defaultLeetSpeak$()
        regSet "leetspeak", s$
    End If
    
    loadLeetspeak
    
    If regGet("charset") = "" Then
        regSet "charset", defaultCharset()
    End If
    loadCharsets
    
    'load file split pref's
    s$ = regGet("split")
    Select Case s$
    Case "50000"
        mnuFileSplit50000.Checked = True
    Case "100000"
        mnuFileSplit100000.Checked = True
    Case "1000000"
        'preset
        mnuFileSplit1000000.Checked = True
    Case ""
        'nothing
        mnuFileSplitNever.Checked = True
    Case Else
        'custom
        mnuFileSplitCustom.Checked = True
        mnuFileSplitCustom.Caption = "every [" + Format(CLng(s$), "###,###") + "] words"
        If LANGUAGE$ <> "english" Then Change_Language
    End Select
    
    'load filter pref's
    s$ = regGet("len_max")
    Select Case s$
    Case "", "0"
        mnuFilterLenMax.Tag = "0"
    Case Else
        mnuFilterLenMax.Caption = "maximum length: " + s$
        mnuFilterLenMax.Checked = True
        mnuFilterLenMax.Tag = s$
    End Select
    s$ = regGet("len_min")
    Select Case s$
    Case "", "0"
        mnuFilterLenMin.Tag = "0"
    Case Else
        mnuFilterLenMin.Caption = "minimum length: " + s$
        mnuFilterLenMin.Checked = True
        mnuFilterLenMin.Tag = s$
    End Select
    s$ = regGet("text_left")
    If s$ <> "" Then
        mnuFilterTextLeft.Caption = "text to the left of [" + s$ + "]"
        mnuFilterTextLeft.Checked = True
        mnuFilterTextLeft.Tag = s$
    End If
    s$ = regGet("text_right")
    If s$ <> "" Then
        mnuFilterTextRight.Caption = "text to the right of [" + s$ + "]"
        mnuFilterTextRight.Checked = True
        mnuFilterTextRight.Tag = s$
    End If
    If regGet("hex") = "-1" Then
        mnuFilterHex.Checked = True
    End If
    If regGet("foreign") = "-1" Then
        mnuFilterForeign.Checked = True
    End If
    
    'load case pref's
    Select Case regGet("case")
    Case "lower"
        mnuCaseLower.Checked = True
    Case "upper"
        mnuCaseUpper.Checked = True
    End Select
    If regGet("case_first") = "-1" Then
        mnuCaseFirst.Checked = True
    End If
    If regGet("case_leet") = "-1" Then
        mnuCaseLeet.Checked = True
    End If
    If regGet("case_everyother") = "-1" Then
        mnuCaseEveryother.Checked = True
    End If
    
    'load append pref's
    s$ = regGet("append_pre")
    If s$ <> "" And Len(Dir(s$)) <> 0 Then
        mnuAppendPre.Caption = "append prefix to list [" + LCase(GetFileName$(s$)) + "]"
        mnuAppendPre.Checked = True
        mnuAppendPre.Tag = s$
    End If
    s$ = regGet("append_post")
    If s$ <> "" And Len(Dir(s$)) <> 0 Then
        mnuAppendPost.Caption = "append postfix to list [" + LCase(GetFileName$(s$)) + "]"
        mnuAppendPost.Checked = True
        mnuAppendPost.Tag = s$
    End If
    
    'filetype
    If regGet("filetype") = "unix" Then
        mnuFileTypeUnix.Checked = True
        mnuFileTypeUnix.Enabled = False
    Else
        mnuFileTypeWin.Checked = True
        mnuFileTypeWin.Enabled = False
    End If
    
    'generator : date : separator pref's
    s$ = regGet("sep")
    If s$ <> "" Then
        mnuGenDateSep.Caption = "separator: " + s$
        mnuGenDateSep.Tag = s$
    End If
    mnuGenDate(0).Caption = "mm" + s$ + "dd" + s$ + "yy"
    mnuGenDate(1).Caption = "mm" + s$ + "dd" + s$ + "yyyy"
    mnuGenDate(2).Caption = "dd" + s$ + "mm" + s$ + "yy"
    mnuGenDate(3).Caption = "dd" + s$ + "mm" + s$ + "yyyy"
    mnuGenDate(4).Caption = "mmm" + s$ + "dd" + s$ + "yy"
    mnuGenDate(5).Caption = "mmm" + s$ + "dd" + s$ + "yyyy"
    mnuGenDate(6).Caption = "dd" + s$ + "mmm" + s$ + "yy"
    mnuGenDate(7).Caption = "dd" + s$ + "mmm" + s$ + "yyyy"
    
    'generator : phone number : separator pref's
    s$ = regGet("sep_phone")
    mnuGenPhoneSep.Tag = s$
    mnuGenPhoneArea.Caption = "[areacode]" + s$ + "[prefix]" + s$ + "####"
    mnuGenPhoneNoarea.Caption = "[prefix]" + s$ + "####"
    mnuGenPhoneSep.Caption = "separator: " + IIf(s$ = "", "[none]", "")
    
    'generator : analyze pref's
    s$ = regGet("min_prefix")
    If s$ = "" Then
        s$ = "4"
    End If
    mnuGenAnalMinPre.Tag = s$
    mnuGenAnalMinPre.Caption = "min. length of prefix: [" + s$ + "]"
    s$ = regGet("min_postfix")
    If s$ = "" Then
        s$ = "3"
    End If
    mnuGenAnalMinPost.Tag = s$
    mnuGenAnalMinPost.Caption = "min. length of postfix: [" + s$ + "]"
    s$ = regGet("pattern_count")
    If s$ = "" Then
        s$ = "5"
    End If
    mnuGenAnalCount.Tag = s$
    mnuGenAnalCount.Caption = "minimum count: [" + s$ + "]"
    If regGet("pattern_case") = "0" Then
        mnuGenAnalCase.Checked = False
    End If
    
    'form location on screen
    s$ = regGet("x")
    If IsNumeric(s$) = True Then
        Me.Left = CLng(s$)
    Else
        Me.Left = (Screen.Width - Me.Width) / 2
    End If
    
    s$ = regGet("y")
    If IsNumeric(s$) = True Then
        Me.Top = CLng(s$)
    Else
        Me.Top = (Screen.Height - Me.Height) / 2
    End If
    
    If Me.Left > Screen.Width Then Me.Left = 0
    If Me.Top > Screen.Height Then Me.Top = 0
    
    'first load gets welcome screen
    If regGet("firstload?") = "" Then
        'hasn't been loaded before
        regSet "firstload?", "Not_anymore"
        picWelcome.Visible = True
    End If
    
    Me.Height = 5115
    Me.Show
    
    'set width of columnheader on list to width of list MINUS system-defined scrollbar width
    lst.ColumnHeaders(1).Width = lst.Width - (GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY)
    
    'set form on top of other windows
    Call setwindowpos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
    
    Erase sarr$()
    
    s$ = regGet("lang")
    If s$ = "" Then s$ = "english"
    LANGUAGE$ = s$
    Change_Language
    
    If regGet("resume") <> "" Then
        RESUME_STRING$ = regGet("resume")
        Dim x As Integer
        x = 0
        If IsNumeric(regGet("resume_index")) = True Then
            x = CInt(regGet("resume_index"))
        End If
        
        regSet "resume", ""
        regSet "resume_index", ""
        
        If frmMsg.MsgBocks("L517 has saved your position in the last list you were generating (" + Me.mnuGenStringX(x).Tag + ")." + vbCrLf + vbCrLf + "Do you want to resume generating this list?", vbQuestion + vbYesNo, "L517") = vbYes Then
            mnuGenStringX_Click x
        Else
            RESUME_STRING$ = ""
        End If
    End If
    
End Sub

Private Sub Form_LostFocus()
    'everytime the form loses focus, put it back ontop
    'formontop loses its ability after a while for some reason
    Call setwindowpos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_MouseDown Button, Shift, x, Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_MouseMove Button, Shift, x, Y
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_MouseUp Button, Shift, x, Y
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 0 And Me.Caption <> "" Then
        Me.Caption = ""
        Me.Width = 3255
        Me.Height = 5115
        Me.Refresh
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim s$
    Cancel = 0
    If B_CHANGE = True And lst.ListItems().count > 0 Then
        Select Case LANGUAGE$
        Case "english"
            s$ = "the list has been edited since it was last saved. do you want to save it now?"
        Case "french"
            s$ = "la liste a été modifié depuis sa dernière sauvegarde. Voulez-vous l'enregistrer maintenant?"
        Case "german"
            s$ = "Die Liste ist so abgefasst, seit es zuletzt gespeichert wurde. wollen Sie es jetzt speichern?"
        Case "spanish"
            s$ = "la lista se ha modificado desde que se guardó por última vez. ¿quieres salvar ahora?"
        End Select
        Select Case frmMsg.MsgBocks(s$, vbQuestion + vbYesNoCancel, "L517")
        Case vbYes
            mnuFileSave_Click
        Case vbNo
            
        Case vbCancel
            Cancel = 1
            Exit Sub
        End Select
    End If
    
    regSet "x", CStr(Me.Left)
    regSet "y", CStr(Me.Top)
    
    DoEvents
    Me.Visible = False
    DoEvents
    
    lst.ListItems().Clear
    
    s$ = App.Path
    If Right(s$, 1) <> "\" Then s$ = s$ + "\"
    
    On Error Resume Next
    Kill s$ + "readme.txt"
    
    End
End Sub

Private Sub lblAppend2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblAppend2.ForeColor = &H808080
    Me.PopupMenu Me.mnuAppend, , lblAppend2.Left - 400, lblAppend2.Top + lblAppend2.Height
    lblAppend2.ForeColor = vbWhite
End Sub

Private Sub lblAppend2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub lblCase2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblCase2.ForeColor = &H808080
    Me.PopupMenu Me.mnuCase, , lblCase2.Left - 400, lblCase2.Top + lblCase2.Height
    lblCase2.ForeColor = vbWhite
End Sub

Private Sub lblCancel_Click()
    UpdateCaption
    lblCancel.Visible = False
End Sub

Private Sub lblCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblCancel.Visible = False
End Sub

Private Sub lblCase2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub lblDismiss_Click()
    picWelcome.Visible = False
End Sub

Private Sub lblEdit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblEdit.ForeColor = &H808080
    Me.PopupMenu Me.mnuList, , lblEdit.Left - 400, lblEdit.Top + lblEdit.Height
    lblEdit.ForeColor = vbWhite
End Sub
Private Sub lblEdit_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub lblFile_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblFile.ForeColor = &H808080
    Me.PopupMenu mnuFile, , lblFile.Left - 500, lblFile.Top + lblFile.Height
    lblFile.ForeColor = vbWhite
End Sub
Private Sub lblFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub lblFilter2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblFilter2.ForeColor = &H808080
    Me.PopupMenu Me.mnuFilter, , lblFilter2.Left - 400, lblFilter2.Top + lblFilter2.Height
    lblFilter2.ForeColor = vbWhite
End Sub
Private Sub lblFilter2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub lblGen_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblGen.ForeColor = &H808080
    Me.PopupMenu mnuGen, , lblGen.Left - 450, lblGen.Top + lblGen.Height
    lblGen.ForeColor = vbWhite
End Sub
Private Sub lblGen_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblHelp.ForeColor = &H808080
    Me.PopupMenu mnuHelp, , lblHelp.Left - 300, lblHelp.Top + lblHelp.Height
    lblHelp.ForeColor = vbWhite
End Sub
Private Sub lblHelp_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub lblExit_Click()
    Unload Me
End Sub
Private Sub lblexit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblExit.ForeColor = &H808080
    Call setcapture(Me.hWnd)
End Sub
Private Sub lblexit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblExit.ForeColor = vbWhite
    Call releasecapture
End Sub
Private Sub lblExit_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub lblMin_Click()
    Me.WindowState = vbMinimized
    Me.Caption = "L517"
End Sub
Private Sub lblMin_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblMin.ForeColor = &H808080
    Call setcapture(Me.hWnd)
End Sub
Private Sub lblMin_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblMin.ForeColor = vbWhite
    Call releasecapture
End Sub
Private Sub lblMin_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_OLEDragDrop Data, Effect, Button, Shift, x, Y
End Sub

Private Sub lblSite_Click()
    mnuHelpSite_Click
End Sub


Private Sub lblStat_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_MouseDown Button, Shift, x, lblStat.Top + Y
End Sub
Private Sub lblStat_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_MouseMove Button, Shift, x, lblStat.Top + Y
End Sub
Private Sub lblStat_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblTitle_MouseUp Button, Shift, x, lblStat.Top + Y
End Sub

Private Sub lblTitle_DblClick()
    If lblTitle.Tag = "" Then
        lblTitle.Tag = "X"
        Do
            DoEvents
            Me.Height = Me.Height - 50
        Loop Until Me.Height < 280
        Me.Height = 280
    Else
        lblTitle.Tag = ""
        Do
            DoEvents
            Me.Height = Me.Height + 50
        Loop Until Me.Height > 5115
        Me.Height = 5115
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If picWelcome.Visible = True Then picWelcome.Visible = False
    If Button = 1 Then
        C_MOVER = True
        C_FORMX = x
        C_FORMY = Y
        Call setcapture(Me.hWnd)
    End If
End Sub
Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lmsg&, lresult&, pap As pointapi
    
    If C_MOVER = True Then
        If lblTitle.ForeColor <> &H808080 Then
            lblTitle.ForeColor = &H808080
            lblMin.ForeColor = &H808080
            lblExit.ForeColor = &H808080
        End If
        
        Call getcursorpos(pap)
        If (pap.x * Screen.TwipsPerPixelX) - C_FORMX < 250 Then
            Me.Left = 0
        ElseIf (pap.x * Screen.TwipsPerPixelX) - C_FORMX > (Screen.Width - Me.Width) - 250 Then
            Me.Left = (Screen.Width) - Me.Width
        Else
            Me.Left = (pap.x * Screen.TwipsPerPixelX) - C_FORMX
        End If
        
        Select Case (pap.Y * Screen.TwipsPerPixelY) - C_FORMY
        Case Is < 250
            Me.Top = -25
        Case Is > (Screen.Height - Me.Height) - 250
            Me.Top = Screen.Height - Me.Height
        Case Else
            Me.Top = (pap.Y * Screen.TwipsPerPixelY) - C_FORMY
        End Select
    End If
End Sub
Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If C_MOVER = True Then
        C_MOVER = False
        lblTitle.ForeColor = vbWhite
        lblMin.ForeColor = vbWhite
        lblExit.ForeColor = vbWhite
        Call releasecapture
    End If
End Sub

Private Sub lblTitle_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim i&, s$, sfiles$
    
    'data.formats:
    '  1 = text
    ' 15 = files!
    
    If Data.GetFormat(15) = True Then
        lst.Visible = False
        For i& = 1 To Data.Files().count
            s$ = Data.Files(i&)
            If LoadList(s$, False) = False Then Exit For
            UpdateCaption
        Next i&
        lst.Visible = True
    ElseIf Data.GetFormat(1) = True Then
        stat "reading dragged text..."
        ParseTextBlock Data.GetData(1)
        '
        UpdateCaption
        stat ""
    Else
        'unknown
        On Error Resume Next
        For i& = 1 To 500
            If Data.GetFormat(i&) = True Then
                sfiles$ = Data.GetData(i&)
                Exit For
            End If
        Next i&
        If sfiles$ = "" Then sfiles$ = "[null]"
        
        On Error GoTo 0
        frmMsg.MsgBocks "unsupported data format." + vbCrLf + "number: (" + CStr(i&) + ")" + vbCrLf + vbCrLf + "data: " + sfiles$, vbExclamation + vbOKOnly
    End If
End Sub

Private Sub lblWelcome_Click()
    lblDismiss_Click
End Sub

Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub lst_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        lblTitle_MouseDown Button, Shift, lst.Left + x, lst.Top + Y
    ElseIf Button = 2 Then
        Me.PopupMenu mnuList
    End If
End Sub
Private Sub lst_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        lblTitle_MouseMove Button, Shift, lst.Left + x, lst.Top + Y
    End If
End Sub
Private Sub lst_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        lblTitle_MouseUp Button, Shift, lst.Left + x, lst.Top + Y
    End If
End Sub

Private Sub lst_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim i&, s$, sfiles$
    
    '1 = text
    '15 = files!
    If Data.GetFormat(15) = True Then
        lst.Visible = False
        For i& = 1 To Data.Files().count
            s$ = Data.Files(i&)
            If LoadList(s$, False) = False Then Exit For
        Next i&
        lst.Visible = True
    ElseIf Data.GetFormat(1) = True Then
        
        stat "reading dragged text..."
        ParseTextBlock Data.GetData(1)
        UpdateCaption
        stat ""
    Else
        On Error Resume Next
        For i& = 1 To 500
            If Data.GetFormat(i&) = True Then
                sfiles$ = Data.GetData(i&)
                Exit For
            End If
        Next i&
        If sfiles$ = "" Then sfiles$ = "[null]"
        
        On Error GoTo 0
        frmMsg.MsgBocks "unsupported data format." + vbCrLf + "number: (" + CStr(i&) + ")" + vbCrLf + vbCrLf + "data: " + sfiles$, vbExclamation + vbOKOnly
    End If
End Sub

Private Sub mnuCaseLeetEdit_Click()
    frmLeetspeak.Show vbModal, Me
    loadLeetspeak
End Sub

Private Sub mnuGenPhoneCustom_Click()
    Dim s$, count&, ff%, i&, imax&, sresult$, j&, current&, char$, snum$
    
    s$ = frmMsg.InputBocks("Enter custom number string to generate, with X's where you want numbers to go:" + vbCrLf + vbCrLf + "ex: 555-5XX-XXXX")
    If s$ = "" Then Exit Sub
    s$ = UCase(s$)
    
    count& = 0
    For i& = 1 To Len(s$)
        If Mid(s$, i&, 1) = "X" Then
            count& = count& + 1
        End If
    Next i&
    
    lst.ListItems().Clear
    lst.Visible = False
    lblCancel.Visible = True
    DoEvents
    
    stat "generating phone numbers..."
    ff% = FreeFile%
    Open App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + s$ + ".txt" For Binary Access Write As #ff%
        
        imax& = 10 ^ count&
        ' loop through every number
        For i& = 0 To imax& - 1
            If i Mod 2500 = 0 Then
                DoEvents
                prog CDbl(i& + 1) / CDbl(imax&)
                If lblCancel.Visible = False Then
                    MsgBox "CANCELLED WTF"
                    GoTo cancelled
                End If
            End If
            snum$ = CStr(i&)
            ' add buffer of 0's to set correct length
            Do While Len(snum$) < count&
                snum$ = "0" + snum$
            Loop
            
            'build the current number
            sresult$ = ""
            current& = 0
            For j& = 1 To Len(s$) ' loop through entire custom string
                char$ = Mid(s$, j&, 1)
                If char$ = "X" Then ' if we're at an X, insert the current number we're at
                    current& = current& + 1
                    sresult$ = sresult$ + Mid(snum$, current&, 1)
                Else
                    sresult$ = sresult$ + char$ ' if we're not at an X, it's part of the custom string
                End If
            Next j&
            Put #ff%, , CStr(sresult$ + IIf(mnuFileTypeUnix.Checked, Chr(10), vbCrLf))
            'lst.ListItems().Add , , sresult$ 'add it to the list
        Next i&
cancelled:
        If i& = 0 Then i& = imax&
        lblCancel.Visible = False
        prog 0
    Close #ff%
    
    lst.Visible = True
    stat "#'s saved to '" + s$ + ".txt'"
End Sub

Private Sub mnuGenStringEdit_Click()
    frmCharset.Show vbModal, Me
    
    loadCharsets
End Sub

Private Sub loadCharsets()
    Dim s$, i&, j&, sarr$()
    
    For i& = 1 To mnuGenStringX.UBound
        Unload mnuGenStringX(i&)
    Next i&
    
    s$ = regGet("charset")
    sarr$() = Split(s$, vbCrLf)
    j& = 0
    For i& = 0 To UBound(sarr$())
        s$ = sarr$(i&)
        If Left(s$, 1) <> "#" And InStr(s$, "[") <> 0 And InStr(s$, "]") <> 0 And InStr(s$, "=") <> 0 Then
            'isn't commented out by #, contains [ ] and =
            
            'If Mid(s$, InStr(s$, "[") + 1, InStr(s$, "]") - InStr(s$, "[") - 1) <> "" Then
            If Mid(s$, InStr(s$, "[") + 1, Len(s$) - InStr(s$, "[") - 1) <> "" Then
                If j& <> 0 Then
                    Load mnuGenStringX(j&)
                End If
                With mnuGenStringX(j&)
                    .Caption = Trim(Left(s$, InStr(s$, "=") - 1))
                    .Tag = Mid(s$, InStr(s$, "[") + 1, Len(s$) - InStr(s$, "[") - 1)
                End With
            End If
            j& = j& + 1
        ElseIf s$ = "" And i& <> UBound(sarr$()) Then
            If s$ = "" Then
                If j& <> 0 Then
                    Load mnuGenStringX(j&)
                End If
                mnuGenStringX(j&).Caption = "-"
                mnuGenStringX(j&).Tag = ""
                j& = j& + 1
            End If
        End If
    Next i&
End Sub

Private Sub mnuHelpAbout_Click()
    Dim s$, version$, day$
    
    version$ = "0.994"
    day$ = "24feb2012"
    
    Select Case LANGUAGE$
    Case "english"
        s$ = "os:" + vbTab + vbTab + "windows 98/xp/vista/seven" + vbCrLf
        s$ = s$ + "version:" + vbTab + vbTab + version$ + vbCrLf
        s$ = s$ + "compiled:" + vbTab + vbTab + day$ + vbCrLf
        s$ = s$ + "author:" + vbTab + vbTab + "derv" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "may this program guide you in building the perfect wordlist." + vbCrLf
    Case "french"
        s$ = "Système d'exploitation:" + vbTab + vbTab + "Windows 98/xp/vista/seven" + vbCrLf
        s$ = s$ + "version:" + vbTab + vbTab + version$ + vbCrLf
        s$ = s$ + "compilé:" + vbTab + vbTab + day$ + vbCrLf
        s$ = s$ + "Auteur:" + vbTab + vbTab + "derv" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "mai ce programme vous guide dans la construction du dictionnaire parfait." + vbCrLf
    Case "german"
        s$ = "Betriebssystem:" + vbTab + vbTab + "windows 98/xp/vista/seven" + vbCrLf
        s$ = s$ + "version:" + vbTab + vbTab + version$ + vbCrLf
        s$ = s$ + "zusammengestellt:" + vbTab + vbTab + day$ + vbCrLf
        s$ = s$ + "Autor:" + vbTab + vbTab + "derv" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "dieses Programm führt Sie beim Aufbau der perfekte Wortliste." + vbCrLf
    Case "spanish"
        s$ = "sistema operativo:" + vbTab + vbTab + "windows 98/xp/vista/seven" + vbCrLf
        s$ = s$ + "versión:" + vbTab + vbTab + version$ + vbCrLf
        s$ = s$ + "compilado:" + vbTab + vbTab + day$ + vbCrLf
        s$ = s$ + "autor:" + vbTab + vbTab + "derv" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "de este programa puede guiarlo en la construcción de la lista de palabras perfecto." + vbCrLf
    End Select
    frmMsg.MsgBocks s$, vbInformation + vbOKOnly, "L517"
End Sub

Private Sub mnuAppendPostAlpha_Click()
    Dim i&, sarr$(), j&
    
    If lst.ListItems().count = 0 Then
        Select Case LANGUAGE$
        Case "english"
            frmMsg.MsgBocks "you need items in the list before you can append words to them!", vbInformation
        Case "french"
            frmMsg.MsgBocks "vous avez besoin d'éléments de la liste avant de pouvoir ajouter des mots à lui!", vbInformation
        Case "german"
            frmMsg.MsgBocks "Sie müssen Elemente in der Liste, bevor Sie mit Worten, um ihn hinzuzufügen!", vbInformation
        Case "spanish"
            frmMsg.MsgBocks "que usted necesita elementos de la lista antes de que usted puede agregar palabras a él!", vbInformation
        End Select
        Exit Sub
    End If
    
    stat "reading list..."
    For i& = 1 To lst.ListItems().count
        If i& Mod 250 Then DoEvents
        ReDim Preserve sarr$(i& - 1)
        sarr$(i& - 1) = lst.ListItems(i&).Text
    Next i&
    
    stat "adding items..."
    
    lst.Visible = False
    DoEvents
    
    For i& = 0 To UBound(sarr$())
        If i& Mod 10 Then DoEvents
        For j& = 0 To 25
            lst.ListItems().add , , sarr$(i&) + Chr(Asc("a") + j&)
        Next j&
    Next i&
    
    Erase sarr$()
    
    stat ""
    UpdateCaption
    DoEvents
    lst.Visible = True
End Sub

Private Sub mnuAppendPostCustom_Click()
    Dim i&, j&, sfile$, sa$(), sadd$(), ff%, s$
    
    If lst.ListItems().count = 0 Then
        Select Case LANGUAGE$
        Case "english"
            frmMsg.MsgBocks "you need items in the list before you can append words to them!", vbInformation
        Case "french"
            frmMsg.MsgBocks "vous avez besoin d'éléments de la liste avant de pouvoir ajouter des mots à lui!", vbInformation
        Case "german"
            frmMsg.MsgBocks "Sie müssen Elemente in der Liste, bevor Sie mit Worten, um ihn hinzuzufügen!", vbInformation
        Case "spanish"
            frmMsg.MsgBocks "que usted necesita elementos de la lista antes de que usted puede agregar palabras a él!", vbInformation
        End Select
        Exit Sub
    End If
    
    sfile$ = getopen$()
    If Len(Dir(sfile$)) = 0 Or sfile$ = "" Then Exit Sub
    
    stat "queueing postfixes..."
    DoEvents
    
    'read first line of file to see if it's linux or unix
    ff% = FreeFile
    Open sfile$ For Input As #ff%
        Input #ff%, s$
    Close #ff%
    
    If InStr(s$, Chr(10)) <> 0 Then
        'unix
        sadd$() = Split(s$, Chr(10))
    Else
        'windows
        i& = 0
        Open sfile$ For Input As #ff%
            While Not EOF(ff%)
                If i& Mod 250 = 0 Then DoEvents
                Input #ff%, s$
                ReDim Preserve sadd$(i&)
                sadd$(i&) = s$
                i& = i& + 1
            Wend
        Close #ff%
    End If
    
    ReDim sa$(lst.ListItems().count - 1)
    
    For i& = 1 To lst.ListItems().count
        sa$(i& - 1) = lst.ListItems(i&).Text
    Next i&
    
    stat "adding items..."
    DoEvents
    For i& = 0 To UBound(sa$())
        If i& Mod 10 = 0 Then
            DoEvents
            UpdateCaption
        End If
        
        For j& = 0 To UBound(sadd$())
            lst.ListItems().add , , sa$(i&) + sadd$(j&)
        Next j&
    Next i&
    
    B_CHANGE = True
    
    Erase sa$()
    Erase sadd$()
    UpdateCaption
    stat ""
End Sub

Private Sub mnuAppendPostDefault_Click()
    Dim s$, sadd$(), sarr$(), i&, j&
    
    s$ = "!|#|$|%|&me|)|*|.|@|]|~|0|00|01|0123|02|03|04|05|06|07|08|09|1|10|11|12|123|1234|12345|123456|13|14|15|159|16|17|18|187|19|2|20|2000|2001|2002|2003|2004|2005|2006|2007|"
    s$ = s$ + "2008|2009|2010|2011|2012|2013|21|22|23|24|25|26|27|28|29|2u|3|30|31|316|357|4|42|420|456|4eva|4ever|4life|4me|5|6|666|6789|69|7|753|777|789|8|9|90|91|92|93|94|95|951|96|97|98|"
    s$ = s$ + "99|abcd|asd|asdf|boy|christ|cxz|doll|dsa|eagle|ewq|face|fuck|god|iloveyou|itsme|jesus|loser|lover|qwe|qwer|s|u2|xx|xxx|zxc|zxcv|"
    sadd$() = Split(s$, "|")
    s$ = ""
    
    If lst.ListItems().count > 0 Then
        
        stat "reading list..."
        For i& = 1 To lst.ListItems().count
            If i& Mod 250 Then DoEvents
            ReDim Preserve sarr$(i& - 1)
            sarr$(i& - 1) = lst.ListItems(i&).Text
        Next i&
    Else
        ReDim sarr$(0)
        sarr$(0) = ""
    End If
    
    stat "adding items..."
    
    lst.Visible = False
    DoEvents
    For i& = 0 To UBound(sarr$())
        If i Mod 100 = 0 Then
            DoEvents
            prog (i& + 1) / UBound(sarr$())
        End If
        For j& = 0 To UBound(sadd$())
            s$ = sarr$(i&) + sadd$(j&)
            If FilterCheck(s$) = True Then
                lst.ListItems().add , , s$
            End If
        Next j&
    Next i&
    
    B_CHANGE = True
    
    Erase sarr$()
    Erase sadd$()
    
    stat ""
    UpdateCaption
    DoEvents
    lst.Visible = True
    
    Erase sadd$()
    Erase sarr$()
End Sub

Private Sub mnuAppendPostNum_Click()
    Dim i&, sarr$(), j&
    
    If lst.ListItems().count = 0 Then
        Select Case LANGUAGE$
        Case "english"
            frmMsg.MsgBocks "you need items in the list before you can append words to them!", vbInformation
        Case "french"
            frmMsg.MsgBocks "vous avez besoin d'éléments de la liste avant de pouvoir ajouter des mots à lui!", vbInformation
        Case "german"
            frmMsg.MsgBocks "Sie müssen Elemente in der Liste, bevor Sie mit Worten, um ihn hinzuzufügen!", vbInformation
        Case "spanish"
            frmMsg.MsgBocks "que usted necesita elementos de la lista antes de que usted puede agregar palabras a él!", vbInformation
        End Select
        Exit Sub
    End If
    
    stat "reading list..."
    For i& = 1 To lst.ListItems().count
        If i& Mod 250 Then DoEvents
        ReDim Preserve sarr$(i& - 1)
        sarr$(i& - 1) = lst.ListItems(i&).Text
    Next i&
    
    stat "adding items..."
    
    lst.Visible = False
    DoEvents
    
    For i& = 0 To UBound(sarr$())
        If i& Mod 25 Then DoEvents
        For j& = 0 To 9
            lst.ListItems().add , , sarr$(i&) + CStr(j&)
        Next j&
    Next i&
    
    B_CHANGE = True
    
    Erase sarr$()
    
    stat ""
    UpdateCaption
    DoEvents
    lst.Visible = True
End Sub

Private Sub mnuAppendPreAlpha_Click()
    Dim i&, sarr$(), j&
    
    If lst.ListItems().count = 0 Then
        Select Case LANGUAGE$
        Case "english"
            frmMsg.MsgBocks "you need items in the list before you can append words to them!", vbInformation
        Case "french"
            frmMsg.MsgBocks "vous avez besoin d'éléments de la liste avant de pouvoir ajouter des mots à lui!", vbInformation
        Case "german"
            frmMsg.MsgBocks "Sie müssen Elemente in der Liste, bevor Sie mit Worten, um ihn hinzuzufügen!", vbInformation
        Case "spanish"
            frmMsg.MsgBocks "que usted necesita elementos de la lista antes de que usted puede agregar palabras a él!", vbInformation
        End Select
        Exit Sub
    End If
    
    stat "reading list..."
    For i& = 1 To lst.ListItems().count
        If i& Mod 250 Then DoEvents
        ReDim Preserve sarr$(i& - 1)
        sarr$(i& - 1) = lst.ListItems(i&).Text
    Next i&
    
    stat "adding items..."
    
    lst.Visible = False
    DoEvents
    
    For i& = 0 To UBound(sarr$())
        If i& Mod 10 Then DoEvents
        For j& = 0 To 25
            lst.ListItems().add , , Chr(Asc("a") + j&) + sarr$(i&)
        Next j&
    Next i&
    
    B_CHANGE = True
    
    Erase sarr$()
    
    stat ""
    UpdateCaption
    DoEvents
    lst.Visible = True
End Sub

Private Sub mnuAppendPreCustom_Click()
    Dim i&, j&, sfile$, sa$(), sadd$(), ff%, s$
    
    If lst.ListItems().count = 0 Then
        Select Case LANGUAGE$
        Case "english"
            frmMsg.MsgBocks "you need items in the list before you can append words to them!", vbInformation
        Case "french"
            frmMsg.MsgBocks "vous avez besoin d'éléments de la liste avant de pouvoir ajouter des mots à lui!", vbInformation
        Case "german"
            frmMsg.MsgBocks "Sie müssen Elemente in der Liste, bevor Sie mit Worten, um ihn hinzuzufügen!", vbInformation
        Case "spanish"
            frmMsg.MsgBocks "que usted necesita elementos de la lista antes de que usted puede agregar palabras a él!", vbInformation
        End Select
        Exit Sub
    End If
    
    sfile$ = getopen$()
    If Len(Dir(sfile$)) = 0 Or sfile$ = "" Then Exit Sub
    
    stat "queueing prefixes..."
    DoEvents
    
    'read first line of file to see if it's linux or unix
    ff% = FreeFile
    Open sfile$ For Input As #ff%
        Input #ff%, s$
    Close #ff%
    
    If InStr(s$, Chr(10)) <> 0 Then
        'unix
        sadd$() = Split(s$, Chr(10))
    Else
        'windows
        i& = 0
        Open sfile$ For Input As #ff%
            While Not EOF(ff%)
                If i& Mod 250 = 0 Then DoEvents
                Input #ff%, s$
                ReDim Preserve sadd$(i&)
                sadd$(i&) = s$
                i& = i& + 1
            Wend
        Close #ff%
    End If
    
    ReDim sa$(lst.ListItems().count - 1)
    
    For i& = 1 To lst.ListItems().count
        sa$(i& - 1) = lst.ListItems(i&).Text
    Next i&
    
    stat "adding items..."
    DoEvents
    For i& = 0 To UBound(sa$())
        If i& Mod 10 = 0 Then
            DoEvents
            UpdateCaption
        End If
        
        For j& = 0 To UBound(sadd$())
            lst.ListItems().add , , sadd$(j&) + sa$(i&)
        Next j&
    Next i&
    
    B_CHANGE = True
    
    Erase sa$()
    Erase sadd$()
    
    UpdateCaption
    stat ""
End Sub

Private Sub mnuAppendPreDefault_Click()
    Dim s$, sadd$(), sarr$(), i&, j&
    
    s$ = "!|#|#1|$|(|*|.|[|~|0|00|0123|1|1234|12345|123456|1234567|12345678|2|3|4|420|5|6|69|7|8|9|a|abc|abcd|abcde|abcdefg|acbdef|adam|adidas|alex|alexis|alyssa|amanda|amber|andrew|angel|animals|anthony|apple|april|arsenal|asdf|ashley|"
    s$ = s$ + "ass|asshole|august|austin|awesome|b|babgirl|baby|babygurl|badboy|bailey|ballen|baller|ballin|balls|banana|barbie|barney|baseball|basketball|bastard|batman|bball|bday|beans|beast|beautiful|bebe|bellabeach|belle|bitch|blue|bobby|boobies|booger|boomer|bowwow|brad|brandon|brian|britney|britt|brittany|broken|brook|brooke|bubble|bubbles|buddy|bulldog|"
    s$ = s$ + "bunny|buster|butt|butterfly|c|cali|cancer|candy|carlos|carmen|carolina|casper|cat|charlie|cheer|cheese|chelsea|cherry|chicken|chocolate|chris|christ|class|classof|coco|cody|computer|confused|cookie|cool|cowboys|crazy|cuddles|cupcake|cutie|daddy|daisy|dakota|dallas|dance|dancer|dani|daniel|dark|dave|david|dead|death|december|demon|"
    s$ = s$ + "devil|diamond|dickhead|dog|dolphin|dolphins|donkey|dragon|dragons|dream|dreams|duck|duckie|ducky|dude|duke|eagle|edward|eeyore|element|elephant|elizabeth|emily|eminem|emma|england|eric|faggot|faith|family|fire|fireman|football|forever|frankie|free|friends|fuck|fucker|fuckme|fucku|fuckyou|futbol|gangsta|george|ginger|god|gold|golf|good|"
    s$ = s$ + "goodbye|google|grandma|green|greenday|gunit|gymnast|hannah|happy|hardcore|harry|heart|heather|hell|hello|hockey|holla|hollister|hollywood|home|honda|honey|horse|hotdog|hottie|house|hunter|hunting|iam|ihateyou|ilike|ilove|ilovehim|iloveu|iloveyou|iluvu|inlove|internet|ipod|irish|irock|jack|jackass|jacob|jake|james|january|jasmine|jason|jazz|"
    s$ = s$ + "jennifer|jessica|jessie|jesus|jimmy|john|johnny|jojo|joker|jordan|josh|joshua|juice|juicy|july|june|justin|katie|kevin|kiki|killa|killer|king|kiss|kissme|kitten|kitty|kittycat|kool|lady|lakers|lemon|letmein|life|lil|lindsey|lipgloss|little|llama|lol|lollipop|london|loser|louise|love|lover|loveyou|lucky|lynn|magic|"
    s$ = s$ + "march|marie|mark|matt|matthew|may|mcfly|metallica|mexico|michael|michelle|mickey|mike|milli|missy|molly|mommy|money|monkey|muffin|music|mustang|mylove|natalie|nathan|newyork|nick|nicole|nigger|nike|nikki|nirvana|nothing|november|number|oliver|omfg|orange|oscar|panda|pantera|panthers|passw0rd|password|patrick|peanut|penguin|penis|people|pepper|"
    s$ = s$ + "pickle|piglet|pimp|pimpin|pink|pinky|pirates|playa|playboy|player|pokemon|poohbear|pookie|poop|poopie|popcorn|power|precious|pretty|princess|punk|purple|pussy|queen|qwerty|rachel|raiders|rainbow|rayray|red|redsox|retard|richard|robert|rock|rockstar|rose|roxy|ryan|sally|sammy|sassy|satan|scarface|school|scooter|scott|secret|september|sexi|"
    s$ = s$ + "sexy|shadow|shit|shorty|silver|simba|skate|skater|slayer|slipknot|smile|smith|snoopy|soccer|softball|sparkle|spiderman|spike|spongebob|star|starwars|sublime|sugar|summer|sunshine|super|superman|surf|sweet|tennis|the|thomas|tiger|tigers|tigger|timothy|tink|tinker|tommy|tony|track|travis|tree|trojan|turtle|tweety|tyler|vagina|vball|verizon|"
    s$ = s$ + "victor|volcom|warriors|water|weed|whatever|whore|wifey|wild|william|winter|wrestling|xx|xxx|yamaha|yankees|yellow|young|yourmom|zach|zero|zxcvbnm|"
    sadd$() = Split(s$, "|")
    s$ = ""
    
    If lst.ListItems().count > 0 Then
        
        stat "reading list..."
        For i& = 1 To lst.ListItems().count
            If i& Mod 250 Then DoEvents
            ReDim Preserve sarr$(i& - 1)
            sarr$(i& - 1) = lst.ListItems(i&).Text
        Next i&
    Else
        ReDim sarr$(0)
        sarr$(0) = ""
    End If
    
    stat "adding items..."
    
    lst.Visible = False
    DoEvents
    For i& = 0 To UBound(sarr$())
        If i Mod 100 = 0 Then DoEvents
        For j& = 0 To UBound(sadd$())
            s$ = sadd$(j&) + sarr$(i&)
            If FilterCheck(s$) = True Then
                lst.ListItems().add , , s$
            End If
        Next j&
    Next i&
    
    B_CHANGE = True
    
    Erase sarr$()
    Erase sadd$()
    
    stat ""
    UpdateCaption
    DoEvents
    lst.Visible = True
End Sub

Private Sub mnuAppendPreNum_Click()
    Dim i&, sarr$(), j&
    
    If lst.ListItems().count = 0 Then
        Select Case LANGUAGE$
        Case "english"
            frmMsg.MsgBocks "you need items in the list before you can append words to them!", vbInformation
        Case "french"
            frmMsg.MsgBocks "vous avez besoin d'éléments de la liste avant de pouvoir ajouter des mots à lui!", vbInformation
        Case "german"
            frmMsg.MsgBocks "Sie müssen Elemente in der Liste, bevor Sie mit Worten, um ihn hinzuzufügen!", vbInformation
        Case "spanish"
            frmMsg.MsgBocks "que usted necesita elementos de la lista antes de que usted puede agregar palabras a él!", vbInformation
        End Select
        Exit Sub
    End If
    
    stat "reading list..."
    For i& = 1 To lst.ListItems().count
        If i& Mod 250 Then DoEvents
        ReDim Preserve sarr$(i& - 1)
        sarr$(i& - 1) = lst.ListItems(i&).Text
    Next i&
    
    stat "adding items..."
    
    lst.Visible = False
    DoEvents
    
    For i& = 0 To UBound(sarr$())
        If i& Mod 25 Then DoEvents
        For j& = 0 To 9
            lst.ListItems().add , , CStr(j&) + sarr$(i&)
        Next j&
    Next i&
    
    B_CHANGE = True
    
    Erase sarr$()
    
    stat ""
    UpdateCaption
    DoEvents
    lst.Visible = True
End Sub

Private Sub mnuCaseEveryother_Click()
    Dim change%
    
    mnuCaseEveryother.Checked = Not (mnuCaseEveryother.Checked)
    regSet "case_everyother", CStr(CInt(mnuCaseEveryother.Checked))
End Sub

Private Sub mnuCaseFirst_Click()
    Dim change%
    
    mnuCaseFirst.Checked = Not (mnuCaseFirst.Checked)
    regSet "case_first", CStr(CInt(mnuCaseFirst.Checked))
End Sub

Private Sub mnuCaseLeet_Click()
    Dim change%
    
    mnuCaseLeet.Checked = Not (mnuCaseLeet.Checked)
    regSet "case_leet", CStr(CInt(mnuCaseLeet.Checked))
End Sub

Private Sub mnuCaseLower_Click()
    Dim change%
    
    change% = CInt(mnuCaseLower.Checked)
    
    mnuCaseLower.Checked = Not (mnuCaseLower.Checked)
    If mnuCaseLower.Checked = False Then
        regSet "case", "none"
    Else
        regSet "case", "lower"
    End If
    
    mnuCaseUpper.Checked = False
    
    If change% - CInt(mnuCaseLower.Checked) > 0 Then
        UpdateList
    End If
End Sub

Private Sub mnuCaseUpper_Click()
    Dim change%
    
    change% = CInt(mnuCaseUpper.Checked)
    
    mnuCaseUpper.Checked = Not (mnuCaseUpper.Checked)
    If mnuCaseUpper.Checked = False Then
        regSet "case", "none"
    Else
        regSet "case", "upper"
    End If
    
    mnuCaseLower.Checked = False
    
    If change% - CInt(mnuCaseUpper.Checked) > 0 Then
        'only update when turning case filtering ON
        UpdateList
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    Dim s$
    If B_CHANGE = True Then
        Select Case LANGUAGE$
        Case "english"
            s$ = "the current list has been changed; do you want to SAVE THE CHANGES?"
        Case "french"
            s$ = "La liste actuelle a été modifié; voulez-vous enregistrer les modifications?"
        Case "german"
            s$ = "die aktuelle Liste geändert wurde, möchten Sie die Änderungen speichern?"
        Case "spanish"
            s$ = "la lista actual ha cambiado, ¿quieres guardar los cambios?"
        End Select
        
        Select Case frmMsg.MsgBocks(s$, vbYesNoCancel + vbQuestion)
        Case vbYes
            mnuFileSave_Click
        Case vbNo
            'do nothing
        Case vbCancel
            Exit Sub
        End Select
        B_CHANGE = False
    End If
    
    If lst.ListItems().count > 0 Then
        lst.ListItems().Clear
        UpdateCaption
    End If
    
    stat "inactive"
End Sub

Private Sub mnuFileOpen_Click()
    Dim s$
    
    s$ = getopen$()
    If Len(Dir(s$)) <> 0 And s$ <> "" Then
        LoadList s$
    End If
End Sub

Private Sub mnuFileProfilesWeb_Click()
    mnuFilterLenMin.Caption = "minimum length: 4"
    mnuFilterLenMin.Checked = True
    mnuFilterLenMin.Tag = "4"
    regSet "len_min", "4"
    
    mnuFilterLenMax.Caption = "maximum length: 12"
    mnuFilterLenMax.Checked = True
    mnuFilterLenMax.Tag = "12"
    regSet "len_max", "12"
    
    If LANGUAGE$ <> "english" Then Change_Language
    UpdateList
End Sub

Private Sub mnuFileSave_Click()
    Dim s$, sl$
    
    s$ = getsave$()
    If s$ <> "" Then
        If InStr(Right(s$, 5), ".") = 0 Then
            s$ = s$ + ".txt"
        End If
        If Len(Dir(s$)) <> 0 Then
            Select Case LANGUAGE$
            Case "english"
                sl$ = "'" + s$ + "' already exists.  do you wish to overwrite this file?"
            Case "french"
                sl$ = "'" + s$ + "' existe déjà. Voulez-vous écraser ce fichier?"
            Case "german"
                sl$ = "'" + s$ + "' bereits vorhanden ist. möchten Sie diese Datei überschreiben?"
            Case "spanish"
                sl$ = "'" + s$ + "' ya existe. ¿Desea sobreescribir este archivo?"
            End Select
            If frmMsg.MsgBocks(sl$, vbYesNo + vbExclamation, "L517") = vbNo Then Exit Sub
            Do While Len(Dir(s$)) <> 0
                DoEvents
                Kill s$
            Loop
        End If
        
        stat "saving '" + GetFileName(s$) + "'"
        lst.Visible = False
        DoEvents
        
        SaveList s$
        DoEvents
        lst.Visible = True
        stat ""
        
        B_CHANGE = False
    End If
End Sub
Private Sub UpdateList()
    'for use when the filter changes and items are already in the list
    Dim i&, s$, count&
    
    If lst.ListItems().count = 0 Then Exit Sub
    
    i& = 1
    count& = 0
    lst.Visible = False
    prog 0.001
    
    
    stat "updating list"
    
    Do While i& < lst.ListItems().count + 1
        If i& Mod 50 = 0 Or count& Mod 50 = 0 Then
            DoEvents
            If lblCancel.Visible = False Then
                prog 0
                lst.Visible = True
                
                stat "canceled, " + IIf(count& = 0, "0", Format(count&, "###,###")) + " removed"
                If count& > 0 Then B_CHANGE = True
                UpdateCaption
                Exit Sub
            End If
            prog i& / lst.ListItems().count
        End If
        s$ = lst.ListItems(i&).Text
        If FilterCheck(s$) = False Then
            lst.ListItems().Remove i&
            count& = count& + 1
        Else
            'If lst.ListItems(i&).Text <> s$ Then
                lst.ListItems(i&).Text = s$
            'End If
            i& = i& + 1
        End If
    Loop
    
    lst.Sorted = True
    lst.SortKey = 0
    lst.Refresh
    
    UpdateCaption
    
    If count& > 0 Then B_CHANGE = True
    
    lst.Visible = True
    Select Case LANGUAGE$
    Case "english"
        
    Case "french"
        
    Case "german"
        
    Case "spanish"
        
    End Select
    
    stat "complete, " + IIf(count& = 0, "0", Format(count&, "###,###")) + " removed"
    prog 0
End Sub


Private Sub mnuFileSplit100000_Click()
    mnuFileSplit100000.Checked = True
    mnuFileSplitNever.Checked = False
    mnuFileSplit50000.Checked = False
    mnuFileSplit1000000.Checked = False
    mnuFileSplitCustom.Checked = False
    regSet "split", "100000"
End Sub

Private Sub mnuFileSplit1000000_Click()
    mnuFileSplit100000.Checked = False
    mnuFileSplitNever.Checked = False
    mnuFileSplit50000.Checked = False
    mnuFileSplit1000000.Checked = True
    mnuFileSplitCustom.Checked = False
    regSet "split", "1000000"
End Sub

Private Sub mnuFileSplit50000_Click()
    mnuFileSplit100000.Checked = False
    mnuFileSplitNever.Checked = False
    mnuFileSplit50000.Checked = True
    mnuFileSplit1000000.Checked = False
    mnuFileSplitCustom.Checked = False
    regSet "split", "50000"
End Sub

Private Sub mnuFileSplitCustom_Click()
    Dim s$, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the amount of items you want saved per file:" + vbCrLf + vbCrLf + _
              "L517 will automatically split large lists into separate files based on this criteria." + vbCrLf + vbCrLf + _
              "enter 0 to never split."
    Case "french"
        sl$ = "Inscrivez le montant des éléments que vous souhaitez sauvé par fichier:" + vbCrLf + vbCrLf + _
              "L517 fractionne automatiquement des listes importantes en fichiers distincts fondés sur ce critère." + vbCrLf + vbCrLf + _
              "entrez 0 pour séparer jamais."
    Case "german"
        sl$ = "Geben Sie die Anzahl der Einträge Sie pro Datei gespeichert:" + vbCrLf + vbCrLf + _
              "L517 wird automatisch so aufgeteilt großen Listen in separate Dateien auf dieser Kriterien." + vbCrLf + vbCrLf + _
              "0 eingeben nie geteilt."
    Case "spanish"
        sl$ = "introducir la cantidad de elementos que desea guardar el archivo por:" + vbCrLf + vbCrLf + _
              "L517 dividirá automáticamente listas grandes en archivos separados sobre la base de este criterio." + vbCrLf + vbCrLf + _
              "introduzca 0 nunca a dividir."
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517", regGet("split"))
    If s$ = "" Or IsNumeric(s$) = False Then Exit Sub
    
    If CLng(s$) <= 0 Then
        regSet "split", "0"
        mnuFileSplitCustom.Tag = s$
        mnuFileSplitCustom.Caption = "split files [never]"
        If LANGUAGE$ <> "english" Then Change_Language
        Exit Sub
    End If
    
    regSet "split", s$
    mnuFileSplitCustom.Tag = s$
    mnuFileSplitCustom.Caption = "every [" + Format(s$, "###,###") + "] words"
    If LANGUAGE$ <> "english" Then Change_Language
    
    mnuFileSplit100000.Checked = False
    mnuFileSplitNever.Checked = False
    mnuFileSplit50000.Checked = False
    mnuFileSplit1000000.Checked = False
    mnuFileSplitCustom.Checked = True
End Sub

Private Sub mnuFileSplitNever_Click()
    mnuFileSplit100000.Checked = False
    mnuFileSplitNever.Checked = True
    mnuFileSplit50000.Checked = False
    mnuFileSplit1000000.Checked = False
    mnuFileSplitCustom.Checked = False
    regSet "split", ""
End Sub

Private Sub mnuFilesProfilesWPA_Click()
    mnuFilterLenMin.Caption = "minimum length: 8"
    mnuFilterLenMin.Checked = True
    mnuFilterLenMin.Tag = "8"
    regSet "len_min", "8"
    
    mnuFilterLenMax.Caption = "maximum length: 64"
    mnuFilterLenMax.Checked = True
    mnuFilterLenMax.Tag = "64"
    regSet "len_max", "64"
    
    If LANGUAGE$ <> "english" Then Change_Language
    UpdateList
End Sub

Private Sub mnuFileTypeUnix_Click()
    mnuFileTypeWin.Checked = False
    mnuFileTypeWin.Enabled = True
    regSet "filetype", "unix"
    mnuFileTypeUnix.Checked = True
    mnuFileTypeUnix.Enabled = False
End Sub

Private Sub mnuFileTypeWin_Click()
    mnuFileTypeWin.Checked = True
    mnuFileTypeWin.Enabled = False
    regSet "filetype", "win"
    mnuFileTypeUnix.Checked = False
    mnuFileTypeUnix.Enabled = True
End Sub

Private Sub mnuFilterForeign_Click()
    mnuFilterForeign.Checked = Not (mnuFilterForeign.Checked)
    regSet "foreign", CStr(CInt(mnuFilterForeign.Checked))
    
    If mnuFilterForeign.Checked = False Then
        UpdateList
    End If
End Sub

Private Sub mnuFilterHex_Click()
    mnuFilterHex.Checked = Not (mnuFilterHex.Checked)
    regSet "hex", CStr(CInt(mnuFilterHex.Checked))
    
    stat "updating list..."
    UpdateList
    stat ""
End Sub

Private Sub mnuFilterLenMax_Click()
    Dim s$, change%, sl$
    
    change% = CInt(mnuFilterLenMax.Tag)
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter maximum length for an item in the list:" + vbCrLf + vbCrLf + _
              "0 for no maximum."
    Case "french"
        sl$ = "Entrez durée maximale d'un élément dans la liste:" + vbCrLf + vbCrLf + _
              "0 pour pas de maximum."
    Case "german"
        sl$ = "Geben Sie die maximale Länge für ein Element in der Liste:" + vbCrLf + vbCrLf + _
              "0 für kein Maximum."
    Case "spanish"
        sl$ = "entrar en longitud máxima de un elemento de la lista:" + vbCrLf + vbCrLf + _
              "0 si no máximo."
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517", mnuFilterLenMax.Tag)
    If s$ = "" Or IsNumeric(s$) = False Then Exit Sub
    mnuFilterLenMax.Tag = s$
    If s$ = "0" Then
        mnuFilterLenMax.Caption = "maximum length: [none]"
        mnuFilterLenMax.Checked = False
    Else
        mnuFilterLenMax.Caption = "maximum length: " + s$
        mnuFilterLenMax.Checked = True
    End If
    regSet "len_max", s$
    
    If LANGUAGE$ <> "english" Then Change_Language
    
    If change% - CInt(s$) > 0 Then
        UpdateList
    End If
End Sub
Private Sub mnuFilterLenMin_Click()
    Dim s$, change%, sl$
    
    change% = CInt(mnuFilterLenMin.Tag)
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter minimum length for an item in the list:" + vbCrLf + vbCrLf + _
              "0 for no minimum."
    Case "french"
        sl$ = "Entrez durée minimum d'un élément dans la liste:" + vbCrLf + vbCrLf + _
              "0 pour pas de minimum."
    Case "german"
        sl$ = "Geben Sie die minimum Länge für ein Element in der Liste:" + vbCrLf + vbCrLf + _
              "0 für kein Minimum."
    Case "spanish"
        sl$ = "entrar en longitud mínimo de un elemento de la lista:" + vbCrLf + vbCrLf + _
              "0 si no mínimo."
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517", mnuFilterLenMin.Tag)
    If s$ = "" Or IsNumeric(s$) = False Then Exit Sub
    
    mnuFilterLenMin.Tag = s$
    
    If s$ = "0" Then
        mnuFilterLenMin.Caption = "minimum length: [none]"
        mnuFilterLenMin.Checked = False
    Else
        mnuFilterLenMin.Caption = "minimum length: " + s$
        mnuFilterLenMin.Checked = True
    End If
    
    regSet "len_min", s$
    
    If LANGUAGE$ <> "english" Then Change_Language
    
    If change% - CInt(s$) < 0 Then
        UpdateList
    End If
End Sub
Private Sub mnuFilterTextLeft_Click()
    Dim s$, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the string (text) which separates what you want from what you don't want:" + vbCrLf + vbCrLf + _
              "if your list is made up of items such as:" + vbCrLf + _
              "user1:asdf123" + vbCrLf + _
              "jsmith:password1" + vbCrLf + _
              "babyman:hunter2" + vbCrLf + vbCrLf + _
              "and you want the text to the left of the colon, you would enter a colon (:)"
    Case "french"
        sl$ = "entrez la chaîne (texte) qui sépare ce que vous voulez de ce que vous ne voulez pas:" + vbCrLf + vbCrLf + _
              "Si votre liste est composée d'éléments tels que:" + vbCrLf + _
              "user1: asdf123" + vbCrLf + _
              "jsmith: password1" + vbCrLf + _
              "babyman: hunter2" + vbCrLf + vbCrLf + _
              "et vous voulez que le texte à gauche du colon, tu veux entrer dans un deux-points (:)"
    Case "german"
        sl$ = "geben Sie die Zeichenfolge (Text), was trennt, was Sie von dem, was Sie nicht wollen, möchten:" + vbCrLf + vbCrLf + _
              "wenn Ihre Liste der Elemente enthalten wie:" + vbCrLf + _
              "user1: asdf123" + vbCrLf + _
              "jsmith: password1" + vbCrLf + _
              "babyman: hunter2" + vbCrLf + vbCrLf + _
              "und Sie möchten den Text auf der linken Seite des Dickdarms, würden Sie einen Doppelpunkt (:) eingeben"
    Case "spanish"
        sl$ = "entrar en la cadena (texto) que separa lo que quiere de lo que no quieren:" + vbCrLf + vbCrLf + _
              "Si su lista se compone de elementos tales como:" + vbCrLf + _
              "user1: asdf123" + vbCrLf + _
              "jsmith: password1" + vbCrLf + _
              "babyman: hunter2" + vbCrLf + vbCrLf + _
              "y desea que el texto a la izquierda del colon, tiene que escribir dos puntos (:)"
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517")
    If s$ = "" Then
        mnuFilterTextLeft.Caption = "text to the left of [string]"
        mnuFilterTextLeft.Checked = False
        mnuFilterTextLeft.Tag = ""
    Else
        mnuFilterTextLeft.Caption = "text to the left of [" + s$ + "]"
        mnuFilterTextLeft.Checked = True
        mnuFilterTextLeft.Tag = s$
    End If
    If LANGUAGE$ <> "english" Then Change_Language
    
    regSet "text_left", s$
    UpdateList
End Sub
Private Sub mnuFilterTextRight_Click()
    Dim s$, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the string (text) which separates what you want from what you don't want:" + vbCrLf + vbCrLf + _
              "if your list is made up of items such as:" + vbCrLf + _
              "user1:asdf123" + vbCrLf + _
              "jsmith:password1" + vbCrLf + _
              "babyman:hunter2" + vbCrLf + vbCrLf + _
              "and you want the text to the right of the colon, you would enter a colon (:)"
    Case "french"
        sl$ = "entrez la chaîne (texte) qui sépare ce que vous voulez de ce que vous ne voulez pas:" + vbCrLf + vbCrLf + _
              "Si votre liste est composée d'éléments tels que:" + vbCrLf + _
              "user1: asdf123" + vbCrLf + _
              "jsmith: password1" + vbCrLf + _
              "babyman: hunter2" + vbCrLf + vbCrLf + _
              "et vous voulez que le texte à droit du colon, tu veux entrer dans un deux-points (:)"
    Case "german"
        sl$ = "geben Sie die Zeichenfolge (Text), was trennt, was Sie von dem, was Sie nicht wollen, möchten:" + vbCrLf + vbCrLf + _
              "wenn Ihre Liste der Elemente enthalten wie:" + vbCrLf + _
              "user1: asdf123" + vbCrLf + _
              "jsmith: password1" + vbCrLf + _
              "babyman: hunter2" + vbCrLf + vbCrLf + _
              "und Sie möchten den Text auf der richtig Seite des Dickdarms, würden Sie einen Doppelpunkt (:) eingeben"
    Case "spanish"
        sl$ = "entrar en la cadena (texto) que separa lo que quiere de lo que no quieren:" + vbCrLf + vbCrLf + _
              "Si su lista se compone de elementos tales como:" + vbCrLf + _
              "user1: asdf123" + vbCrLf + _
              "jsmith: password1" + vbCrLf + _
              "babyman: hunter2" + vbCrLf + vbCrLf + _
              "y desea que el texto a la derecho del colon, tiene que escribir dos puntos (:)"
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517")
    If s$ = "" Then
        mnuFilterTextRight.Caption = "text to the right of [string]"
        mnuFilterTextRight.Checked = False
        mnuFilterTextRight.Tag = ""
    Else
        mnuFilterTextRight.Caption = "text to the right of [" + s$ + "]"
        mnuFilterTextRight.Checked = True
        mnuFilterTextRight.Tag = s$
    End If
    If LANGUAGE$ <> "english" Then Change_Language
    
    regSet "text_right", s$
    UpdateList
End Sub

Public Sub HeapSort1(ByRef pvarArray As Variant)
    'sorts array alphabetically, very quickly!
    'used in the analyzer
    Dim i&, iMin&, imax&, varSwap As Variant
   
    iMin = LBound(pvarArray)
    imax = UBound(pvarArray)
    For i = (imax + iMin) \ 2 To iMin Step -1
        Heap1 pvarArray, i, iMin, imax
    Next i
    For i = imax To iMin + 1 Step -1
        varSwap = pvarArray(i)
        pvarArray(i) = pvarArray(iMin)
        pvarArray(iMin) = varSwap
        Heap1 pvarArray, iMin, iMin, i - 1
    Next i
End Sub
Private Sub Heap1(ByRef pvarArray As Variant, ByVal i As Long, iMin As Long, imax As Long)
    'used in heap sort
    Dim lngLeaf&, varSwap As Variant
    
    Do
        lngLeaf = i + i - (iMin - 1)
        Select Case lngLeaf
            Case Is > imax: Exit Do
            Case Is < imax: If pvarArray(lngLeaf + 1) > pvarArray(lngLeaf) Then lngLeaf = lngLeaf + 1
        End Select
        If pvarArray(i) > pvarArray(lngLeaf) Then Exit Do
        varSwap = pvarArray(i)
        pvarArray(i) = pvarArray(lngLeaf)
        pvarArray(lngLeaf) = varSwap
        i = lngLeaf
    Loop
End Sub
Private Function ReverseString$(s$)
    'returns the mirror of a string
    'abcd would return dcba
    Dim i&, sresult$
    sresult$ = ""
    For i& = 1 To Len(s$)
        sresult$ = Mid(s$, i&, 1) + sresult$
    Next i&
    ReverseString$ = sresult$
End Function
Private Sub mnuGenAnalANAL_Click()
    'analyze list
    Dim i&, s$, s2$, j%, c_min_pre%, c_min_post%, c_count%, spatt$, sarr$(), iarr&
    Dim srev$(), irev&, ff%, sfile$, sfile2$, sMsg$
    Dim spre$(), cpre%(), ipre&, k&, bgot!
    
    If lst.ListItems().count < 2 Then
        Select Case LANGUAGE$
        Case "english"
            s$ = "list analysis requires more than 1 item in the list." + vbCrLf + vbCrLf + "the analysis finds and sorts repeating patterns within a list; useful for password-list generation."
        Case "french"
            s$ = "analyse liste exige plus de 1 point dans la liste." + vbCrLf + vbCrLf + "les trouvailles d'analyse et les trie en répétant les tendances au sein d'une liste; génération utile pour un mot de passe liste."
        Case "german"
            s$ = "Liste Analyse erfordert mehr als 1 Element in der Liste." + vbCrLf + vbCrLf + "die Analyse gesucht und sortiert werden sich wiederholende Muster innerhalb einer Liste; nützlich für die Passwort-Liste Generation."
        Case "spanish"
            s$ = "el análisis de la lista requiere más de 1 punto en la lista." + vbCrLf + vbCrLf + "los hallazgos y el análisis de tipo de patrones repetitivos dentro de una lista, la generación de la lista de útiles para la contraseña."
        End Select
        frmMsg.MsgBocks s$, vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    sfile$ = App.Path + IIf(Right(App.Path, 1) = "\", "", "\")
    sfile2$ = sfile$ + "_postfix.txt"
    sfile$ = sfile$ + "_prefix.txt"
    
    c_min_pre% = CInt(mnuGenAnalMinPre.Tag)   'minimum length of prefix
    c_min_post% = CInt(mnuGenAnalMinPost.Tag) 'minimum length of postfix
    c_count% = CInt(mnuGenAnalCount.Tag)      'minimum count to be a 'pattern'
    
    stat "analyzing prefixes..."
    DoEvents
    ipre& = 0
    
    lst.Visible = False
    prog 0.001
    
    For i& = 1 To lst.ListItems.count()
        If i Mod 250 = 0 Then
            DoEvents
            If lblCancel.Visible = False Then
                'stop!
                lst.Visible = True
                stat "canceled"
                Exit Sub
            End If
            prog i / lst.ListItems().count
        End If
        
        s$ = lst.ListItems(i&).Text
        If mnuGenAnalCase.Checked = True Then s$ = LCase(s$)
        
        'srev$() holds the reversed strings [for post-fixes]
        ReDim Preserve srev$(i& - 1)
        srev$(i& - 1) = ReverseString$(s$)
        
        If i& < lst.ListItems().count And Len(s$) >= c_min_pre% Then
            s2$ = lst.ListItems(i& + 1).Text
            If mnuGenAnalCase.Checked = True Then s2$ = LCase(s2$)
            
            For j% = Len(s2$) To c_min_pre% Step -1
                If Left(s$, j%) = Left(s2$, j%) Then
                    'we have a match
                    spatt$ = Left(s$, j%)
                    If ipre& > 0 Then
                        For k& = 0 To UBound(spre$())
                            If spre$(k&) = spatt$ Then
                                'found it
                                cpre%(k&) = cpre%(k&) + 1
                                If cpre%(k&) >= c_min_pre% Then GoTo AFTER_J
                                
                                k& = -1
                                Exit For
                            End If
                        Next k&
                        If k& <> -1 Then
                            ReDim Preserve spre$(ipre&)
                            ReDim Preserve cpre%(ipre&)
                            spre$(ipre&) = spatt$
                            cpre%(ipre&) = 1
                            ipre& = ipre& + 1
                        End If
                    Else
                        ReDim Preserve spre$(ipre&)
                        ReDim Preserve cpre%(ipre&)
                        spre$(ipre&) = spatt$
                        cpre%(ipre&) = 1
                        ipre& = ipre& + 1
                    End If
                End If
            Next j%
AFTER_J:
        End If
    Next i&
    
    'ipre& is the last index of the cpre/spre arrays
    '0 means no patterns found
    If ipre& = 0 Then
        sMsg$ = "" 'no prefix patterns found."
        sfile$ = ""
    Else
        Erase sarr$()
        iarr& = 0
        For i& = 0 To UBound(cpre%())
            If cpre%(i&) >= c_count% Then
                ReDim Preserve sarr$(iarr&)
                sarr$(iarr&) = spre$(i&)
                iarr& = iarr& + 1
            End If
        Next i&
        sMsg$ = "" + CStr(iarr&) + " prefix patterns found, saved to '" + sfile$ + "'"
        Do While Len(Dir(sfile$)) <> 0
            DoEvents
            Kill sfile$
        Loop
        
        ff% = FreeFile
        Open sfile$ For Binary Access Write As #ff%
            For i& = 0 To UBound(sarr$())
                Put #ff%, , CStr(sarr$(i&) + IIf(mnuFileTypeUnix.Checked, Chr(10), vbCrLf))
            Next i&
        Close #ff%
        DoEvents
        Erase sarr$()
        Erase cpre%()
        Erase spre$()
    End If
    
    stat "sorting postfixes..."
    DoEvents
    HeapSort1 srev$()
    DoEvents
    
    stat "analyzing postfixes..."
    DoEvents
     
    ipre& = 0
    For i& = 1 To UBound(srev$()) - 1
        If i Mod 250 = 0 Then
            DoEvents
            If lblCancel.Visible = False Then
                'stop!
                lst.Visible = True
                stat "canceled"
                Exit Sub
            End If
            prog i / UBound(srev$())
        End If
        
        s$ = srev$(i&)
        If mnuGenAnalCase.Checked = True Then s$ = LCase(s$)
        
        If Len(s$) >= c_min_post% Then
            s2$ = srev$(i& + 1)
            If mnuGenAnalCase.Checked = True Then s2$ = LCase(s2$)
            
            For j% = Len(s2$) To c_min_post% Step -1
                If Left(s$, j%) = Left(s2$, j%) Then
                    'we have a match
                    spatt$ = Left(s$, j%)
                    If ipre& > 0 Then
                        For k& = 0 To UBound(spre$())
                            If spre$(k&) = spatt$ Then
                                'found it
                                cpre%(k&) = cpre%(k&) + 1
                                If cpre%(k&) >= c_min_post% Then GoTo AFTER_J_AGAIN
                                
                                k& = -1
                                Exit For
                            End If
                        Next k&
                        If k& <> -1 Then
                            ReDim Preserve spre$(ipre&)
                            ReDim Preserve cpre%(ipre&)
                            spre$(ipre&) = spatt$
                            cpre%(ipre&) = 1
                            ipre& = ipre& + 1
                        End If
                    Else
                        ReDim Preserve spre$(ipre&)
                        ReDim Preserve cpre%(ipre&)
                        spre$(ipre&) = spatt$
                        cpre%(ipre&) = 1
                        ipre& = ipre& + 1
                    End If
                End If
            Next j%
AFTER_J_AGAIN:
        End If
    Next i&
    
    stat "finalizing..."
    DoEvents
    
    'ipre& is the last index of the cpre/spre arrays
    '0 means no patterns found
    If ipre& = 0 Then
        sMsg$ = "" 'no prefix patterns found."
        sfile$ = ""
    Else
        Erase sarr$()
        iarr& = 0
        For i& = 0 To UBound(cpre%())
            If cpre%(i&) >= c_count% Then
                ReDim Preserve sarr$(iarr&)
                sarr$(iarr&) = spre$(i&)
                iarr& = iarr& + 1
            End If
        Next i&
        sMsg$ = sMsg$ + vbCrLf + "" + CStr(iarr&) + " postfix patterns found, saved to '" + sfile2$ + "'"
        Do While Len(Dir(sfile2$)) <> 0
            DoEvents
            Kill sfile2$
        Loop
        
        ff% = FreeFile
        Open sfile2$ For Binary Access Write As #ff%
            For i& = 0 To UBound(sarr$())
                Put #ff%, , CStr(ReverseString$(sarr$(i&)) + IIf(mnuFileTypeUnix.Checked, Chr(10), vbCrLf))
            Next i&
        Close #ff%
        DoEvents
        Erase sarr$()
        Erase cpre%()
        Erase spre$()
    End If
    
    If sfile$ <> "" Or sfile2$ <> "" Then
        'we got a list of prefixes OR postfixes...
        Select Case LANGUAGE$
        Case "english"
            s$ = "do you want to open the file in notepad?"
        Case "french"
            s$ = "Voulez-vous ouvrir le fichier dans le Notepad?"
        Case "german"
            s$ = "wollen Sie die Datei(n) in Notepad zu öffnen?"
        Case "spanish"
            s$ = "¿quieres abrir el archivo(s) en notepad?"
        End Select
        If frmMsg.MsgBocks(sMsg$ + vbCrLf + vbCrLf + s$, vbYesNo + vbQuestion) = vbYes Then
            If sfile$ <> "" Then
                Shell "explorer " + Chr(34) + sfile$ + Chr(34), vbNormalFocus
            End If
            If sfile2$ <> "" Then
                Shell "explorer " + Chr(34) + sfile2$ + Chr(34), vbNormalFocus
            End If
        End If
    End If
    
    Erase sarr$()
    Erase srev$()
    Erase spre$()
    Erase cpre%()
    
    prog 0
    stat ""
    lst.Visible = True
End Sub

Private Sub mnuGenAnalCase_Click()
    mnuGenAnalCase.Checked = Not (mnuGenAnalCase.Checked)
    regSet "pattern_case", CStr(CInt(mnuGenAnalCase.Checked))
End Sub

Private Sub mnuGenAnalCount_Click()
    'pattern_count
    Dim s$, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the minimum number of times a pattern has to appear in the list before it is recognized as a pattern:"
    Case "french"
        sl$ = "Entrez le nombre minimum de fois un modèle qui doit apparaître dans la liste avant qu'elle soit reconnue comme un motif:"
    Case "german"
        sl$ = "Geben Sie die minimale Anzahl, wie oft ein Muster hat sich in der Liste angezeigt wird, bevor es als ein Muster erkannt wird:"
    Case "spanish"
        sl$ = "introducir el número mínimo de veces que un patrón tiene que aparecer en la lista antes de que sea reconocido como un patrón:"
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517", regGet("pattern_count"))
    If s$ = "" Or IsNumeric(s$) = False Then Exit Sub
    
    If CInt(s$) <= 0 Then Exit Sub
    
    mnuGenAnalCount.Caption = "minimum count: [" + s$ + "]"
    mnuGenAnalCount.Tag = s$
    
    If LANGUAGE$ <> "english" Then Change_Language
    regSet "pattern_count", s$
End Sub

Private Sub mnuGenAnalHelp_Click()
    Dim s$
    Select Case LANGUAGE$
    Case "english"
        s$ = "when L517 'analyzes' a pre-loaded list, it goes through both prefixes and post-fixes, looking for patterns." + vbCrLf + vbCrLf + _
             "many common passwords begin and end with the same characters (some start with colors 'blue' or 'red', end with '123' or 'abc')." + vbCrLf + vbCrLf + _
             "the analyzer will filter out patterns found at the beginning and end of list items and separate them into lists (_prefix.txt and _postfix.txt), and will save them to the L517's current directory." + vbCrLf + vbCrLf + _
             "this option is ONLY useful if you have a list of passwords (preferably phished)."
             
    Case "french"
        s$ = "quand L517 'analyse' une liste pré-chargé, il va à travers les deux préfixes et post-fixe, la recherche de modèles." + vbCrLf + vbCrLf + _
             "de nombreux mots de passe courants commencent et se terminent avec les mêmes personnages (certains commencent par des couleurs «bleu» ou «rouge», se terminent par «123» ou «ABC»)." + vbCrLf + vbCrLf + _
             "l'analyseur de filtrer les modèles observés au début et de fin de liste et de les séparer dans des listes (_prefix.txt et _postfix.txt), et les enregistrer dans le répertoire courant du L517." + vbCrLf + vbCrLf + _
             "Cette option est uniquement utile si vous avez une liste de mots de passe (de préférence victimes de phishing)."
    Case "german"
        s$ = "wenn L517 'analysiert' ein Pre-Liste geladen ist, geht es durch die beiden Prä-und Post-Korrekturen auf der Suche nach Mustern." + vbCrLf + vbCrLf + _
             "viele gemeinsame Passwörter beginnen und enden mit dem gleichen Zeichen (einige beginnen mit den Farben 'blau' oder 'rot', enden mit '123' oder 'abc')." + vbCrLf + vbCrLf + _
             "das Analysegerät wird herausgefiltert Muster am Anfang und am Ende der Liste Produkte und trennen Sie sie bitte in Listen (_prefix.txt und _postfix.txt) gefunden, und wird sie in die aktuelle Verzeichnis der L517 zu retten." + vbCrLf + vbCrLf + _
             "Diese Option ist nur dann sinnvoll, wenn Sie eine Liste von Passwörtern (die vorzugsweise Phishing)."
    Case "spanish"
        s$ = "cuando se analiza 'L517' pre-cargado lista, pasa por ambos prefijos y post-fija, en busca de patrones." + vbCrLf + vbCrLf + _
             "muchas contraseñas comunes comienzan y terminan con los mismos personajes (algunos empiezan con los colores 'azul' o 'rojo', con el fin '123 'o' abc ')." + vbCrLf + vbCrLf + _
             "el analizador filtrar los patrones encontrados en el comienzo y el final de elementos de lista y en listas separadas (_prefix.txt y _postfix.txt), y guardarlos en el directorio actual del L517's." + vbCrLf + vbCrLf + _
             "Esta opción sólo es útil si usted tiene una lista de contraseñas (de preferencia phishing)."
    End Select
    frmMsg.MsgBocks s$, vbInformation + vbOKOnly
End Sub

Private Sub mnuGenAnalMinPost_Click()
    'minimum length of postfix pattern
    Dim s$, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the minimum length of a postfix pattern to search for in the list:"
    Case "french"
        sl$ = "Entrez la durée minimale d'un motif postfix à rechercher dans la liste:"
    Case "german"
        sl$ = "geben Sie die minimale Länge einer postfix Muster für die Suche nach in der Liste:"
    Case "spanish"
        sl$ = "introducir la duración mínima de un patrón de postfix para buscar en la lista:"
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517", regGet("min_postfix"))
    If s$ = "" Or IsNumeric(s$) = False Then Exit Sub
    
    If CInt(s$) <= 0 Then Exit Sub
    
    mnuGenAnalMinPost.Caption = "min. length of postfix: [" + s$ + "]"
    mnuGenAnalMinPost.Tag = s$
    regSet "min_postfix", s$
    
    If LANGUAGE$ <> "english" Then Change_Language
End Sub

Private Sub mnuGenAnalMinPre_Click()
    'minimum length of prefix pattern
    Dim s$, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the minimum length of a prefix pattern to search for in the list:"
    Case "french"
        sl$ = "Entrez la durée minimale d'un motif préfixe à rechercher dans la liste:"
    Case "german"
        sl$ = "geben Sie die minimale Länge einer prefix Muster für die Suche nach in der Liste:"
    Case "spanish"
        sl$ = "introducir la duración mínima de un patrón de prefijo para buscar en la lista:"
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517", regGet("min_prefix"))
    If s$ = "" Or IsNumeric(s$) = False Then Exit Sub
    
    If CInt(s$) <= 0 Then Exit Sub
    
    mnuGenAnalMinPre.Caption = "min. length of prefix: [" + s$ + "]"
    mnuGenAnalMinPre.Tag = s$
    regSet "min_prefix", s$
    
    If LANGUAGE$ <> "english" Then Change_Language
End Sub

Private Sub mnuGenDate_Click(Index As Integer)
    Dim sstart$, sstop$, ista%, isto%, id%, im%, iy%
    Dim sd$, sm$, sy$, sdate$, count&, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the STARTING two-digit year:"
    Case "french"
        sl$ = "Entrez le point de départ année à deux chiffres:"
    Case "german"
        sl$ = "Geben Sie den Anfangspunkt zweistellige Jahreszahl:"
    Case "spanish"
        sl$ = "introducir los dos dígitos del año de partida:"
    End Select
    sstart$ = frmMsg.InputBocks(sl$ + vbCrLf + vbCrLf + "i.e. 91", "L517")
    If sstart$ = "" Or IsNumeric(sstart$) = False Then Exit Sub
    
    ista% = CInt(sstart$)
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the ENDING two-digit year:"
    Case "french"
        sl$ = "Entrez la fin année à deux chiffres:"
    Case "german"
        sl$ = "Geben Sie die Endung zweistellige Jahreszahl:"
    Case "spanish"
        sl$ = "introducir los dos dígitos del año de termina:"
    End Select
    sstop$ = frmMsg.InputBocks(sl$ + vbCrLf + vbCrLf + "i.e. 99", "l517")
    If sstop$ = "" Or IsNumeric(sstop$) = False Then Exit Sub
    
    isto% = CInt(sstop$)
    
    If ista% > 99 Or isto% > 99 Or ista% = isto% Then Exit Sub
    
    'convert to 4 digit years
    If ista% > isto% Then
        ista% = 1900 + ista%
        isto% = 2000 + isto%
    Else
        If ista% > 20 Then
            ista% = 1900 + ista%
            isto% = 1900 + isto%
        Else
            ista% = 2000 + ista%
            isto% = 2000 + isto%
        End If
    End If
    
    count& = 1
    
    stat "generating dates..."
    prog 0.001
    lst.Visible = False
    
    For iy% = ista% To isto%
        sy$ = CStr(iy%)
        If Index% Mod 2 = 0 Then sy$ = Right(sy$, 2)
        
        DoEvents
        If lblCancel.Visible = False Then
            prog 0
            stat "canceled"
            UpdateCaption
            lst.Visible = True
            B_CHANGE = True
            Exit Sub
        End If
        
        prog count& / ((isto% - ista%) * 365)
        
        For im% = 1 To 12
            sm$ = CStr(im%)
            If Index% >= 4 Then
                sm$ = NumToMonth$(im%)
            Else
                If Len(sm$) = 1 Then sm$ = "0" + sm$
            End If
            
            For id% = 1 To 31
                sd$ = CStr(id%)
                If Len(sd$) = 1 Then sd$ = "0" + sd$
                
                Select Case im%
                Case 4, 6, 9, 11
                    If id% = 31 Then Exit For
                Case 2
                    If iy% Mod 4 = 0 Then
                        If id% = 16 Then Exit For
                    Else
                        If id% = 15 Then Exit For
                    End If
                End Select
                
                Select Case Index%
                Case 0, 1, 4, 5
                    sdate$ = sm$ + sd$ + sy$
                Case 2, 3, 6, 7
                    sdate$ = sd$ + mnuGenDateSep.Tag + sm$ + mnuGenDateSep.Tag + sy$
                End Select
                
                If FilterCheck(sdate$) = True Then
                    lst.ListItems().add , , sdate$
                End If
                count& = count& + 1
            Next id%
        Next im%
    Next iy%
    
    B_CHANGE = True
    lst.Visible = True
    prog 0
    stat ""
    UpdateCaption
End Sub
Private Function NumToMonth$(inum%)
    'converts month number to name
    Dim s$
    
    Select Case inum%
    Case 1
        s$ = "january"
    Case 2
        s$ = "february"
    Case 3
        s$ = "march"
    Case 4
        s$ = "april"
    Case 5
        s$ = "may"
    Case 6
        s$ = "june"
    Case 7
        s$ = "july"
    Case 8
        s$ = "august"
    Case 9
        s$ = "september"
    Case 10
        s$ = "october"
    Case 11
        s$ = "november"
    Case 12
        s$ = "december"
    End Select
    
    NumToMonth$ = s$
End Function

Private Sub mnuGenDateSep_Click()
    Dim s$, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter character or string to separate the days, months, and years in the date:" + vbCrLf + vbCrLf + _
              "pressing Cancel will default to no separator"
    Case "french"
        sl$ = "Entrez caractère ou une chaîne de séparer les jours, des mois et des années à la date:" + vbCrLf + vbCrLf + _
              "appuyant sur Annuler sera par défaut sans séparateur"
    Case "german"
        sl$ = "Geben Sie Zeichen oder eine Zeichenkette zur Trennung der Tage, Monate und Jahre in der Ausgabe:" + vbCrLf + vbCrLf + _
              "Drücken von Abbrechen wird kein Trennzeichen Standard"
    Case "spanish"
        sl$ = "entrar carácter o cadena para separar los días, meses y años en la fecha:" + vbCrLf + vbCrLf + _
              "pulsa Cancelar se pondrá por defecto no separador"
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517")
    
    mnuGenDateSep.Caption = "separator: " + IIf(s$ = "", "[none]", s$)
    mnuGenDateSep.Tag = s$
    regSet "sep", s$
    
    mnuGenDate(0).Caption = "mm" + s$ + "dd" + s$ + "yy"
    mnuGenDate(1).Caption = "mm" + s$ + "dd" + s$ + "yyyy"
    mnuGenDate(2).Caption = "dd" + s$ + "mm" + s$ + "yy"
    mnuGenDate(3).Caption = "dd" + s$ + "mm" + s$ + "yyyy"
    mnuGenDate(4).Caption = "mmm" + s$ + "dd" + s$ + "yy"
    mnuGenDate(5).Caption = "mmm" + s$ + "dd" + s$ + "yyyy"
    mnuGenDate(6).Caption = "dd" + s$ + "mmm" + s$ + "yy"
    mnuGenDate(7).Caption = "dd" + s$ + "mmm" + s$ + "yyyy"
    
    If LANGUAGE$ <> "english" Then Change_Language
End Sub

Private Sub mnuGenFiles_Click()
    Dim i&, s$, j&, spath$
    
    'recursively scan directories
    
    'learn to read words from:
    ' word documents
    ' mp3's
    ' jpg's
    ' etc
    
    'this is a project in and of itself
    
    spath$ = regGet("last_path")
    If spath$ = "" Or Len(Dir(spath$)) = 0 Then
        spath$ = App.Path
    End If
    
    'get directory
    s$ = GetFolder$("select folder" + vbCrLf + "*** this will include ALL sub-directories! ***", spath$, False)
    If s$ = "" Then Exit Sub
    
    regSet "last_path", s$ + IIf(Right(s$, 1) = "\", "", "\")
    
    stat "searching subfolders"
    'collect subfolders
    lstDir.Clear
    lstDir.AddItem s$
    Do Until i& = lstDir.ListCount
        dir1.Path = lstDir.List(i&)
        If dir1.ListCount > 0 Then
            For j& = 0 To dir1.ListCount - 1
                lstDir.AddItem dir1.List(j&)
            Next j&
        End If
        i& = i& + 1
    Loop
    
    prog 0.001
    DoEvents
    
    lstFile.Clear
    
    stat "reading folders"
    'collect files from subfolders
    For i& = 0 To lstDir.ListCount - 1
        DoEvents
        If lblCancel.Visible = False Then
            prog 0
            stat "canceled"
            Exit Sub
        End If
        
        prog (i& + 1) / lstDir.ListCount
        
        s$ = Dir(lstDir.List(i&) + IIf(Right(lstDir.List(i&), 1) = "\", "", "\") + "*.*")
        If s$ <> "" Then
            Do
                lstFile.AddItem lstDir.List(i&) + IIf(Right(lstDir.List(i&), 1) = "\", "", "\") + s$
                s$ = Dir$
            Loop Until s$ = ""
        End If
    Next i&
    
    For i& = 0 To lstFile.ListCount - 1
        DoEvents
        If lblCancel.Visible = False Then
            prog 0
            stat "canceled"
            Exit Sub
        End If
        
        stat "parsing '" + GetFileName$(lstFile.List(i&)) + "'"
        prog (i& + 1) / lstFile.ListCount
        
        If LoadList(lstFile.List(i&), False) = False Then
            prog 0
            stat "canceled"
            lst.Visible = True
            Exit Sub
        End If
        
        lblCancel.Visible = True
    Next i&
    
    lst.Visible = True
    prog 0
    stat ""
    
    lstDir.Clear
    lstFile.Clear
End Sub

Private Sub mnuGenPhoneArea_Click()
    Dim scity$, sDat$, i&, j&, s$, sep$, sarr$(), iarr&
    Dim ff%, sfile$, isplit&, icount&, sl$
    
    icount& = 1
    isplit& = 0
    s$ = regGet("split")
    If IsNumeric(s$) = True Then isplit& = CLng(s$)
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the name of the city you want to grab all the phone numbers for (United States only!):"
    Case "french"
        sl$ = "Entrez le nom de la ville que vous voulez récupérer tous les numéros de téléphone pour (États-Unis seulement!):"
    Case "german"
        sl$ = "Geben Sie den Namen der Stadt, die Sie wollen alle Telefonnummern (Vereinigte Staaten nur greifen!):"
    Case "spanish"
        sl$ = "escriba el nombre de la ciudad que desee para agarrar todos los números de teléfono (Estados Unidos solamente!):"
    End Select
    scity$ = Trim(frmMsg.InputBocks(sl$, "L517", ""))
    If scity$ = "" Then Exit Sub
    
    stat "grabbing phone prefixes..."
    DoEvents
    
    scity$ = Replace(scity$, " ", "+")
    sDat$ = webgetsource$("http://www.melissadata.com/lookups/phonelocation.asp?number=" + scity$ + "")
    
    iarr& = 0
    j& = 0
    sep$ = mnuGenPhoneSep.Tag
    
    stat "parsing prefixes..."
    Do
        DoEvents
        i& = InStr(j + 1, sDat$, "<a href=" + Chr(34) + "phonelocation.asp?number=")
        j& = InStr(i& + Len("<a href=Xphonelocation.asp?number="), sDat$, Chr(34))
        If i& = 0 Or j& = 0 Then Exit Do
        i& = i& + Len("<a href=Xphonelocation.asp?number=")
        s$ = Mid(sDat$, i&, j& - i&)
        If s$ <> "" Then
            ReDim Preserve sarr$(iarr&)
            sarr$(iarr&) = Mid(s$, 1, 3) + sep$ + Mid(s$, 4, 3)
            iarr& = iarr& + 1
        End If
    Loop Until j& = 0
    
    Do While Len(Dir(App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + scity$ + ".txt")) <> 0
        Kill App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + scity$ + ".txt"
        DoEvents
    Loop
    
    lst.Visible = False
    scity$ = scity$ + "(area)"
    stat "saving to '" + scity$ + ".txt'"
    prog 0.001
    DoEvents
    
    ff% = FreeFile
    Open App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + scity$ + ".txt" For Binary Access Write As #ff%
        For i& = 0 To UBound(sarr$())
            If i& Mod 5 = 0 Then
                If lblCancel.Visible = False Then
                    'cancel was clicked
                    Exit For
                End If
                UpdateCaption
                prog (i& + 1) / UBound(sarr$())
                DoEvents
            End If
            
            'last 4 digits
            For j& = 0 To 9999
                'split check
                If isplit& > 0 Then
                    If icount& >= isplit& Then
                        DoEvents
                        icount& = 0
                        'close old file
                        Close #ff%
                        'find next file name in sequence
                        sfile$ = NextFile$(App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + scity$ + ".txt")
                        'open new file
                        ff% = FreeFile
                        Open sfile$ For Binary Access Write As #ff%
                        DoEvents
                    End If
                End If
                
                Put #ff%, , CStr(sarr$(i&) + sep$ + String(4 - Len(CStr(j&)), "0") + CStr(j&)) + CStr(IIf(mnuFileTypeUnix.Checked, Chr(10), vbCrLf))
                icount& = icount& + 1
            Next j&
        Next i&
    Close #ff%
    
    lst.Visible = True
    prog 0
    
    Erase sarr$()
    
    stat "saved to '" + scity$ + ".txt'"
    Select Case LANGUAGE$
    Case "english"
        s$ = "phone numbers for the city '" + scity$ + "' have been saved to '" + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + _
             scity$ + ".txt'." + vbCrLf + vbCrLf + "do you want to open the containing folder?"
    Case "french"
        s$ = "numéros de téléphone pour la ville '" + scity$ + "' ont été enregistrées dans '" + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + _
             scity$ + ".txt'." + vbCrLf + vbCrLf + "Voulez-vous ouvrez le dossier contenant?"
    Case "german"
        s$ = "Telefonnummern für die Stadt '" + scity$ + "' gespeichert wurden, um '" + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + _
             scity$ + ".txt'." + vbCrLf + vbCrLf + "wollen Sie mit dem Ordner zu öffnen?"
    Case "spanish"
        s$ = "números de teléfono de la ciudad '" + scity$ + "' se han guardado en '" + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + _
             scity$ + ".txt'." + vbCrLf + vbCrLf + "¿quieres abrir la carpeta que contiene?"
    End Select
    If frmMsg.MsgBocks(s$, vbQuestion + vbYesNo, "L517") = vbYes Then
        Shell "explorer /select," + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + scity$ + ".txt", vbNormalFocus
    End If
End Sub

Private Sub mnuGenPhoneNoarea_Click()
    Dim scity$, sDat$, i&, j&, s$, sep$, sarr$(), iarr&
    Dim ff%, sfile$, isplit&, icount&, sl$
    
    icount& = 1
    isplit& = 0
    s$ = regGet("split")
    If IsNumeric(s$) = True Then isplit& = CLng(s$)
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the name of the city you want to grab all the phone numbers for (United States only!):"
    Case "french"
        sl$ = "Entrez le nom de la ville que vous voulez récupérer tous les numéros de téléphone pour (États-Unis seulement!):"
    Case "german"
        sl$ = "Geben Sie den Namen der Stadt, die Sie wollen alle Telefonnummern (Vereinigte Staaten nur greifen!):"
    Case "spanish"
        sl$ = "escriba el nombre de la ciudad que desee para agarrar todos los números de teléfono (Estados Unidos solamente!):"
    End Select
    scity$ = Trim(frmMsg.InputBocks(sl$, "L517", ""))
    If scity$ = "" Then Exit Sub
    
    stat "grabbing phone prefixes..."
    DoEvents
    
    scity$ = Replace(scity$, " ", "+")
    sDat$ = webgetsource$("http://www.melissadata.com/lookups/phonelocation.asp?number=" + scity$ + "")
    
    iarr& = 0
    j& = 0
    sep$ = mnuGenPhoneSep.Tag
    
    stat "parsing prefixes..."
    Do
        DoEvents
        i& = InStr(j + 1, sDat$, "<a href=" + Chr(34) + "phonelocation.asp?number=")
        j& = InStr(i& + Len("<a href=Xphonelocation.asp?number="), sDat$, Chr(34))
        If i& = 0 Or j& = 0 Then Exit Do
        i& = i& + Len("<a href=Xphonelocation.asp?number=")
        s$ = Mid(sDat$, i&, j& - i&)
        If s$ <> "" Then
            ReDim Preserve sarr$(iarr&)
            sarr$(iarr&) = Mid(s$, 4, 3)
            iarr& = iarr& + 1
        End If
    Loop Until j& = 0
    
    Do While Len(Dir(App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + scity$ + ".txt")) <> 0
        Kill App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + scity$ + ".txt"
        DoEvents
    Loop
    
    lst.Visible = False
    stat "saving to '" + scity$ + ".txt'"
    prog 0.001
    DoEvents
    
    ff% = FreeFile
    Open App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + scity$ + ".txt" For Binary Access Write As #ff%
        For i& = 0 To UBound(sarr$())
            If i& Mod 5 = 0 Then
                If lblCancel.Visible = False Then
                    'cancel was clicked
                    Exit For
                End If
                UpdateCaption
                prog (i& + 1) / UBound(sarr$())
                DoEvents
            End If
            
            'last 4 digits
            For j& = 0 To 9999
                'split check
                If isplit& > 0 Then
                    If icount& >= isplit& Then
                        DoEvents
                        icount& = 0
                        'close old file
                        Close #ff%
                        'find next file name in sequence
                        sfile$ = NextFile$(App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + scity$ + ".txt")
                        'open new file
                        ff% = FreeFile
                        Open sfile$ For Binary Access Write As #ff%
                        DoEvents
                    End If
                End If
                
                Put #ff%, , CStr(sarr$(i&) + sep$ + String(4 - Len(CStr(j&)), "0") + CStr(j&)) + CStr(IIf(mnuFileTypeUnix.Checked, Chr(10), vbCrLf))
                icount& = icount& + 1
            Next j&
        Next i&
    Close #ff%
    
    lst.Visible = True
    prog 0
    stat "saved to '" + scity$ + ".txt'"
    
    Erase sarr$()
    
    Select Case LANGUAGE$
    Case "english"
        s$ = "phone numbers for the city '" + scity$ + "' have been saved to '" + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + _
             scity$ + ".txt'." + vbCrLf + vbCrLf + "do you want to open the containing folder?"
    Case "french"
        s$ = "numéros de téléphone pour la ville '" + scity$ + "' ont été enregistrées dans '" + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + _
             scity$ + ".txt'." + vbCrLf + vbCrLf + "Voulez-vous ouvrez le dossier contenant?"
    Case "german"
        s$ = "Telefonnummern für die Stadt '" + scity$ + "' gespeichert wurden, um '" + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + _
             scity$ + ".txt'." + vbCrLf + vbCrLf + "wollen Sie mit dem Ordner zu öffnen?"
    Case "spanish"
        s$ = "números de teléfono de la ciudad '" + scity$ + "' se han guardado en '" + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + _
             scity$ + ".txt'." + vbCrLf + vbCrLf + "¿quieres abrir la carpeta que contiene?"
    End Select
    If frmMsg.MsgBocks(s$, vbQuestion + vbYesNo, "L517") = vbYes Then
        Shell "explorer /select," + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + scity$ + ".txt", vbNormalFocus
    End If
End Sub

Private Sub mnuGenPhoneSep_Click()
    Dim s$, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter character or string to separate the area code, prefix, and 4 digits number in the phone number:" + vbCrLf + vbCrLf + _
              "note: pressing cancel will default to a null separator!"
    Case "french"
        sl$ = "Entrez caractère ou une chaîne pour séparer l'indicatif régional, préfixe et numéro de 4 chiffres dans le numéro de téléphone:" + vbCrLf + vbCrLf + _
              "Remarque: appuyer sur Annuler par défaut à un séparateur null!"
    Case "german"
        sl$ = "Zeichen oder eine Zeichenkette eingeben, um die Telefonnummer mit Vorwahl-Präfix zu trennen, und 4-stellig Zahl in der Telefonnummer:" + vbCrLf + vbCrLf + _
              "Hinweis: Druck auf beenden wird standardmäßig auf einen Null-Trennzeichen!"
    Case "spanish"
        sl$ = "entrar en cadena de caracteres o para separar el código de área, prefijo y el número de 4 dígitos en el número de teléfono:" + vbCrLf + vbCrLf + _
              "Nota: Cancelar presionando de forma predeterminada a un separador null!"
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517", regGet("sep_phone"))
    
    regSet "sep_phone", s$
    mnuGenPhoneSep.Caption = "separator: " + IIf(s$ = "", "[none]", s$)
    mnuGenPhoneSep.Tag = s$
    mnuGenPhoneArea.Caption = "[areacode]" + s$ + "[prefix]" + s$ + "####"
    mnuGenPhoneNoarea.Caption = "[prefix]" + s$ + "####"
    
    If LANGUAGE$ <> "english" Then Change_Language
End Sub

Private Sub mnuGenStringHelp_Click()
    Dim s$
    
    Select Case LANGUAGE$
    Case "english"
        s$ = "each item in the menu contains a specific charset - a charset is a string of letters, numbers, and/or symbols that are used by this program to generate words." + vbCrLf + vbCrLf
        s$ = s$ + "select the charset you want to generate with from the list." + vbCrLf + vbCrLf
        s$ = s$ + "after you click on a charset, you will be prompted to enter the length of the string you want to generate." + vbCrLf
        s$ = s$ + "the program will then generate every possible combination of the charset that is of the specified length." + vbCrLf + vbCrLf
        s$ = s$ + "the charset.lst file is editable and contained in the same directory as this program.  would you like to view the charset.lst file?"
    Case "french"
             s$ = "chaque élément dans le menu contient un jeu de caractères spécifique - un jeu de caractères est une chaîne de lettres, de chiffres et / ou les symboles qui sont utilisés par ce programme pour générer des mots." + vbCrLf
        s$ = s$ + "Sélectionnez le jeu de caractères que vous souhaitez générer à partir de la liste." + vbCrLf
        s$ = s$ + "Après avoir cliqué sur un jeu de caractères, vous serez invité à entrer la longueur de la chaîne que vous voulez générer." + vbCrLf
        s$ = s$ + "le programme va alors générer toutes les combinaisons possibles du jeu de caractères qui est de la longueur spécifiée." + vbCrLf
        s$ = s$ + "charset.lst le fichier est éditable et contenues dans le même répertoire que ce programme. Souhaitez-vous afficher le fichier charset.lst?" + vbCrLf
    Case "german"
             s$ = "jedes Element in dem Menü enthält eine spezifische charset - ein Zeichensatz ist eine Zeichenfolge aus Buchstaben, Zahlen und / oder Symbole, die von diesem Programm, um Wörter zu erzeugen verwendet werden." + vbCrLf
        s$ = s$ + "Wählen Sie den Zeichensatz, die Sie mit von der Liste zu generieren." + vbCrLf
        s$ = s$ + "nachdem Sie auf einen Zeichensatz, werden Sie aufgefordert werden, um die Länge der Zeichenfolge, die Sie erzeugen wollen geben." + vbCrLf
        s$ = s$ + "das Programm generiert dann jede mögliche Kombination von den Zeichensatz, dass mit der angegebenen Länge ist." + vbCrLf
        s$ = s$ + "charset.lst der Datei ist leicht zu bearbeiten und die in dem gleichen Verzeichnis wie dieses Programm. Möchten Sie die charset.lst Datei anzuzeigen?" + vbCrLf
    Case "spanish"
             s$ = "cada elemento en el menú contiene un conjunto de caracteres específicos - un conjunto de caracteres es una cadena de letras, números y / o símbolos que se utilizan en este programa para generar palabras." + vbCrLf
        s$ = s$ + "seleccionar el juego de caracteres que desea generar con la de la lista." + vbCrLf
        s$ = s$ + "Después de hacer clic en un juego de caracteres, se le pedirá que introduzca la longitud de la cadena que desea generar." + vbCrLf
        s$ = s$ + "Entonces, el programa va a generar todas las combinaciones posibles del juego de caracteres que es de la longitud especificada." + vbCrLf
        s$ = s$ + "charset.lst el archivo es editable y que figuran en el mismo directorio que este programa. ¿Te gustaría ver el archivo charset.lst?" + vbCrLf
    End Select
    
    If frmMsg.MsgBocks(s$, vbYesNo + vbInformation, "L517") = vbYes Then
        Shell "explorer " + Chr(34) + App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + "charset.lst", vbNormalFocus
    End If
End Sub

Private Sub mnuGenStringX_Click(Index As Integer)
    Dim a$, ff%, f$, slen$, lword&(), i&, j&, s$, per_ttl#, per_cur#, isplit&, icount&, timah#, xtemp#, sl$
    Dim sarr$()
    
    icount& = 0
    isplit& = 0
    s$ = regGet("split")
    If IsNumeric(s$) = True Then isplit& = CLng(s$)
    
    s$ = regGet("1mb")
    If s$ = "" Or IsNumeric(s$) = False Then
        'need to run timing check
        
        stat "checking data-write speed..."
        DoEvents
        Do While Len(Dir(App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + "_test.txt")) <> 0
            Kill App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + "_test.txt"
            DoEvents
        Loop
        ff% = FreeFile
        Open App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + "_test.txt" For Binary Access Write As #ff%
            timah# = Timer
            For i& = 0 To (1024 * (1024 / 8))
                If i Mod 256 = 0 Then DoEvents
                Put #ff, , CStr(String(IIf(mnuFileTypeUnix.Checked, 7, 6), "0") + IIf(mnuFileTypeUnix.Checked, Chr(10), vbCrLf))
            Next i&
            timah# = Timer - timah#
        Close #ff%
        Do While Len(Dir(App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + "_test.txt")) <> 0
            Kill App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + "_test.txt"
            DoEvents
        Loop
        regSet "1mb", CStr(timah#)
        stat ""
    Else
        timah# = CDbl(s$)
    End If
    
    a$ = mnuGenStringX(Index).Tag
    
    If RESUME_STRING$ = "" Then
        Select Case LANGUAGE$
        Case "english"
            sl$ = "enter the (numeric) length of the string you want to generate from the charset:"
        Case "french"
            sl$ = "entrez le (numérique) longueur de la chaîne que vous voulez générer à partir du jeu de caractères:"
        Case "german"
            sl$ = "die (numerisch) Länge der Zeichenfolge, die Sie aus den Zeichensatz zu erzeugen:"
        Case "spanish"
            sl$ = "entrar en el (numérico) de longitud de la cadena que desea generar en el conjunto de caracteres:"
        End Select
        slen$ = frmMsg.InputBocks(sl$ + " '" + mnuGenStringX(Index).Caption + "':")
        If slen$ = "" Or IsNumeric(slen$) = False Then Exit Sub
        
        per_ttl# = CDbl(Len(a$) ^ CDbl(slen$))
        
        xtemp# = per_ttl# * (CLng(slen$) + IIf(frmMain.mnuFileTypeUnix.Checked, 1, 2))
        
        
        ReDim lword&(CLng(slen$) - 1)
        For i& = 0 To UBound(lword&())
            lword(i&) = 0&
        Next i&
        
        Select Case LANGUAGE$
        Case "english"
                 s$ = "character set: " + vbTab + a$ + vbCrLf
            s$ = s$ + "estimated size: " + vbTab + BytesToString$(xtemp#) + vbCrLf
            s$ = s$ + "estimated time: " + vbTab + CalcETA(xtemp#, timah#) + vbCrLf
            s$ = s$ + "filename: " + vbTab + vbTab + "'" + slen$ + "-" + mnuGenStringX(Index).Caption + ".txt'" + vbCrLf + vbCrLf
            s$ = s$ + "do you want to generate this list?"
        Case "french"
                 s$ = "ensemble de lettres: " + vbTab + a$ + vbCrLf
            s$ = s$ + "estimé l'espace disque dur: " + vbTab + BytesToString$(xtemp#) + vbCrLf
            s$ = s$ + "Estimation du temps restant: " + vbTab + CalcETA(xtemp#, timah#) + vbCrLf
            s$ = s$ + "le nom du fichier: " + vbTab + vbTab + "'" + slen$ + "-" + mnuGenStringX(Index).Caption + ".txt'" + vbCrLf + vbCrLf
            s$ = s$ + "Voulez-vous générer cette liste?"
        Case "german"
                 s$ = "Reihe von Briefen: " + vbTab + a$ + vbCrLf
            s$ = s$ + "geschätzten Platz auf der Festplatte: " + vbTab + BytesToString$(xtemp#) + vbCrLf
            s$ = s$ + "Geschätzte verbleibende Zeit: " + vbTab + CalcETA(xtemp#, timah#) + vbCrLf
            s$ = s$ + "Dateinamen: " + vbTab + vbTab + "'" + slen$ + "-" + mnuGenStringX(Index).Caption + ".txt'" + vbCrLf + vbCrLf
            s$ = s$ + "wollen Sie diese Liste zu generieren?"
        Case "spanish"
                 s$ = "conjunto de cartas: " + vbTab + a$ + vbCrLf
            s$ = s$ + "Estimación del espacio de disco duro: " + vbTab + BytesToString$(xtemp#) + vbCrLf
            s$ = s$ + "tiempo restante estimado: " + vbTab + CalcETA(xtemp#, timah#) + vbCrLf
            s$ = s$ + "el nombre de archivo: " + vbTab + vbTab + "'" + slen$ + "-" + mnuGenStringX(Index).Caption + ".txt'" + vbCrLf + vbCrLf
            s$ = s$ + "¿quieres generar esta lista?"
        End Select
        
        If frmMsg.MsgBocks(s$, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        per_cur# = 1
        f$ = App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + slen$ + "-" + mnuGenStringX(Index).Caption + ".txt"
        Do While Len(Dir(f$)) <> 0
            Kill f$
            DoEvents
        Loop
        
    Else
        ' resume from previous list generation
        sarr$() = Split(RESUME_STRING$, ",")
        slen$ = CStr(CLng(UBound(sarr$()) + 1))
        ReDim lword&(UBound(sarr$()))
        For i& = 0 To UBound(lword&())
            lword&(i&) = CLng(sarr$(i&))
        Next i&
        
        per_ttl# = CDbl(Len(a$) ^ CDbl(slen$))
        xtemp# = per_ttl# * (CLng(slen$) + IIf(frmMain.mnuFileTypeUnix.Checked, 1, 2))
        
        If IsNumeric(regGet("resume_percent")) Then
            per_cur# = CDbl(regGet("resume_percent"))
        Else
            per_cur# = 1
        End If
        
        f$ = App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + slen$ + "-" + mnuGenStringX(Index).Caption + ".txt"
        Do While Len(Dir(f$)) <> 0
            DoEvents
            f$ = NextFile$(f$)
        Loop
    End If
    
    stat "generating file..."
    prog 0.001
    lst.Visible = False
    lblCancel.Visible = True
    DoEvents
    
    ff% = FreeFile
    Open f$ For Binary Access Write As #ff%
        Do
            If per_cur# Mod 2500 = 0 Then
                DoEvents
                If lblCancel.Visible = False Then
                    'cancel was clicked
                    per_ttl# = -1
                    GoTo cancelled
                End If
                prog per_cur# / per_ttl#
            End If
            
            'split check
            If isplit& > 0 Then
                If icount& >= isplit& Then
                    DoEvents
                    icount& = 0
                    'close old file
                    Close #ff%
                    'find next file name in sequence
                    f$ = NextFile$(f$)
                    'open new file
                    ff% = FreeFile
                    Open f$ For Binary Access Write As #ff%
                    DoEvents
                End If
            End If
            
            'build word from numeric array
            s$ = ""
            For i& = 0 To UBound(lword&())
                s$ = s$ + Mid(a$, lword&(i&) + 1, 1)
            Next i&
            
            If Len(s$) <> CInt(slen$) Then Exit Do
            
            Put #ff%, , CStr(s$ + IIf(mnuFileTypeUnix.Checked, Chr(10), vbCrLf))
            icount& = icount& + 1
            
            per_cur# = per_cur# + 1
            
            'increment the array
            For i& = 0 To UBound(lword&())
                If lword&(i&) + 1 < Len(a$) Then
                    lword&(i&) = lword&(i&) + 1
                    For j& = 0 To i& - 1
                        lword&(j&) = 0
                    Next j&
                    i& = -1
                    Exit For
                End If
            Next i&
        Loop Until i& <> -1
cancelled:
    If lblCancel.Visible = False Then
        If frmMsg.MsgBocks("Do you want to save your progress and start from the same point next time you load L517?" + vbCrLf + vbCrLf + "L517 can resume wordlist generation at a later time.  By selecting 'Yes', L517 will ask if you want to resume generating this list the next time the program is launched.", vbQuestion + vbYesNo, "L517") = vbYes Then
            s$ = ""
            For i = 0 To UBound(lword())
                s$ = s$ + CStr(lword(i)) + ","
            Next i
            s$ = Left(s$, Len(s$) - 1)
            regSet "resume", s$
            regSet "resume_index", CStr(Index)
            regSet "resume_percent", CStr(per_cur#)
        End If
    End If
    Close #ff%
    DoEvents
    prog 0
    stat ""
    lst.Visible = True
    
    Select Case LANGUAGE$
    Case "english"
        If per_ttl# = -1 Then
            s$ = "list generation was cancelled." + vbCrLf + vbCrLf
        Else
            s$ = "list generation is complete." + vbCrLf + vbCrLf
        End If
        s$ = s$ + "do you want to open the folder containing the generated file"
    Case "french"
        If per_ttl# = -1 Then
            s$ = "la génération de listes a été annulée." + vbCrLf + vbCrLf
        Else
            s$ = "génération liste est complète." + vbCrLf + vbCrLf
        End If
        s$ = s$ + "Voulez-vous ouvrez le dossier contenant le fichier généré"
    Case "german"
        If per_ttl# = -1 Then
            s$ = "Liste Generation wurde abgesagt." + vbCrLf + vbCrLf
        Else
            s$ = "Liste Generation abgeschlossen ist." + vbCrLf + vbCrLf
        End If
        s$ = s$ + "wollen Sie den Ordner mit der erzeugten Datei öffnen"
    Case "spanish"
        If per_ttl# = -1 Then
            s$ = "generación de la lista fue cancelada." + vbCrLf + vbCrLf
        Else
            s$ = "la generación de la lista está completa." + vbCrLf + vbCrLf
        End If
        s$ = s$ + "¿quieres abrir la carpeta que contiene el archivo generado?"
    End Select
    
    s$ = s$ + IIf(isplit& > 0, "(s)", "") + ": '" + GetFileName$(f$) + "'?"
    If frmMsg.MsgBocks(s$, vbQuestion + vbYesNo) = vbYes Then
        Shell "explorer /select," + f$, vbNormalFocus
    End If
End Sub

Private Sub mnuGenWeb_Click()
    frmWeb.Show vbModal, Me
    
    Exit Sub
    Dim sdata$, s$, sl$, lbefore&, lafter&
    
    s$ = frmMsg.InputBocks("enter the url or web address of the site you want to grab words from:" + vbCrLf + vbCrLf + "i.e. http://www.myspace.com/tilatequila", "L517", regGet("last_url"))
    If s$ = "" Then Exit Sub
    
    regSet "last_url", s$
    
    stat "loading site; please wait"
    DoEvents
    
    sdata$ = webgetsource$(s$)
    If sdata$ <> "" Then
        lbefore& = lst.ListItems().count
        ParseWebData sdata$
        lafter& = lst.ListItems().count
        stat "loaded site; " + CStr(lafter& - lbefore&) + " items found"
    Else
        stat "unable to load site data"
    End If
    
    B_CHANGE = True
    UpdateCaption
End Sub

Private Sub mnuHelpEnglish_Click()
    LANGUAGE$ = "english"
    regSet "lang", LANGUAGE$
    Change_Language
End Sub

Private Sub mnuHelpFrench_Click()
    LANGUAGE$ = "french"
    regSet "lang", LANGUAGE$
    Change_Language
End Sub

Private Sub mnuHelpGerman_Click()
    LANGUAGE$ = "german"
    regSet "lang", LANGUAGE$
    Change_Language
End Sub

Private Sub mnuHelpGetlists_Click()
    Dim s$
    
    Select Case LANGUAGE$
    Case "english"
        s$ = "do you want to download a zip file containing 1.1MB of wordlists?" + vbCrLf + vbCrLf + _
             "the zip has seven very good, common password lists."
    Case "french"
        s$ = "Voulez-vous télécharger un fichier zip contenant 1.1MB des listes de mots?" + vbCrLf + vbCrLf + _
             "le .ZIP a des listes de mot de passe sept très bons, ordinaires."
    Case "german"
        s$ = "wollen Sie ein Zip-Download-Datei mit 1,1 MB von Wortlisten?" + vbCrLf + vbCrLf + _
             "die .ZIP hat sieben sehr gute, gemeinsame Passwort-Listen."
    Case "spanish"
        s$ = "¿quieres descargar un archivo zip que contiene 1.1MB de listas de palabras?" + vbCrLf + vbCrLf + _
             "de la .ZIP tiene siete muy buenas, listas de contraseñas comunes."
    End Select
    If frmMsg.MsgBocks(s$, vbYesNo + vbQuestion) = vbYes Then
        Shell "explorer " + Chr(34) + "http://l517.googlecode.com/files/Common%20Lists.zip" + Chr(34), vbNormalFocus
    End If
End Sub

Private Sub mnuHelpHelp_Click()
    Dim s$, timah#, readme$, ff%
    
    Select Case LANGUAGE$
    Case "english"
        s$ = "L517 : wordlist generator : readme" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "shortcuts" + vbCrLf
        s$ = s$ + "----------" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "1. drag-and-drop file(s) anywhere on this application to add" + vbCrLf
        s$ = s$ + "     files quickly." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "2. you can also drag-and-drop selected text (from files," + vbCrLf
        s$ = s$ + "     emails, webpages) into the list." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "3. you can PASTE text into the list! press ctrl+V anywhere" + vbCrLf
        s$ = s$ + "     in the application." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "4. double-click titlebar to 'shrink' the program." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "5. the menu has lots of shortcuts that work anywhere in " + vbCrLf
        s$ = s$ + "     the program:" + vbCrLf
        s$ = s$ + "" + vbTab + "ctrl+O     : open file" + vbCrLf
        s$ = s$ + "" + vbTab + "ctrl+V     : load words from clipboard" + vbCrLf
        s$ = s$ + "" + vbTab + "ctrL+D     : dupe-kill list" + vbCrLf
        s$ = s$ + "" + vbTab + "ctrl+F     : find item in list" + vbCrLf
        s$ = s$ + "" + vbTab + "Delete     : remove selected item" + vbCrLf
        s$ = s$ + "" + vbTab + "Shift+Del  : clear list" + vbCrLf
        s$ = s$ + "" + vbTab + "...and then some." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "menu" + vbCrLf
        s$ = s$ + "----------" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[FILE]" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">new" + vbTab + vbTab + " clears list." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">open" + vbTab + "" + vbTab + " opens wordlist file, adds each found word " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + " to the list. (handles -almost- all file " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + " types/formats)" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">save as..." + vbTab + " saves wordlist to file." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">split into file saves list to a new file every X items." + vbCrLf
        s$ = s$ + "" + vbTab + vbTab + " uses sequential file-naming '1.txt, 2.txt'" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">save type>" + vbCrLf
        s$ = s$ + "" + vbTab + ">windows adds carriage-return to end of lines; saves" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + " files that work with windows applications" + vbCrLf
        s$ = s$ + "" + vbTab + ">unix" + vbTab + " only uses line-feed at end of lines; works" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + " with unix [and some windows] applications" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">presets>" + vbCrLf
        s$ = s$ + "" + vbCrLf + "[note: these presets change the *filter length* for passwords]" + vbCrLf
        s$ = s$ + "" + vbTab + ">wpa " + vbTab + " sets minimum/maximum length for  " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + " wpa passwords (8-64)" + vbCrLf
        s$ = s$ + "" + vbTab + ">web " + vbTab + " minimum/maximum length for " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + " web-based passwords (4-12) + vbCrLf" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">exit" + vbTab + "" + vbTab + " quits." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[EDIT]" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">remove duplicates" + vbTab + "removes duplicate items." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">clear list" + vbTab + "" + vbTab + "clears list (after a prompt)." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">remove item" + vbTab + "" + vbTab + "removes selected item, same as " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "pressing delete key." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">remove items with" + vbTab + "removes items that DO contain a " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "string of text." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">remove items without" + vbTab + "removes items that DON'T contain a " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "string of text." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">find" + vbTab + "" + vbTab + "" + vbTab + "searches for word [or part of word]" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "in list." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">find next" + vbTab + "" + vbTab + "finds next occurrence of previously" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "-searched word." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[GENERATE]" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">words from website" + vbTab + "extracts words from website, " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  filtering out HTML. (BUG: LAGS)" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "for a simpler method, visit the site " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  in IE/FireFox, select the entire " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  webpage (ctrl+A)," + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  then click-and-drag the text into" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  this program's list." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">words from folder(s)" + vbTab + "scans directory (& subdirectories). " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  extracts words from jpg, mp3, doc," + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  srt, and others." + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "excludes: rar, avi, mpg, mp4, iso, " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "  " + vbTab + "  zip, msi, exe, .torrent, dat, tar" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">string from charset" + vbTab + "generates string of desired length" + vbCrLf
        s$ = s$ + "" + vbTab + vbTab + vbTab + "  based on 'charset.lst'," + vbCrLf
        s$ = s$ + vbTab + vbTab + vbTab + "  a character-set file." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">dates>" + vbCrLf
        s$ = s$ + "" + vbTab + ">separator" + vbTab + "decides what the months, days and" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "years are separated by. can be null." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbTab + ">[mm/dd/yy]" + vbTab + "user sets the start and stop year" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + " and the program generates, based on" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + " the format, all the dates within" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + " that time period." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">phone numbers>" + vbCrLf
        s$ = s$ + "" + vbTab + ">separator" + vbTab + "what character/string separates the " + vbCrLf
        s$ = s$ + vbTab + vbTab + vbTab + "areacode, prefix, and 4 digit number" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + vbTab + ">[areacode][prefix]####" + vbTab + "generates every phone number" + vbCrLf
        s$ = s$ + vbTab + vbTab + vbTab + vbTab + "within a specific city/state." + vbCrLf
        s$ = s$ + vbTab + vbTab + vbTab + vbTab + "saves to file-too big for list" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + vbTab + ">[prefix]####" + vbTab + vbTab + "same as above" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "options" + vbCrLf
        s$ = s$ + "---------" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[CASE]" + vbCrLf
        s$ = s$ + "convert to:" + vbCrLf
        s$ = s$ + "  >lowercase" + vbTab + vbTab + "every item in list is changed to Lcase" + vbCrLf
        s$ = s$ + "  >UPPERCASE" + vbTab + vbTab + "every item is converted to upper-case" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "make copies with:" + vbCrLf
        s$ = s$ + "  >First Letter Upper" + vbTab + "saves original list (as shown), plus " + vbCrLf + vbTab + vbTab + vbTab + "  list is converted to 1stLetterUpper" + vbCrLf
        s$ = s$ + "  >EvErY oThEr UpPeR" + vbTab + "saves original format + every other" + vbCrLf + vbTab + vbTab + vbTab + "   uppercase format" + vbCrLf
        s$ = s$ + "  >1337C453" + vbTab + vbTab + "saves the list, and also saves the" + vbCrLf + vbTab + vbTab + vbTab + "  list converted to 'leetspeak'" + vbCrLf
        s$ = s$ + vbTab + vbTab + vbTab + "leet dictionary is in 'leetspeak.txt'," + vbCrLf + vbTab + vbTab + vbTab + "  and is editable!" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[FILTER]" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">set min/max length" + vbTab + "removes (and will not add) words" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  smaller/greater than the min/max." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">text to right/left..." + vbTab + "splits word from each line " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  in a file." + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "can be useful for a lot of reasons. " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "i.e.:" + vbCrLf
        s$ = s$ + " " + vbTab + "" + vbTab + "" + vbTab + "supposed a wordlist is numbered: " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "" + vbTab + "1. word1" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "" + vbTab + "2. word2" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "" + vbTab + "3. word3" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "to extract just the words (no #'s)," + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "you would select 'text to right of'" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "and enter '. ' without quotes." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">convert !@#...to hex" + vbTab + "this will change special (non alpha-" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  numeric) characters to their hex" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  equivalent. i.e. %21%40%22%23" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "useful when building a wordlist that" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  will be used over the web." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">include foreign chars" + vbTab + "allows more than just alpha-numeric" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  words, includes accented symbols" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  like â é ï etc" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[APPEND]" + vbCrLf
        s$ = s$ + ">add to left of item>" + vbTab + "prepends prefix to every item in list." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbTab + ">custom list" + vbTab + "user selects file, each item from the " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  list is added to the beginning of " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  each item in the list." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbTab + ">def. prefixes" + vbTab + "program uses default, chosen-by-author" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + " prefixes, adds them to each item in " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + " the list." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbTab + ">numeric" + vbTab + "adds 0-9 to beginning of each item in " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  the list.  Only does one number" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  at-a-time." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbTab + ">alpha" + vbTab + "" + vbTab + "adds a-z to beginning of each item in " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  the list.  Only does one letter" + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  at-a-time." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + ">add to right of item>" + vbTab + "appends postfix to every item in the " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  list. (same as above, but adds to " + vbCrLf
        s$ = s$ + "" + vbTab + "" + vbTab + "" + vbTab + "  the END of each item)." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "usage" + vbCrLf
        s$ = s$ + "--------" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "the options in L517 are designed to provide the tools " + vbCrLf
        s$ = s$ + " necessary to build the 'perfect wordlist'." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "this is a re-write and expansion upon an earlier program," + vbCrLf
        s$ = s$ + " 'LST', which was slow, cumbersome, and lacked the options and " + vbCrLf
        s$ = s$ + " features that I wanted to use." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "L517 is small, lightweight, and does its best at generating" + vbCrLf + "  wordlists." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "some lists will take longer to load than others.  some options " + vbCrLf
        s$ = s$ + " (words-from-website) can lock the program up for minutes at a time." + vbCrLf
        s$ = s$ + " dupekilling large lists can take HOURS." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "my point is: this program is stable." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "godspeed, and may this program generate a long-sought-after" + vbCrLf
        s$ = s$ + " password for you." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "-derv" + vbCrLf
        s$ = s$ + "" + vbCrLf
    Case "french"
        s$ = "L517: générateur de liste de mots: lisez-moi" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "raccourcis" + vbCrLf
        s$ = s$ + "----------" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "1. glisser-déposer une goutte (s) n'importe où sur cette demande d'ajouter" + vbCrLf
        s$ = s$ + "     fichiers rapidement." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "2. vous pouvez également glisser-déposer du texte sélectionné (à partir de fichiers," + vbCrLf
        s$ = s$ + "     e-mails, pages Web) dans la liste." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "3. Vous pouvez coller du texte dans la liste! appuyez sur ctrl V partout" + vbCrLf
        s$ = s$ + "     dans la demande." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "4. barre de titre de double-cliquer sur 'shrink' le programme." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "5. Le menu a beaucoup de raccourcis que de travailler partout dans" + vbCrLf
        s$ = s$ + "     Au programme:" + vbCrLf
        s$ = s$ + "Ctrl O: Ouvrir le fichier" + vbCrLf
        s$ = s$ + "Ctrl V: les mots de charge de la planchette" + vbCrLf
        s$ = s$ + "Ctrl D: Dupe-kill liste" + vbCrLf
        s$ = s$ + "F Ctrl: trouver l'élément dans la liste" + vbCrLf
        s$ = s$ + "Supprimer: Supprimer l'élément sélectionné" + vbCrLf
        s$ = s$ + "Shift Del: liste claire" + vbCrLf
        s$ = s$ + "... et puis certains." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "menu" + vbCrLf
        s$ = s$ + "----------" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[DOSSIER]" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> new efface liste." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Ouvrir ouvre le fichier base de mots, ajoute chaque retrouve mot" + vbCrLf
        s$ = s$ + "à la liste. (poignées-presque-tous les fichiers" + vbCrLf
        s$ = s$ + "types / formats)" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> enregistrer sous ... wordlist enregistre dans un fichier." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> split dans le fichier enregistre la liste à un nouveau fichier de tous les articles X." + vbCrLf
        s$ = s$ + "utilise un fichier séquentiel de nommage «1. txt, 2.txt '" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> save type>" + vbCrLf
        s$ = s$ + "> Windows ajoute le transport-retour à la fin des lignes; sauve" + vbCrLf
        s$ = s$ + "des fichiers qui fonctionnent avec les applications Windows" + vbCrLf
        s$ = s$ + "> UNIX utilise uniquement line-feed en fin de ligne, les œuvres" + vbCrLf
        s$ = s$ + "sous Unix [et certaines fenêtres] applications" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Presets>" + vbCrLf
        s$ = s$ + "> WPA définit l'minimal adéquat / longueur maximale" + vbCrLf
        s$ = s$ + "des mots de passe WPA (8-64)" + vbCrLf
        s$ = s$ + "> web ensembles min bon / Longueur max pour les web-based" + vbCrLf
        s$ = s$ + "mots de passe (4-12) vbCrLf" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> sortie se ferme." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[EDIT]" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> supprimer les doublons supprime les doublons." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> liste claire efface la liste (après une invite)." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Supprimer article supprime l'élément sélectionné, même chose que" + vbCrLf
        s$ = s$ + "appuyant sur Suppr." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> supprimer des éléments à supprimer les articles qui ne contiennent une" + vbCrLf
        s$ = s$ + "chaîne de texte." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Supprimer des éléments sans supprime des éléments qui ne contiennent pas une" + vbCrLf
        s$ = s$ + "chaîne de texte." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> TROUVER recherches pour mot [ou partie du mot]" + vbCrLf
        s$ = s$ + "dans la liste." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> TROUVER prochaines trouve prochaine occurrence précédemment" + vbCrLf
        s$ = s$ + "-mot recherché." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[Produire]" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> mots de mots extraits du site web site web," + vbCrLf
        s$ = s$ + "filtrer HTML. (BUG: GAL)" + vbCrLf
        s$ = s$ + "pour une méthode plus simple, visitez le site" + vbCrLf
        s$ = s$ + "dans IE / Firefox, sélectionnez l'ensemble du" + vbCrLf
        s$ = s$ + "page (ctrl)," + vbCrLf
        s$ = s$ + "puis cliquer-glisser le texte dans" + vbCrLf
        s$ = s$ + "liste de ce programme." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "mots> de dossier (s) balaie répertoire (et sous-répertoires)." + vbCrLf
        s$ = s$ + "extraits de mots à partir jpg, mp3, doc," + vbCrLf
        s$ = s$ + "srt, et d'autres." + vbCrLf
        s$ = s$ + "ne comprend pas: avi, avi, mpg, mp4, ISO," + vbCrLf
        s$ = s$ + "zip, msi, exe,. torrent, dat, goudron" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> string charset génère de chaîne d'une longueur désirée" + vbCrLf
        s$ = s$ + "fondée sur la «charset.lst '," + vbCrLf
        s$ = s$ + "un personnage-ensemble du fichier." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> dates>" + vbCrLf
        s$ = s$ + "> séparateur décide ce que les mois, les jours et" + vbCrLf
        s$ = s$ + "ans sont séparés par. peut être nul." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> [mm / jj / aa] utilisateur fixe le début et de fin année" + vbCrLf
        s$ = s$ + "et le programme génère, sur la base" + vbCrLf
        s$ = s$ + "le format, toutes les dates dans" + vbCrLf
        s$ = s$ + "cette période." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Numéros de téléphone>" + vbCrLf
        s$ = s$ + "> séparateur quel caractère / chaîne sépare le" + vbCrLf
        s$ = s$ + "indicatif régional, préfixe et numéro à 4 chiffres" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> [indicatif régional] []#### préfixe génère chaque numéro de téléphone" + vbCrLf
        s$ = s$ + "au sein d'une ville spécifique ou des États." + vbCrLf
        s$ = s$ + "enregistre à file-trop gros pour la liste" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> []#### même préfixe comme ci-dessus" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "options" + vbCrLf
        s$ = s$ + "---------" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[CAS]" + vbCrLf
        s$ = s$ + "Convert to:" + vbCrLf
        s$ = s$ + "  > minuscules chaque élément dans la liste est modifiée à Lcase" + vbCrLf
        s$ = s$ + "  > MAJUSCULES chaque article est converti en majuscules" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "faire des copies avec:" + vbCrLf
        s$ = s$ + "  > Première Lettre Haute sauve liste originale (comme indiqué), plus" + vbCrLf
        s$ = s$ + "liste est converti en 1stLetterUpper" + vbCrLf
        s$ = s$ + "  > Haute autre chaque format d'origine enregistre tous les autres" + vbCrLf
        s$ = s$ + "majuscules format" + vbCrLf
        s$ = s$ + "  > 1337C453 enregistre la liste, et permet également d'économiser les" + vbCrLf
        s$ = s$ + "liste converti en «Leetspeak '" + vbCrLf
        s$ = s$ + "Dictionnaire Leet est en «leetspeak.txt '," + vbCrLf
        s$ = s$ + "et est modifiable!" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[Filtre]" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> min set / longueur max supprime (et ne pourra pas ajouter) les mots" + vbCrLf
        s$ = s$ + "plus petite / plus grande que le min / max." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Texte à droite / à gauche ... scissions mot de chaque ligne" + vbCrLf
        s$ = s$ + "dans un fichier." + vbCrLf
        s$ = s$ + "peut être utile pour beaucoup de raisons." + vbCrLf
        s$ = s$ + "à savoir:" + vbCrLf
        s$ = s$ + " supposait une liste de mots est numérotée:" + vbCrLf
        s$ = s$ + "1. mot1" + vbCrLf
        s$ = s$ + "2. mot2" + vbCrLf
        s$ = s$ + "3. mot3" + vbCrLf
        s$ = s$ + "pour extraire simplement les mots (pas de # 's)," + vbCrLf
        s$ = s$ + "vous devez sélectionner «texte à droite de" + vbCrLf
        s$ = s$ + "et entrez . Sans les guillemets. + vbCrLf"
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> !@#... convertir en hexadécimal, cela va changer alpha (Special non -" + vbCrLf
        s$ = s$ + "numérique) à leurs caractères hex" + vbCrLf
        s$ = s$ + "équivalent. à savoir! @ #" + vbCrLf
        s$ = s$ + "utile lors de la construction d'une liste de mots que" + vbCrLf
        s$ = s$ + "sera utilisé sur le Web." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> inclure les caractères étrangers permet plus que l'alpha-numérique" + vbCrLf
        s$ = s$ + "mots, comprend accentués symboles" + vbCrLf
        s$ = s$ + "comme a i E, etc" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[APPEND]" + vbCrLf
        s$ = s$ + "> ajouter à gauche du texte:> Ajoute préfixe à chaque élément de liste." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> user liste personnalisée sélectionne le fichier, chaque élément de la" + vbCrLf
        s$ = s$ + "liste est ajoutée au début de" + vbCrLf
        s$ = s$ + "chaque élément de la liste." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> def. préfixes programme utilise par défaut, choisi par l'auteur" + vbCrLf
        s$ = s$ + "préfixes, les ajoute à chaque élément de" + vbCrLf
        s$ = s$ + "la liste." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> numérique ajoute 0-9 pour début de chaque élément dans" + vbCrLf
        s$ = s$ + "la liste. Ne, un seul numéro" + vbCrLf
        s$ = s$ + "AT-A-temps." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> alpha ajoute une-z au début de chaque élément dans" + vbCrLf
        s$ = s$ + "la liste. Ne fait qu'une seule lettre" + vbCrLf
        s$ = s$ + "AT-A-temps." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> ajouter à droite du point> ajoute Postfix pour chaque élément de la" + vbCrLf
        s$ = s$ + "liste. (le même que ci-dessus, mais il ajoute à" + vbCrLf
        s$ = s$ + "la fin de chaque article)." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "utilisation" + vbCrLf
        s$ = s$ + "--------" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "les options dans L517 sont conçus pour fournir les outils" + vbCrLf
        s$ = s$ + " nécessaire pour construire la «liste de mots parfait»." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "Il s'agit d'une ré-écriture et de l'expansion sur un programme antérieur," + vbCrLf
        s$ = s$ + " «LST», qui était lente, lourde, et ne disposaient pas des options et" + vbCrLf
        s$ = s$ + " caractéristiques que je voulais utiliser." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "L517 est petit, léger et fait de son mieux à générer des" + vbCrLf
        s$ = s$ + "  des listes de mots." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "certaines listes prendra plus de temps à charger que d'autres. quelques options" + vbCrLf
        s$ = s$ + " (mots-from-site) peut verrouiller le programme en place pendant quelques minutes à la fois." + vbCrLf
        s$ = s$ + " dupekilling grandes listes peuvent prendre des heures." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "Mon point est: ce programme est stable." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "Godspeed, mai ce programme et générer un long convoité" + vbCrLf
        s$ = s$ + " mot de passe pour vous." + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "-Derv" + vbCrLf
        s$ = s$ + "" + vbCrLf
    Case "spanish"
        s$ = "L517: generador de lista de palabras: readme " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "accesos directos " + vbCrLf
        s$ = s$ + "---------- " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "1. de arrastrar y soltar el archivo (s) en cualquier lugar de esta aplicación para agregar " + vbCrLf
        s$ = s$ + "     archivos de forma rápida. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "2. También puede arrastrar y soltar texto seleccionado (a partir de archivos, " + vbCrLf
        s$ = s$ + "     correos electrónicos, páginas web) en la lista. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "3. Puede pegar el texto en la lista! pulse Ctrl V en cualquier lugar " + vbCrLf
        s$ = s$ + "     en la aplicación. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "4. haga doble clic en barra de título a 'encoger' el programa. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "5. el menú tiene un montón de atajos que trabajar en cualquier lugar " + vbCrLf
        s$ = s$ + "     el programa: " + vbCrLf
        s$ = s$ + "Ctrl O: Abrir el archivo " + vbCrLf
        s$ = s$ + "Ctrl V: Palabras de carga desde el portapapeles " + vbCrLf
        s$ = s$ + "D ctrl: engañar, matar a la lista " + vbCrLf
        s$ = s$ + "F Ctrl: encontrar el artículo en la lista " + vbCrLf
        s$ = s$ + "Eliminar: quitar elemento seleccionado " + vbCrLf
        s$ = s$ + "Shift Supr: Borrar lista " + vbCrLf
        s$ = s$ + "... y algo más. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "menú " + vbCrLf
        s$ = s$ + "---------- " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[ARCHIVO] " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Nueva borra la lista. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Abra el archivo se abre lista de palabras, añade cada uno ha encontrado la palabra " + vbCrLf
        s$ = s$ + "a la lista. (administra-casi-todos los archivos " + vbCrLf
        s$ = s$ + "tipos de formatos) " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Guardar como ... guarda lista de palabras en un fichero. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> dividido en Guarda el archivo de lista a un nuevo archivo de todos los artículos de X. " + vbCrLf
        s$ = s$ + "utiliza nombres de archivo secuencial «1. txt, 2.txt ' " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Guardar tipo> " + vbCrLf
        s$ = s$ + "> Windows añade retorno de carro al final de las líneas; salva " + vbCrLf
        s$ = s$ + "los archivos que trabajan con las aplicaciones de Windows " + vbCrLf
        s$ = s$ + "> Unix sólo utiliza de línea de alimentación al final de las líneas, las obras " + vbCrLf
        s$ = s$ + "con Unix [y algunas ventanas] aplicaciones " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Presets> " + vbCrLf
        s$ = s$ + "> WPA establece el mínimo adecuado / longitud máxima " + vbCrLf
        s$ = s$ + "para las contraseñas WPA (8-64) " + vbCrLf
        s$ = s$ + "> web pone min adecuada / longitud máxima para los basados en la web " + vbCrLf
        s$ = s$ + "contraseñas (4-12) vbCrLf " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> salida se cierra. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[EDIT] " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> eliminar duplicados elimina los elementos duplicados. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Borrar lista borra la lista (después de un sistema). " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "item> quitar elimina seleccionados, al igual que " + vbCrLf
        s$ = s$ + "presionando la tecla Supr. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Eliminar elementos con elimina los elementos que contienen una " + vbCrLf
        s$ = s$ + "cadena de texto. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> eliminar elementos sin elimina los elementos que no contienen una " + vbCrLf
        s$ = s$ + "cadena de texto. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Buscar búsquedas por palabra [o parte de la palabra] " + vbCrLf
        s$ = s$ + "en la lista. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Buscar siguiente se encuentra la siguiente aparición de antes " + vbCrLf
        s$ = s$ + "-palabra buscada. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[Generar] " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "palabras> del sitio web de los extractos de las palabras de sitio web, " + vbCrLf
        s$ = s$ + "filtrado de HTML. (ERROR: GAL) " + vbCrLf
        s$ = s$ + "por un método más sencillo, visite el sitio " + vbCrLf
        s$ = s$ + "en IE, Firefox, seleccione todo el " + vbCrLf
        s$ = s$ + "página web (Ctrl A), " + vbCrLf
        s$ = s$ + "a continuación, haga clic y arrastre el texto a " + vbCrLf
        s$ = s$ + "la lista de este programa. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "palabras> de la carpeta (s) de exploraciones de directorio (y subdirectorios). " + vbCrLf
        s$ = s$ + "extractos de las palabras de jpg, mp3, doc, " + vbCrLf
        s$ = s$ + "SRT, y otros. " + vbCrLf
        s$ = s$ + "Excluye: wmv, avi, mpg, mp4, iso, " + vbCrLf
        s$ = s$ + "zip, MSI, exe,. torrente, dat, alquitrán " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> string de juego de caracteres genera cuerda de longitud deseada " + vbCrLf
        s$ = s$ + "charset.lst basado en, " + vbCrLf
        s$ = s$ + "un carácter de conjunto de archivos. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Fechas> " + vbCrLf
        s$ = s$ + "> Separador decide lo que los meses, días y " + vbCrLf
        s$ = s$ + "años están separados por. puede ser nulo. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> [dd / mm / aa] usuario establece el inicio y parada de año " + vbCrLf
        s$ = s$ + "y el programa genera, sobre la base de " + vbCrLf
        s$ = s$ + "el formato, todas las fechas dentro de " + vbCrLf
        s$ = s$ + "ese período de tiempo. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> números de teléfono> " + vbCrLf
        s$ = s$ + "> separador de qué personaje / cadena separa el " + vbCrLf
        s$ = s$ + "codigoDeArea, prefijo y número de 4 dígitos " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> [codigoDeArea] []#### prefijo genera cada número de teléfono " + vbCrLf
        s$ = s$ + "dentro de una determinada ciudad / estado. " + vbCrLf
        s$ = s$ + "guarda en archivo demasiado grande para la lista " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> []#### mismo prefijo que el anterior " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "opciones " + vbCrLf
        s$ = s$ + "--------- " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[CASE] " + vbCrLf
        s$ = s$ + "convertir a: " + vbCrLf
        s$ = s$ + "  > minúsculas cada elemento de la lista se cambia a Lcase " + vbCrLf
        s$ = s$ + "  > MAYÚSCULAS cada artículo se convierte a mayúsculas " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "hacer copias de: " + vbCrLf
        s$ = s$ + "  > Primera letra mayúscula guarda la lista original (como se muestra), más " + vbCrLf
        s$ = s$ + "la lista se convierte en 1stLetterUpper " + vbCrLf
        s$ = s$ + "  > Alta Todos los demás guarda el formato original de todos los demás " + vbCrLf
        s$ = s$ + "formato de mayúsculas " + vbCrLf
        s$ = s$ + "  > 1337C453 guarda la lista, y también guarda la " + vbCrLf
        s$ = s$ + "Lista de convertirse en 'leetspeak' " + vbCrLf
        s$ = s$ + "Leet diccionario se encuentra en 'leetspeak.txt', " + vbCrLf
        s$ = s$ + "y se puede editar! " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[Filter] " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> min set / longitud máxima quita (y no agrega) las palabras " + vbCrLf
        s$ = s$ + "menor / mayor que el min / max. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> texto a la derecha / izquierda ... divide palabra de cada línea de " + vbCrLf
        s$ = s$ + "en un archivo. " + vbCrLf
        s$ = s$ + "puede ser útil para un montón de razones. " + vbCrLf
        s$ = s$ + "es decir: " + vbCrLf
        s$ = s$ + " suponía una lista de palabras está numerada: " + vbCrLf
        s$ = s$ + "1. word1 " + vbCrLf
        s$ = s$ + "2. word2 " + vbCrLf
        s$ = s$ + "3. word3 " + vbCrLf
        s$ = s$ + "para extraer sólo las palabras (no # 's), " + vbCrLf
        s$ = s$ + "que usted seleccione 'texto a la derecha de' " + vbCrLf
        s$ = s$ + "y escriba '. ' Sin las comillas. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Convertir a hexadecimal !@#... esto va a cambiar especiales (alfa no " + vbCrLf
        s$ = s$ + "numérico) a sus caracteres hexadecimales " + vbCrLf
        s$ = s$ + "equivalente. es decir,! @ # " + vbCrLf
        s$ = s$ + "útil cuando se construye una lista de palabras que " + vbCrLf
        s$ = s$ + "se utilizará en la web. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> permite incluir caracteres extranjeros más que el alfa-numérico " + vbCrLf
        s$ = s$ + "palabras, incluye los símbolos acentuados " + vbCrLf
        s$ = s$ + "como un EX I, etc " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[Añadir] " + vbCrLf
        s$ = s$ + "> Añadir a la a la izquierda del elemento> antepone el prefijo a cada elemento de la lista. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> de usuario personalizada lista selecciona el archivo, cada elemento de la " + vbCrLf
        s$ = s$ + "lista se agrega al principio de " + vbCrLf
        s$ = s$ + "cada elemento de la lista. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> def. prefijos programa utiliza por defecto, elegido por el autor " + vbCrLf
        s$ = s$ + "prefijos, añade a cada elemento en " + vbCrLf
        s$ = s$ + "la lista. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> numérico añade 0-9 al comienzo de cada elemento en " + vbCrLf
        s$ = s$ + "la lista. Sólo tiene un número de " + vbCrLf
        s$ = s$ + "en-un-tiempo. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> añade un alfa-Z a principios de cada elemento en " + vbCrLf
        s$ = s$ + "la lista. Sólo hace una carta " + vbCrLf
        s$ = s$ + "en-un-tiempo. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Añadir a la derecha del punto> añade postfix a cada elemento de la " + vbCrLf
        s$ = s$ + "lista. (igual al anterior, pero añade que " + vbCrLf
        s$ = s$ + "Al final de cada tema). " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "el uso de " + vbCrLf
        s$ = s$ + "-------- " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "las opciones en L517 están diseñados para proporcionar las herramientas " + vbCrLf
        s$ = s$ + " necesarios para construir la 'lista de palabras perfecto'. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "esta es una re-escritura y la expansión a un programa anterior, " + vbCrLf
        s$ = s$ + " 'LST', que era lento, engorroso, y carecía de las opciones y " + vbCrLf
        s$ = s$ + " características que yo quería usar. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "L517 es pequeño, ligero, y hace todo lo posible a la generación de " + vbCrLf
        s$ = s$ + "  listas de palabras. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "algunas listas tardará más tiempo en la carga que otros. algunas de las opciones " + vbCrLf
        s$ = s$ + " (palabras-de-página web) puede bloquear el programa hasta por minutos a la vez. " + vbCrLf
        s$ = s$ + " dupekilling listas grandes pueden tardar horas. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "mi punto es: este programa es estable. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "Buena suerte, y este programa puede generar un largo buscados " + vbCrLf
        s$ = s$ + " contraseña. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "-derv" + vbCrLf
    Case "german"
        s$ = "L517: Wortliste Generator: Readme " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "Shortcuts " + vbCrLf
        s$ = s$ + "---------- " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "1. Drag-and-Drop-Datei (en), die beliebig zu diesem Antrag in den " + vbCrLf
        s$ = s$ + "     Dateien schnell. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "2. Sie können auch per Drag-and-Drop den markierten Text (von Dateien, " + vbCrLf
        s$ = s$ + "     E-Mails, Webseiten) in der Liste. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "3. können Sie Text einfügen in die Liste! drücken Sie Strg + V überall " + vbCrLf
        s$ = s$ + "     in der Anwendung. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "4. Doppelklick auf Titelleiste zu 'schrumpfen' das Programm. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "5. Die Speisekarte bietet viele Verknüpfungen, die Arbeit überall in " + vbCrLf
        s$ = s$ + "     das Programm: " + vbCrLf
        s$ = s$ + "Strg O: Datei öffnen " + vbCrLf
        s$ = s$ + "Strg + V: load Worte aus der Zwischenablage " + vbCrLf
        s$ = s$ + "Strg D: dupe-Kill-Liste " + vbCrLf
        s$ = s$ + "ctrl F: Suche Element in der Liste " + vbCrLf
        s$ = s$ + "Löschen: Entfernen Sie ausgewählte Element " + vbCrLf
        s$ = s$ + "Shift-Entf: Liste löschen " + vbCrLf
        s$ = s$ + "... und noch einige mehr. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "Menü " + vbCrLf
        s$ = s$ + "---------- " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[FILE] " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> neue Liste löscht. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> open öffnet Wortliste, fügt jedes gefundene Wort " + vbCrLf
        s$ = s$ + "auf der Liste. (Griffe-fast-alles-Datei " + vbCrLf
        s$ = s$ + "Arten / Formate) " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Speichern unter ... Wortliste speichert die Datei. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> split in Datei speichert Liste, um eine neue Datei alle X-Artikel. " + vbCrLf
        s$ = s$ + "verwendet sequentielle Datei-naming '1. txt, 2.txt' " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> save type> " + vbCrLf
        s$ = s$ + "> Windows fügt Wagen zurück zum Ende der Zeilen und speichert die " + vbCrLf
        s$ = s$ + "Dateien, die mit Windows-Anwendungen " + vbCrLf
        s$ = s$ + "> Unix verwendet nur Line Feed am Ende der Linien, Werke " + vbCrLf
        s$ = s$ + "mit Unix [und einige Fenster] Anwendungen " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Voreinstellungen> " + vbCrLf
        s$ = s$ + "> WPA setzt die ordnungsgemäße minimale / maximale Länge " + vbCrLf
        s$ = s$ + "für die WPA-Passwörter (8-64) " + vbCrLf
        s$ = s$ + "> web-Sets ordnungsgemäße min / max Länge für web-basierte " + vbCrLf
        s$ = s$ + "Passwörter (4-12) vbCrLf " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Exit beendet wird. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[EDIT] " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Entfernen von Duplikaten entfernt doppelte Elemente. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Liste löschen löscht Liste (nach Rückfrage). " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> remove Artikel ausgewählt Punkt entfernt, wie " + vbCrLf
        s$ = s$ + "Delete-Taste drücken. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Entfernen von Elementen mit entfernt Elemente, die enthalten ein " + vbCrLf
        s$ = s$ + "Text-String. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Entfernen von Elementen, ohne entfernt Elemente, die NICHT enthalten " + vbCrLf
        s$ = s$ + "Text-String. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> find sucht nach Wort [oder einen Teil des Wortes] " + vbCrLf
        s$ = s$ + "in der Liste. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Weitersuchen Sucht das nächste Vorkommen von zuvor " + vbCrLf
        s$ = s$ + "-gesuchte Wort. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[GENERATE] " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Worte von der Website Auszüge Worte von der Website, " + vbCrLf
        s$ = s$ + "Herausfiltern von HTML. (BUG: LAG) " + vbCrLf
        s$ = s$ + "für eine einfachere Methode, besuchen Sie die Website " + vbCrLf
        s$ = s$ + "in IE / Firefox, wählen Sie die gesamte " + vbCrLf
        s$ = s$ + "Webseite (Strg A), " + vbCrLf
        s$ = s$ + "klicken Sie dann auf und ziehen Sie den Text in " + vbCrLf
        s$ = s$ + "Das Programm-Liste. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Worte aus dem Ordner (n) Scans Verzeichnis (und Unterverzeichnisse). " + vbCrLf
        s$ = s$ + "Auszüge Wörter aus jpg, mp3, doc, " + vbCrLf
        s$ = s$ + "srt, und andere. " + vbCrLf
        s$ = s$ + "umfasst nicht: RAR, AVI, MPG, MP4, ISO, " + vbCrLf
        s$ = s$ + "zip, msi, exe,. torrent, dat, tar " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> string aus charset String generiert die gewünschte Länge " + vbCrLf
        s$ = s$ + "auf der Grundlage 'charset.lst', " + vbCrLf
        s$ = s$ + "ein character-set-Datei. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Termine> " + vbCrLf
        s$ = s$ + "> separator entscheidet, was die Monate, Tage und " + vbCrLf
        s$ = s$ + "Jahre sind getrennt. null sein. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> [mm / tt / jj] Benutzer legt die Start-und Stopp Jahr " + vbCrLf
        s$ = s$ + "und das Programm erzeugt, basierend auf " + vbCrLf
        s$ = s$ + "das Format, die alle Daten innerhalb " + vbCrLf
        s$ = s$ + "dieser Zeitspanne besteht. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Telefonnummern> " + vbCrLf
        s$ = s$ + "Trennzeichen, welchen Charakter / string trennt die " + vbCrLf
        s$ = s$ + "Ortsvorwahl, Vorwahl als auch 4-stellige Zahl " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> [Ortsvorwahl] [prefix ]#### generiert jede Telefonnummer " + vbCrLf
        s$ = s$ + "innerhalb einer bestimmten Stadt / Land. " + vbCrLf
        s$ = s$ + "speichert die Datei-Liste zu groß für " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> [prefix ]#### gleiche wie oben " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "Optionen " + vbCrLf
        s$ = s$ + "--------- " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[CASE] " + vbCrLf
        s$ = s$ + "Convert to: " + vbCrLf
        s$ = s$ + "  > Kleinbuchstaben jedes Element in der Liste wird geändert, um LCase " + vbCrLf
        s$ = s$ + "  > UPPERCASE jedem Punkt auf der oberen umgewandelt Fall " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "Kopien mit: " + vbCrLf
        s$ = s$ + "  > First Letter Upper spart ursprüngliche Liste (siehe Abbildung), plus " + vbCrLf
        s$ = s$ + "Liste umgewandelt wird 1stLetterUpper " + vbCrLf
        s$ = s$ + "  > Jedes andere Upper ursprünglichen Format speichert jede andere " + vbCrLf
        s$ = s$ + "Groß-Format " + vbCrLf
        s$ = s$ + "  > 1337C453 speichert die Liste, und spart auch die " + vbCrLf
        s$ = s$ + "Liste umgewandelt 'leetspeak' " + vbCrLf
        s$ = s$ + "leet Wörterbuch ist in 'leetspeak.txt', " + vbCrLf
        s$ = s$ + "und editierbar! " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[FILTER] " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> set min / max Länge entfernt (und wird nicht bewertet) Worte " + vbCrLf
        s$ = s$ + "kleiner / größer als die min / max. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> Text nach rechts / links ... spaltet Wort aus jeder Zeile " + vbCrLf
        s$ = s$ + "in einer Datei. " + vbCrLf
        s$ = s$ + "kann für eine Vielzahl von Gründen sinnvoll. " + vbCrLf
        s$ = s$ + "d. h.: " + vbCrLf
        s$ = s$ + " soll eine Wortliste ist nummeriert: " + vbCrLf
        s$ = s$ + "1. word1 " + vbCrLf
        s$ = s$ + "2. word2 " + vbCrLf
        s$ = s$ + "3. word3 " + vbCrLf
        s$ = s$ + "nur die Worte Extrakt (no # 's), " + vbCrLf
        s$ = s$ + "Sie wählen Sie 'Text rechts' " + vbCrLf
        s$ = s$ + "und geben Sie '. '  Ohne Anführungszeichen. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> convert !@#... zu hex dies spezielle (nicht alpha ändern " + vbCrLf
        s$ = s$ + "numerisch)-Zeichen in ihre hex " + vbCrLf
        s$ = s$ + "-Äquivalent. d. h.! @ # " + vbCrLf
        s$ = s$ + "sinnvoll, wenn Sie eine Wortliste, dass " + vbCrLf
        s$ = s$ + "wird über das Internet verwendet werden. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> include ausländischen Zeichen können mehr als nur alpha-numerischen " + vbCrLf
        s$ = s$ + "Wörtern und enthält Symbole mit Akzent " + vbCrLf
        s$ = s$ + "wie eine e ï etc " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "[APPEND] " + vbCrLf
        s$ = s$ + "> in den links item> Präfix vorangestellt zu jedem Element in der Liste. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> benutzerdefinierte Liste Benutzer wählt Datei, jedes Element aus dem " + vbCrLf
        s$ = s$ + "Liste ist bis zum Beginn des zugesetzten " + vbCrLf
        s$ = s$ + "jedes Element in der Liste. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> def. Präfixe Programm verwendet standardmäßig ausgewählt-by-Autor " + vbCrLf
        s$ = s$ + "Präfixe, fügt sie jedes Element in " + vbCrLf
        s$ = s$ + "der Liste. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> numerische fügt 0-9 zu Beginn jedes Element in " + vbCrLf
        s$ = s$ + "der Liste. Erst eine Zahl " + vbCrLf
        s$ = s$ + "at-a-time. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> alpha fügt a-z zu Beginn jedes Element in " + vbCrLf
        s$ = s$ + "der Liste. Erst ein Schreiben " + vbCrLf
        s$ = s$ + "at-a-time. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "> nach rechts von Element hinzufügen> hängt postfix zu jedem Element in der " + vbCrLf
        s$ = s$ + "Liste. (gleiche wie oben, jedoch ergänzt " + vbCrLf
        s$ = s$ + "Ende jedes Artikels). " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "Nutzung " + vbCrLf
        s$ = s$ + "-------- " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "die Optionen in L517 sind entworfen, um die Werkzeuge liefern, " + vbCrLf
        s$ = s$ + " erforderlich sind, um beim Aufbau der 'perfekten Wortliste.' " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "dies ist ein neu zu schreiben und die Expansion auf einem früheren Programm " + vbCrLf
        s$ = s$ + " 'LST', die langsam war, umständlich und nicht über die Optionen und " + vbCrLf
        s$ = s$ + " Funktionen, die ich nutzen wollte. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "L517 ist klein, leicht und tut sein Bestes, um rechtzeitig die " + vbCrLf
        s$ = s$ + "  Wortlisten. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "einigen Listen wird länger dauern als andere zu laden. einige Optionen " + vbCrLf
        s$ = s$ + " (Wörter-from-Website) kann das Programm Sperre für Minuten auf einmal. " + vbCrLf
        s$ = s$ + " dupekilling umfangreichen Listen kann Stunden dauern. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "Mein Punkt ist: dieses Programm ist stabil. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "viel Glück und kann dieses Programm erzeugen eine lange gesuchten nach " + vbCrLf
        s$ = s$ + " Passwort für Sie. " + vbCrLf
        s$ = s$ + "" + vbCrLf
        s$ = s$ + "-derv" + vbCrLf
    End Select
    
    readme$ = App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + "readme.txt"
    Do While Len(Dir(readme$)) <> 0
        DoEvents
        Kill readme$
    Loop
    
    ff% = FreeFile
    Open readme$ For Binary Access Write As #ff%
        Put #ff%, , s$
    Close #ff%
    
    DoEvents
    Shell "explorer " + Chr(34) + readme$ + Chr(34), vbNormalFocus
    DoEvents
    
    'timah# = Timer
    'Do While Timer - timah# < 2
    '    DoEvents
    'Loop
    
    'timah# = Timer
    'Do While Timer - timah# < 5
    '    DoEvents
    '    If Len(Dir(readme$)) <> 0 Then
    '        Kill readme$
    '        Exit Do
    '    End If
    'Loop
End Sub

Private Sub mnuHelpItems_Click()
    Dim s$
    
    Select Case LANGUAGE$
    Case "english"
        s$ = "are certain items not being added to the list during the load?" + vbCrLf + vbCrLf + _
             "make sure your filters are turned off (or set to an desired level) before loading items." + vbCrLf + vbCrLf + _
             "sometimes items that would normally be added to the list are ignored because of the min/max length filter."
    Case "french"
        s$ = "sont certains articles ne sont pas ajoutés à la liste au cours de la charge?" + vbCrLf + vbCrLf + _
             "Assurez-vous que vos filtres sont hors tension (ou un ensemble à un niveau désiré) avant le chargement articles." + vbCrLf + vbCrLf + _
             "parfois des éléments qui seraient normalement ajoutés à la liste sont ignorées en raison de la min / filtre de longueur max."
    Case "german"
        s$ = "sind bestimmte Posten nicht in die Liste aufgenommen während der Last?" + vbCrLf + vbCrLf + _
             "Vergewissern Sie sich, Ihre Filter deaktiviert sind (oder zu einem gewünschten Niveau festgesetzt) vor dem Laden von Elementen." + vbCrLf + vbCrLf + _
             "manchmal Elemente, die normalerweise in die Liste aufgenommen werden, weil der ignoriert min / max Länge Filter."
    Case "spanish"
        s$ = "ciertos artículos no se añade a la lista durante la carga?" + vbCrLf + vbCrLf + _
             "Asegúrese de que sus filtros están apagados (o conjunto a un nivel deseado) antes de cargar los elementos." + vbCrLf + vbCrLf + _
             "A veces los artículos que normalmente se añade a la lista son ignorados por el min / filtro de longitud máxima."
    End Select
    frmMsg.MsgBocks s$, vbInformation + vbOKOnly
End Sub

Private Sub mnuHelpSite_Click()
    Shell "explorer " + Chr(34) + "http://code.google.com/p/l517/" + Chr(34), vbNormalFocus
End Sub

Private Sub mnuHelpSpanish_Click()
    LANGUAGE$ = "spanish"
    regSet "lang", LANGUAGE$
    Change_Language
End Sub

Private Sub mnuListClear_Click()
    Dim s$
    If lst.ListItems().count = 0 Then Exit Sub
    
    Select Case LANGUAGE$
    Case "english"
        s$ = "are you sure you want to clear this list?"
    Case "french"
        s$ = "Etes-vous sûr de vouloir effacer cette liste?"
    Case "german"
        s$ = "Sie sind sicher, dass Sie diese Liste zu löschen?"
    Case "spanish"
        s$ = "¿Está seguro que quiere borrar esta lista?"
    End Select
    If frmMsg.MsgBocks(s$, vbQuestion + vbYesNo, "L517") = vbYes Then
        
        stat "clearing list..."
        DoEvents
        lst.ListItems().Clear
        
        stat "list cleared"
        UpdateCaption
    End If
End Sub

Private Sub arrAdd(arr$(), add$)
    Dim ln&
    ln& = UBound(arr$())
    ln& = ln& + 1
    ReDim Preserve arr$(ln&)
    arr$(ln&) = add$
End Sub
Private Sub arrClear(arr$())
    ReDim arr$(0)
End Sub
Private Function arrFind(sarr$(), item$) As Boolean
    Dim i&
    
    For i& = 0 To UBound(sarr$())
        If sarr$(i&) = item$ Then
            arrFind = True
            Exit Function
        End If
    Next i&
    
    arrFind = False
End Function

Private Sub dupekill()
    Dim i&, count&, ttl&, tempfile$, ff%, lastitem$, sarr$(), item$, removed$, donotadd As Boolean
    
    ReDim sarr$(0)
    
    ttl& = lst.ListItems().count
    count& = 0
    
    lst.Visible = False
    
    stat "filtering..."
    tempfile$ = App.Path + IIf(Right(App.Path, 1) = "\", "", "\") + "_temp.txt"
    ff% = FreeFile
    Open tempfile$ For Binary Access Write As #ff%
        For i& = 1 To ttl&
            If i& Mod 2500 = 0 Then
                prog i& / (ttl& * 2)
                DoEvents
            End If
            
            donotadd = False
            
            item$ = lst.ListItems(i&).Text
            
            If UBound(sarr$()) > 0 Then
                If LCase(item$) = LCase(sarr$(1)) Then
                    If arrFind(sarr$(), item$) = True Then
                        donotadd = True
                    End If
                Else
                    arrClear sarr$()
                End If
            End If
            
            If donotadd = False Then
                count& = count& + 1
                Put #ff%, , CStr(item$ + vbCrLf)
            End If
            
            arrAdd sarr$(), item$
        Next i&
    Close #ff%
    
    stat "clearing..."
    lst.ListItems().Clear
    
    stat "loading..."
    i& = 0
    ff% = FreeFile
    Open tempfile$ For Input As #ff%
        While Not EOF(ff%)
            i& = i& + 1
            If i& Mod 2500 = 0 Then
                prog (count& + i&) / (count& * 2)
                DoEvents
            End If
            Input #ff%, item$
            lst.ListItems().add , , item$
        Wend
    Close #ff%
    
    stat "displaying..."
    lst.Visible = True
    
    On Error Resume Next
    Kill tempfile$
    On Error GoTo 0
    
    removed$ = Format(ttl& - count&, "###,###")
    If removed$ = "" Then removed$ = "0"
    stat "inactive; " + removed$ + " removed"
    
    prog 0
    UpdateCaption
End Sub

Private Sub mnuListDupekill_Click()
    Dim i&, count&
    
    dupekill
    Exit Sub
'
'    i& = 1
'    count& = 0
'
'    lst.Visible = False
'
'    stat "dupekilling"
'
'    prog 0.001
'
'    Do While i& < lst.ListItems().count
'        If i& Mod 250 = 0 Or (count& + i&) Mod 20 = 0 Then
'            DoEvents
'
'            If lblCancel.Visible = False Then
'                stat "canceled, " + IIf(count& = 0, "0", Format(count&, "###,###")) + " removed"
'                prog 0
'                lst.Visible = True
'                If count& > 0 Then B_CHANGE = True
'                UpdateCaption
'                Exit Sub
'            End If
'
'            UpdateCaption
'            stat "dupekilling; " + IIf(count& = 0, "0", Format(count&, "###,###")) + " removed" '(" + Format(i&, "###,###") + "/" + Format(lst.ListItems().count, "###,###") + ")"
'            prog i& / lst.ListItems().count
'        End If
'
'        If lst.ListItems(i&).Text = lst.ListItems(i& + 1).Text Then
'            lst.ListItems().Remove i&
'            count& = count& + 1
'            i& = i& - 1
'            If i& < 1 Then i& = 1
'        Else
'            If LCase(lst.ListItems(i&).Text) = LCase(lst.ListItems(i& + 1).Text) And i& < lst.ListItems().count - 1 Then
'                If lst.ListItems(i&).Text = lst.ListItems(i& + 2).Text Then 'Or lst.ListItems(i&).Text = lst.ListItems(i& + 3).Text Then
'                    lst.ListItems().Remove i&
'                    'i& = i& - 1
'                    'If i& < 1 Then i& = 1
'                    count& = count& + 1
'                Else
'                    i& = i& + 1
'                End If
'            Else
'                i& = i& + 1
'            End If
'        End If
'
'    Loop
'
'    If count& > 0 Then B_CHANGE = True
'
'    prog 0
'    lst.Visible = True
'
'    UpdateCaption
'
'    stat "inactive; " + IIf(count& = 0, "0", Format(count&, "###,###")) + " removed"
End Sub

Private Sub mnuListFind_Click()
    Dim s$, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "enter the string (or part of the string) you want to find:"
    Case "french"
        sl$ = "entrez la chaîne (ou une partie de la chaîne) que vous souhaitez trouver:"
    Case "german"
        sl$ = "geben Sie die Zeichenfolge (oder ein Teil des Strings) Sie suchen möchten:"
    Case "spanish"
        sl$ = "entrar en la cadena (o parte de la cadena) que desea buscar:"
    End Select
    s$ = frmMsg.InputBocks(sl$, "L517", mnuListFind.Tag)
    If s$ = "" Then Exit Sub
    
    mnuListFind.Tag = s$
    mnuListFindNext_Click
End Sub

Private Sub mnuListFindNext_Click()
    Dim s$, start&, i&
    
    s$ = LCase(mnuListFind.Tag)
    If s$ = "" Then Exit Sub
    
    If mnuListFindNext.Tag = "" Then
        start& = 1
    Else
        start& = CLng(mnuListFindNext.Tag) + 1
    End If
    
    stat "searching..."
    For i& = start& To lst.ListItems().count
        If InStr(LCase(lst.ListItems(i&).Text), s$) <> 0 Then
            lst.ListItems(i&).Selected = True
            lst.ListItems(i&).EnsureVisible
            mnuListFindNext.Tag = CStr(i&)
            stat "found '" + s$ + "'!"
            Exit Sub
        End If
    Next i&
    
    stat "'" + s$ + "' not found"
    mnuListFindNext.Tag = "1"
End Sub

Private Sub mnuListPaste_Click()
    Form_KeyDown Asc("V"), 2
End Sub

Private Sub mnuListRemove_Click()
    On Error GoTo Err_Occ
    
    'If lst.SelectedItem.Selected = True Then
        lst.ListItems().Remove lst.SelectedItem.Index
        UpdateCaption
    'End If
Err_Occ:
End Sub

Private Sub mnuListRemoveItemsWith_Click()
    Dim s$, i&, count&, sl$
    
    Select Case LANGUAGE$
    Case "english"
        sl$ = "remove all items containing the text below:"
    Case "french"
        sl$ = "supprimer tous les éléments contenant le texte ci-dessous:"
    Case "german"
        sl$ = "alle Einträge mit dem Text:"
    Case "spanish"
        sl$ = "eliminar todos los elementos que contiene el siguiente texto:"
    End Select
    s$ = LCase(frmMsg.InputBocks(sl$))
    If s$ = "" Then Exit Sub
    
    i& = 1
    count& = 0
    
    lst.Visible = False
    
    prog 0.001
    
    stat "filtering list"
    
    Do While i& < lst.ListItems().count
        If i& Mod 100 = 0 Then
            DoEvents
            prog i& / lst.ListItems().count
            If lblCancel.Visible = False Then
                lst.Visible = True
                UpdateCaption
                Exit Sub
            End If
            UpdateCaption
            
            stat "removing, " + CStr(count&) + " removed"
        End If
        
        If InStr(LCase(lst.ListItems(i&).Text), s$) <> 0 Then
            lst.ListItems().Remove i&
        Else
            i& = i& + 1
        End If
    Loop
    
    prog 0
    stat "removed " + CStr(count&)
    UpdateCaption
    lst.Visible = True
    
End Sub
