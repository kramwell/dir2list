VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dir2list - KramWell.com"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check12 
      Caption         =   "Pauses after each screenful of information"
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4935
      Begin VB.CommandButton Command3 
         Caption         =   "Save to Text File"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   4320
         Width           =   4695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Timefield"
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   3600
         Width           =   4695
         Begin VB.CheckBox Check17 
            Caption         =   "Last Written"
            Height          =   255
            Left            =   2520
            TabIndex        =   23
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox Check16 
            Caption         =   "Last Access"
            Height          =   195
            Left            =   1200
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox Check15 
            Caption         =   "Creation"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Attributes"
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4695
         Begin VB.CheckBox Check6 
            Caption         =   "Files are list sorted by column"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   2415
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Bare format (no heading or summary)"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   3015
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Files ready for archiving"
            Height          =   255
            Left            =   2280
            TabIndex        =   9
            Top             =   480
            Width           =   2055
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Read-only files"
            Height          =   255
            Left            =   2280
            TabIndex        =   8
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Hidden files"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Directories"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Uses lowercase"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   4320
            Y1              =   840
            Y2              =   840
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sortorder"
         Height          =   1575
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   4695
         Begin VB.CheckBox Check14 
            Caption         =   "Displays files in specified directory and all subdirectories"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1200
            Width           =   4215
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Display the owner of the file"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   3375
         End
         Begin VB.CheckBox Check11 
            Caption         =   "By date/time (oldest first)"
            Height          =   255
            Left            =   2280
            TabIndex        =   17
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox Check10 
            Caption         =   "By extension (alphabetic)"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox Check9 
            Caption         =   "By size (smallest first)"
            Height          =   255
            Left            =   2280
            TabIndex        =   15
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox Check8 
            Caption         =   "By name (alphabetic)"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2175
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   4320
            Y1              =   840
            Y2              =   840
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save to..."
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "select where the file should be saved!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "click browse and get the folder location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by KramWell.com - 11/SEPT/2007
'This script will output a text file for all files in a selected folder (with advanced options)

'declarations for the shell command
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Declare Function FindExecutable Lib "shell32.dll" Alias _
         "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
         String, ByVal lpResult As String) As Long
         
Dim strDir As String
Dim strRead As String
Dim strHidden As String
Dim strReady As String

Dim strBare As String
Dim strSort As String
Dim strLower As String

Dim strName As String
Dim strSize As String
Dim strExten As String
Dim strDate As String

Dim strPause As String
Dim strOwner As String
Dim strSubdir As String

Dim strCreation As String
Dim strLast As String
Dim strWritten As String

 Dim txtdisplayname As String



'shell stop
         


Private Sub Check1_Click()

If Check1.Value = 1 Then
strDir = ":D"
    Else
strDir = ""
End If


End Sub

Private Sub Check10_Click()
If Check10.Value = 1 Then
strExten = ":E"
    Else
strExten = ""
End If
End Sub

Private Sub Check11_Click()
If Check11.Value = 11 Then
strDate = ":D"
    Else
strDate = ""
End If
End Sub

'Private Sub Check12_Click()
'If Check12.Value = 1 Then
'strPause = " /P"
'    Else
'strPause = ""
'End If
'End Sub

Private Sub Check13_Click()
If Check13.Value = 1 Then
strOwner = " /Q"
    Else
strOwner = ""
End If
End Sub

Private Sub Check14_Click()
If Check14.Value = 1 Then
strSubdir = " /S"
    Else
strSubdir = ""
End If
End Sub

Private Sub Check15_Click()
If Check15.Value = 1 Then
strCreation = ":C"
    Else
strCreation = ""
End If
End Sub

Private Sub Check16_Click()
If Check16.Value = 1 Then
strLast = ":A"
    Else
strLast = ""
End If
End Sub

Private Sub Check17_Click()
If Check17.Value = 1 Then
strWritten = ":W"
    Else
strWritten = ""
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
strHidden = ":H"
    Else
strHidden = ""
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
strReady = ":R"
    Else
strReady = ""
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
strRead = ":A"
    Else
strRead = ""
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
strBare = " /B"
    Else
strBare = ""
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
strSort = " /D"
    Else
strSort = ""
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then
strLower = " /L"
    Else
strLower = ""
End If
End Sub

Private Sub Check8_Click()
If Check8.Value = 1 Then
strName = ":N"
    Else
strName = ""
End If
End Sub

Private Sub Check9_Click()
If Check9.Value = 1 Then
strSize = ":S"
    Else
strSize = ""
End If
End Sub

Private Sub Command1_Click()
  Dim BI As BROWSEINFO
  Dim nFolder As Long
  Dim IDL As ITEMIDLIST
  Dim pIdl As Long
  Dim sPath As String
  Dim SHFI As SHFILEINFO
 
 
  With BI
    '// The dialog'//s owner window...
    .hOwner = Me.hwnd
   
    '// Initialize the buffer that rtns the display name of the selected folder
    .pszDisplayName = String$(MAX_PATH, 0)
   
    '// Set the dialog'//s banner text
    .lpszTitle = "Browse for Folder"
   
    '// Set the type of folders to display & return
    '// -play with these option constants to see what can be returned
    .ulFlags = GetReturnType()
   
  End With
 
  '// Show the Browse dialog
  pIdl = SHBrowseForFolder(BI)
 
  '// If the dialog was cancelled...
  If pIdl = 0 Then Exit Sub
   
  '// Fill sPath w/ the selected path from the id list
  '// (will rtn False if the id list can'//t be converted)
  sPath = String$(MAX_PATH, 0)
  SHGetPathFromIDList ByVal pIdl, ByVal sPath

  '// Display the path and the name of the selected folder
  Label1.Caption = Left(sPath, InStr(sPath, vbNullChar) - 1)
  txtdisplayname = Left$(BI.pszDisplayName, _
                             InStr(BI.pszDisplayName, vbNullChar) - 1)
 
  '// Frees the memory SHBrowseForFolder()
  '// allocated for the pointer to the item id list
  CoTaskMemFree pIdl
End Sub

Private Sub Command2_Click()
  Dim BI As BROWSEINFO
  Dim nFolder As Long
  Dim IDL As ITEMIDLIST
  Dim pIdl As Long
  Dim sPath As String
  Dim SHFI As SHFILEINFO
 
  With BI
    '// The dialog'//s owner window...
    .hOwner = Me.hwnd
   
    '// Initialize the buffer that rtns the display name of the selected folder
    .pszDisplayName = String$(MAX_PATH, 0)
   
    '// Set the dialog'//s banner text
    .lpszTitle = "Browse for Folder"
   
    '// Set the type of folders to display & return
    '// -play with these option constants to see what can be returned
    .ulFlags = GetReturnType()
   
  End With
 
  '// Show the Browse dialog
  pIdl = SHBrowseForFolder(BI)
 
  '// If the dialog was cancelled...
  If pIdl = 0 Then Exit Sub
   
  '// Fill sPath w/ the selected path from the id list
  '// (will rtn False if the id list can'//t be converted)
  sPath = String$(MAX_PATH, 0)
  SHGetPathFromIDList ByVal pIdl, ByVal sPath

  '// Display the path and the name of the selected folder
  Label2.Caption = Left(sPath, InStr(sPath, vbNullChar) - 1)
  'txtDisplayName = Left$(BI.pszDisplayName, _
                             InStr(BI.pszDisplayName, vbNullChar) - 1)
 
  '// Frees the memory SHBrowseForFolder()
  '// allocated for the pointer to the item id list
  CoTaskMemFree pIdl
End Sub

Private Sub Command3_Click()

Dim Taskid As Long
Dim strLong As String
Dim strBrowsepath As String
Dim strBrowsepath2 As String
Dim strTotal As String

Dim strAttrib As String
Dim strSortorder As String
Dim strTime As String


If Label1.Caption = "click browse and get the folder location" Or Label2.Caption = "select where the file should be saved!" Then
MsgBox "Please choose the file location and/or where it has to be saved to...", vbOKOnly, "dir2list"

Else

    strBrowsepath = Label1.Caption
    strBrowsepath2 = Label2.Caption
    
    strAttrib = " /A" + strDir + strRead + strHidden + strReady
    strSortorder = " /O" + strName + strSize + strExten + strDate
    strTime = " /T" + strCreation + strLast + strWritten
    
    strLong = strAttrib + strBare + strSort + strLower + strSortorder + strPause + strOwner + strSubdir + strTime
    
    
    strTotal = "cmd.exe /c dir " + """" + strBrowsepath + """" + strLong + " > " + """" + strBrowsepath2 + "\dir2list(" + txtdisplayname + ").txt" + """"
    
    'MsgBox strTotal, vbOKOnly, "Title"
    
    Taskid = Shell(strTotal, vbNormalFocus)

End If

End Sub

'// get the options
Private Function GetReturnType() As Long
  Dim dwRtn As Long
  GetReturnType = dwRtn
End Function

Private Sub Label1_Click()
If Label1.Caption <> "" Then
    MsgBox Label1.Caption, vbOKOnly, "dir2list"
        End If
End Sub

Private Sub Label2_Click()
If Label2.Caption <> "" Then
    MsgBox Label2.Caption, vbOKOnly, "dir2list"
        End If
End Sub
