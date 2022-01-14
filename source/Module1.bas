Attribute VB_Name = "Module1"
' Maximun long filename path length
Public Const MAX_PATH = 260


Type SHITEMID   ' mkid
    cb As Long       ' Size of the ID (including cb itself)
    abID() As Byte  ' The item ID (variable length)
End Type

Type SHFILEINFO   ' shfi
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Type ITEMIDLIST   ' idl
    mkid As SHITEMID
End Type
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                              (ByVal pIdl As Long, ByVal pszPath As String) As Long

' Frees memory allocated by SHBrowseForFolder()
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                              (lpBrowseInfo As BROWSEINFO) As Long ' ITEMIDLIST

Public Type BROWSEINFO   ' bi

    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long

    iImage As Long

End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4



