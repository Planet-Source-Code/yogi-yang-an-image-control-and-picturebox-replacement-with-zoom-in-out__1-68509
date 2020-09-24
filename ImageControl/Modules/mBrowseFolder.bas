Attribute VB_Name = "mBrowseFolder"
'//  ====================================================
'//  Mod. name: FileFound
'//  Author:    Steve Anderson (Transf. to function)
'//  ----------------------------------------------------
'//  Function BrowseFolder(Tittle As String, OwnerForm As
'//                        Form) As String
'//  ====================================================

Option Explicit

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260

Public Declare Function SHBrowseForFolder Lib "shell32" _
                        (lpbi As BrowseInfo) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32" _
                        (ByVal pidList As Long, _
                        ByVal lpBuffer As String) As Long

Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                        (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type



Public Function BrowseFolder(ByVal Tittle As String, ByRef OwnerForm As Form) As String

  Dim lpIDList As Long
  Dim sBuffer As String
  Dim tBrowseInfo As BrowseInfo

    With tBrowseInfo
        .hWndOwner = OwnerForm.hwnd
        .lpszTitle = lstrcat(Tittle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseFolder = sBuffer
      Else
        BrowseFolder = ""
    End If

End Function

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 11:39:21) 34 + 27 = 61 Lines
