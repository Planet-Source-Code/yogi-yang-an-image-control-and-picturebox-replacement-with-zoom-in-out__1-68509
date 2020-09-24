Attribute VB_Name = "mFileFound"
'//  ====================================================
'//  Mod. name: FileFound
'//  Author:    Eric Russell
'//  Created:   1998-03-19
'//  Revised:   1998-03-20
'//  ----------------------------------------------------
'//  Function FileFound(strFileName As String) As Boolean
'//
'//  Parameters: (string) Filename or folder.
'//  Returns: (Boolean) True if the file/folder was found
'//           or False if not.
'//  ====================================================

Option Explicit

Public Const MAX_PATH = 260

Type FILETIME '8 Bytes
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA '318 Bytes
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved_ As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Declare Function FindFirstFile& Lib "kernel32" Alias "FindFirstFileA" _
                        (ByVal lpFileName As String, _
                        lpFindFileData As WIN32_FIND_DATA)

Public Declare Function FindClose Lib "kernel32" _
                        (ByVal hFindFile As Long) As Long
                        
                        

Public Function FileFound(strFileName As String) As Boolean

  Dim lpFindFileData As WIN32_FIND_DATA
  Dim hFindFirst As Long

    hFindFirst = FindFirstFile(strFileName, lpFindFileData)

    If (hFindFirst > 0) Then
        FindClose hFindFirst
        FileFound = -1
      Else
        FileFound = 0
    End If

End Function

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 11:41:31) 41 + 18 = 59 Lines
