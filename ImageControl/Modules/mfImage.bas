Attribute VB_Name = "mfImage"
'//
'// fImage control module
'//

Option Explicit

'Public fIm As fImage                '// Current image 'pointer'
Public fImPstBits() As Byte         '// Current paste bits
Public fImPasting As Boolean        '// Pasting flag
Public fImToolApplied As Boolean    '// Image changed by tool applying
Public fImToolChanged As Boolean    '// Tool has changed
Public fIDLast As Long              '// Last ID

Public txtBits() As Byte            '// Text bits
Public stmBits() As Byte            '// Stamp bits

Private fID As Long                 '// Image ID 'Copies/New sub-counter'
Private fIDAbs As Long              '// Absolute counter





'Public Function GetActiveImage(ByRef frm As fImage) As Boolean
'
'    If Not (fMain.ActiveForm Is Nothing) Then
'        If (fMain.ActiveForm.name = "fImage") Then
'            Set frm = fMain.ActiveForm
'            GetActiveImage = -1
'        End If
'    End If
'
'End Function

Public Function GetImageTag(ByVal frmTag As String) As String

  Dim i As Long
  Dim iExists As Boolean

    For i = 0 To Forms.Count - 1
        If (Forms(i).name = "fImage") Then
            If (frmTag = Forms(i).Tag) Then
                iExists = -1
            End If
        End If
    Next i

    If (iExists) Then
        fID = fID + 1
        GetImageTag = frmTag & "[" & fID & "]"
      Else
        GetImageTag = frmTag
    End If
    
    fIDAbs = fIDAbs + 1

End Function

Public Function GetImageName(ByVal frmTag As String) As String

    GetImageName = RTrim(Mid(frmTag, InStrRev(frmTag, "\") + 1))

End Function

Public Function GetImageID() As Long

    GetImageID = fIDAbs

End Function

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 11:42:35) 14 + 91 = 105 Lines
