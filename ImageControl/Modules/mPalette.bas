Attribute VB_Name = "mPalette"
'//
'// Palette module (Used to Apply palette)
'// Special thanks vbaccelerator.com
'//

Option Explicit

Public Type RGBPalette
    Rp As Long
    Gp As Long
    Bp As Long
End Type

Public Type tPalette
    nColors As Long
    Color() As RGBPalette
End Type

Public Pal As tPalette

Private Type cCol
    lRGB As Long
    Count As Long
End Type



Public Sub ClosestColor(ByVal Rf As Long, ByVal Gf As Long, ByVal Bf As Long, _
                        R As Long, G As Long, B As Long)

  Dim i As Long
  Dim lER As Long, lEB As Long, lEG As Long
  Dim lMinER As Long, lMinEB As Long, lMinEG As Long
  Dim lMinIndex As Long

    lMinER = 255
    lMinEG = 255
    lMinEB = 255

    For i = 0 To Pal.nColors - 1
        With Pal.Color(i)
            If (Rf = .Rp) And (Gf = .Gp) And (Bf = .Bp) Then
                lMinIndex = i
                Exit For
              Else
                lER = Abs(Rf - .Rp)
                lEG = Abs(Gf - .Gp)
                lEB = Abs(Bf - .Bp)
                If (lER + lEG + lEB < lMinER + lMinEG + lMinEB) Then
                    lMinER = lER
                    lMinEG = lEG
                    lMinEB = lEB
                    lMinIndex = i
                End If
            End If
        End With
    Next i

    R = Pal.Color(lMinIndex).Rp
    G = Pal.Color(lMinIndex).Gp
    B = Pal.Color(lMinIndex).Bp

End Sub

Public Sub LoadPalette256(ByVal FileName As String)

  Dim ff As Integer

    ReDim Pal.Color(255)
    ff = FreeFile
    Open FileName For Binary Access Read As #ff
      Get #ff, , Pal
    Close #ff
    ff = 0
    ReDim Preserve Pal.Color(Pal.nColors - 1)

End Sub

Public Sub SavePalette256(ByVal FileName As String)

  Dim ff As Integer

    ff = FreeFile
    Open FileName For Binary As #ff
      Put #ff, , Pal
    Close #ff
    ff = 0

End Sub

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 11:43:51) 24 + 65 = 89 Lines
