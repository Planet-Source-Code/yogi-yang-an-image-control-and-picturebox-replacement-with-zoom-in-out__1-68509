Attribute VB_Name = "mBitmap"
'//
'// Bitmap module
'//

Option Explicit

Public Type BITMAPINFOHEADER
    biSize           As Long
    biWidth          As Long
    biHeight         As Long
    biPlanes         As Integer
    biBitCount       As Integer
    biCompression    As Long
    biSizeImage      As Long
    biXPelsPerMeter  As Long
    biYPelsPerMeter  As Long
    biClrUsed        As Long
    biClrImportant   As Long
End Type

Public Type BITMAPINFO
    Header As BITMAPINFOHEADER '// (bmiHeader)
    Bits() As Byte             '// (bmiColors)
End Type

Public Const DIB_RGB_COLORS As Long = 0&

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type




Public Sub CreateBitmap(ByRef srcBM As BITMAPINFO, _
                        ByRef srcBits() As Byte, _
                        ByVal W As Long, ByVal H As Long)

    ReDim srcBM.Bits(3, W - 1, H - 1) As Byte
    With srcBM.Header
        .biSize = 40
        .biBitCount = 32
        .biPlanes = 1
        .biWidth = W
        .biHeight = -H
    End With
    srcBM.Bits = srcBits

End Sub

Public Sub BestFitSize(ByVal srcW As Long, ByVal srcH As Long, _
                       ByVal dstW As Long, ByVal dstH As Long, _
                       ByRef bfW As Long, bfH As Long)

  Dim cW As Single, cH As Single

    If (srcW > dstW) Or (srcH > dstH) Then
        cW = dstW / srcW
        cH = dstH / srcH
        If (cW < cH) Then
            bfW = dstW
            bfH = srcH * cW
          Else
            bfH = dstH
            bfW = srcW * cH
        End If
      Else
        bfW = srcW
        bfH = srcH
    End If

End Sub

Public Function TakeBitsFromPicture(ByRef srcPicture As StdPicture, ByVal srcW As Long, ByVal srcH As Long) As Byte()

  Dim lhDC As Long
  Dim lhBmpOld As Long
  Dim tBM As BITMAPINFO

    lhDC = CreateCompatibleDC(0)
    If (lhDC <> 0) Then
        With tBM.Header
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = srcW
            .biHeight = -srcH
            ReDim tBM.Bits(3, srcW - 1, srcH - 1)
        End With

        lhBmpOld = SelectObject(lhDC, srcPicture.Handle)
        SelectObject lhDC, lhBmpOld
        GetDIBits lhDC, srcPicture.Handle, 0, srcH, tBM.Bits(0, 0, 0), tBM, DIB_RGB_COLORS
        DeleteObject lhDC

        TakeBitsFromPicture = tBM.Bits
        Erase tBM.Bits
    End If

End Function

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 11:38:48) 36 + 69 = 105 Lines
