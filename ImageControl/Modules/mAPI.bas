Attribute VB_Name = "mAPI"
'//
'// Main API declarations
'//

Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                        (pDst As Any, _
                        pSrc As Any, _
                        ByVal ByteLen As Long)

Public Declare Function SetStretchBltMode Lib "gdi32" _
                        (ByVal hDC As Long, _
                        ByVal nStretchMode As Long) As Long

Public Declare Function GetDIBits Lib "gdi32" _
                        (ByVal aHDC As Long, _
                        ByVal hBitmap As Long, _
                        ByVal nStartScan As Long, ByVal nNumScans As Long, _
                        lpBits As Any, lpbi As BITMAPINFO, _
                        ByVal wUsage As Long) As Long

Public Declare Function StretchDIBits Lib "gdi32" _
                        (ByVal hDC As Long, _
                        ByVal x As Long, ByVal y As Long, _
                        ByVal dx As Long, ByVal dy As Long, _
                        ByVal SrcX As Long, ByVal SrcY As Long, _
                        ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, _
                        lpBits As Any, lpBitsInfo As BITMAPINFO, _
                        ByVal wUsage As Long, _
                        ByVal dwRop As Long) As Long

Public Declare Function SetDIBitsToDevice Lib "gdi32" _
                        (ByVal hDC As Long, _
                        ByVal x As Long, ByVal y As Long, _
                        ByVal dx As Long, ByVal dy As Long, _
                        ByVal SrcX As Long, ByVal SrcY As Long, _
                        ByVal Scan As Long, ByVal NumScans As Long, _
                        Bits As Any, BitsInfo As BITMAPINFO, _
                        ByVal wUsage As Long) As Long

Public Declare Function BitBlt Lib "gdi32" _
                        (ByVal hDestDC As Long, _
                        ByVal x As Long, ByVal y As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long, _
                        ByVal hSrcDC As Long, _
                        ByVal xSrc As Long, ByVal ySrc As Long, _
                        ByVal dwRop As Long) As Long

Public Const SRCCOPY As Long = &HCC0020
Public Const BLACKNESS As Long = &H42

Public Declare Function CreateCompatibleBitmap Lib "gdi32" _
                        (ByVal hDC As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
                        (ByVal hDC As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
                        (ByVal hDC As Long) As Long

Public Declare Function SelectObject Lib "gdi32" _
                        (ByVal hDC As Long, _
                        ByVal hObject As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
                        (ByVal hObject As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function SetRect Lib "user32" _
                        (lpRect As RECT, _
                        ByVal x1 As Long, ByVal y1 As Long, _
                        ByVal x2 As Long, ByVal y2 As Long) As Long

Public Declare Function IsRectEmpty Lib "user32" _
                        (lpRect As RECT) As Long

Public Declare Function SetRectEmpty Lib "user32" _
                        (lpRect As RECT) As Long

Public Declare Function IntersectRect Lib "user32" _
                        (lpDestRect As RECT, _
                        lpSrc1Rect As RECT, _
                        lpSrc2Rect As RECT) As Long

Public Declare Function InflateRect Lib "user32" _
                        (lpRect As RECT, _
                        ByVal dx As Long, _
                        ByVal dy As Long) As Long

Public Declare Function OffsetRect Lib "user32" _
                        (lpRect As RECT, _
                        ByVal x As Long, _
                        ByVal y As Long) As Long

Public Declare Function GetCursorPos Lib "user32" _
                        (lpPoint As POINTAPI) As Long

Public Declare Function PtInRect Lib "user32" _
                        (lpRect As RECT, _
                        ByVal x As Long, ByVal y As Long) As Long

Public Declare Function GetWindowRect Lib "user32" _
                        (ByVal hwnd As Long, _
                        lpRect As RECT) As Long

Public Declare Function CreateEllipticRgn Lib "gdi32" _
                        (ByVal x1 As Long, ByVal y1 As Long, _
                        ByVal x2 As Long, ByVal y2 As Long) As Long

Public Declare Function PtInRegion Lib "gdi32" _
                        (ByVal hRgn As Long, _
                        ByVal x As Long, ByVal y As Long) As Long

Public Declare Function ClipCursor Lib "user32" _
                        (lpRect As Any) As Long

Public Declare Function ClientToScreen Lib "user32" _
                        (ByVal hwnd As Long, _
                        lpPoint As POINTAPI) As Long

Public Declare Function OpenClipboard Lib "user32" _
                        (ByVal hwnd As Long) As Long

Public Declare Function CloseClipboard Lib "user32" () As Long

Public Declare Function SetClipboardData Lib "user32" _
                        (ByVal wFormat As Long, ByVal hMem As Long) As Long

Public Declare Function EmptyClipboard Lib "user32" () As Long

Public Const CF_BITMAP As Long = 2

Public Declare Function CreatePolygonRgn Lib "gdi32" _
                        (lpPoint As Any, _
                        ByVal nCount As Long, _
                        ByVal nPolyFillMode As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                        (ByVal hwnd As Long, _
                        ByVal wMsg As Long, _
                        ByVal wParam As Long, lParam As Any) As Long

Public Const LB_SELECTSTRING As Long = &H18C

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Declare Function ScrollDC Lib "user32" _
                        (ByVal hDC As Long, _
                        ByVal dx As Long, ByVal dy As Long, _
                        lprcScroll As RECT, _
                        lprcClip As RECT, _
                        ByVal hrgnUpdate As Long, _
                        ByVal lprcUpdate As Long) As Long

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" _
                        (ByVal hDC As Long, _
                        ByVal lpStr As String, ByVal nCount As Long, _
                        lpRect As RECT, _
                        ByVal wFormat As Long) As Long

Public Const DT_CENTER As Long = &H1
Public Const DT_CALCRECT As Long = &H400
Public Const DT_LEFT As Long = &H0
Public Const DT_RIGHT As Long = &H2

Public Declare Function GetPixel Lib "gdi32" _
                        (ByVal hDC As Long, _
                        ByVal x As Long, ByVal y As Long) As Long

Public Declare Function PathCompactPath Lib "shlwapi.dll" Alias "PathCompactPathA" _
                        (ByVal hDC As Long, _
                        ByVal pszPath As String, _
                        ByVal dx As Long) As Long

Public Declare Function SetCapture Lib "user32" _
                        (ByVal hwnd As Long) As Long





Public Function Split(Expression As String, Delimiter As String) As String()

  Dim i As Long
  Dim lenDel As Long
  Dim lstPos As Long, newPos As Long
  Dim tmpArr() As String
  Dim lstDim As Long

    lstPos = 1
    newPos = InStr(Expression, Delimiter)
    lenDel = Len(Delimiter)

    Do While (newPos > 0)
        ReDim Preserve tmpArr(lstDim)
        tmpArr(lstDim) = Mid(Expression, lstPos, newPos - lstPos)
        lstDim = lstDim + 1
        lstPos = newPos + lenDel
        newPos = InStr(lstPos, Expression, Delimiter)
    Loop

    ReDim Preserve tmpArr(lstDim)
    tmpArr(lstDim) = Mid(Expression, lstPos)

    Split = tmpArr

End Function

Public Function GetArrayDim(Arr2D As Variant) As Integer

    On Error Resume Next
      GetArrayDim = UBound(Arr2D, 1) + 1
    On Error GoTo 0
    
End Function

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 11:22:44) 186 + 36 = 222 Lines
