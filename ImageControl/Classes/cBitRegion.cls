VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// cBitRegion Class
'// Carles P.V. - 2002
'//
'// -----------------------------------------------
'// Special thanks to vbAPI.com & vbaccelerator.com
'// -----------------------------------------------

Option Explicit

'// Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As XFORM, ByVal nCount As Long, lpRgnData As RGNDATA) As Long
'// Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As RGNDATA) As Long

'// Private Type XFORM
'//     eM11 As Double
'//     eM12 As Double
'//     eM21 As Double
'//     eM22 As Double
'//     eDx As Double
'//     eDy As Double
'// End Type

'// Private Type RGNDATAHEADER
'//     dwSize As Long
'//     iType As Long
'//     nCount As Long
'//     nRgnSize As Long
'//     rcBound As RECT    -> Region bound rectangle
'// End Type

'// Private Type RGNDATA
'//     rdh As RGNDATAHEADER
'//     Buffer() As Byte   -> Rectangles coord. (Left-Top -> Right-Bottom)
'// End Type

Private Declare Function ExtCreateRegion Lib "gdi32" ( _
                         lpXform As Any, _
                         ByVal nCount As Long, _
                         lpRgnData As Any) As Long

Private Declare Function GetRegionData Lib "gdi32" ( _
                         ByVal hRgn As Long, _
                         ByVal dwCount As Long, _
                         lpRgnData As Any) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function CreateEllipticRgn Lib "gdi32" ( _
                         ByVal x1 As Long, ByVal y1 As Long, _
                         ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" ( _
                         ByVal x1 As Long, ByVal y1 As Long, _
                         ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function CreatePolygonRgn Lib "gdi32" ( _
                         lpPoint As Any, _
                         ByVal nCount As Long, _
                         ByVal nPolyFillMode As Long) As Long

Private Const ALTERNATE = 1
Private Const WINDING = 2

Private Declare Function IsRectEmpty Lib "user32" ( _
                         lpRect As RECT) As Long

Private Declare Function FrameRgn Lib "gdi32" ( _
                         ByVal hdc As Long, _
                         ByVal hRgn As Long, _
                         ByVal hBrush As Long, _
                         ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function OffsetRgn Lib "gdi32" ( _
                         ByVal hRgn As Long, _
                         ByVal dx As Long, ByVal dy As Long) As Long

Private Declare Function GetRgnBox Lib "gdi32" ( _
                         ByVal hRgn As Long, _
                         lpRect As RECT) As Long

Private Const NULLREGION = 1
Private Const SIMPLEREGION = 2
Private Const COMPLEXREGION = 3

Private Declare Function CombineRgn Lib "gdi32" ( _
                         ByVal hDestRgn As Long, _
                         ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, _
                         ByVal nCombineMode As Long) As Long

Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_DIFF = 4
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_MAX = RGN_COPY
Private Const RGN_MIN = RGN_AND

Public Enum CombineRgnMode
    [RGN_Add] = RGN_OR
    [RGN_Subtract] = RGN_DIFF
End Enum

Private Declare Function SelectObject Lib "gdi32" ( _
                         ByVal hdc As Long, ByVal _
                         hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" ( _
                         ByVal hObject As Long) As Long

Private Declare Function CreatePatternBrush Lib "gdi32" ( _
                         ByVal hBitmap As Long) As Long

Private Declare Function CreateBitmap Lib "gdi32" ( _
                         ByVal nWidth As Long, ByVal nHeight As Long, _
                         ByVal nPlanes As Long, _
                         ByVal nBitCount As Long, _
                         lpBits As Integer) As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private m_RegionMain As Long            '// Main region
Private m_Region As Long                '// Original region
Private m_RegionScaled As Long          '// Scaled region

Private m_ZoomFactor As Single          '// Scale factor (>= 0.01)

Private m_BufferData() As Long          '// Original region data (1:1)
Private m_BufferDataScaled() As Long    '// Scaled region data   (ZoomFactor:1)

Private m_hBrush(7) As Long             '// Countour region...
Private m_nBrush As Long                '// Cureent brush index





Private Sub Class_Initialize()

  Dim i As Long, j As Long
  Dim hBitmap As Long
  Dim tBytes(1 To 8) As Integer
  Dim bBytes(1 To 8) As Integer

    tBytes(1) = 240
    tBytes(2) = 120
    tBytes(3) = 60
    tBytes(4) = 30
    tBytes(5) = 15
    tBytes(6) = 135
    tBytes(7) = 195
    tBytes(8) = 225

    hBitmap = CreateBitmap(8, 8, 1, 1, tBytes(1))
    m_hBrush(0) = CreatePatternBrush(hBitmap)
    DeleteObject hBitmap

    For i = 7 To 1 Step -1
        For j = 1 To 8
            If (i + j < 9) Then
                bBytes(j) = tBytes(i + j)
              Else
                bBytes(j) = tBytes(i + j - 8)
            End If
        Next j
        hBitmap = CreateBitmap(8, 8, 1, 1, bBytes(1))
        m_hBrush(i) = CreatePatternBrush(hBitmap)
        DeleteObject hBitmap
    Next i

    Erase tBytes
    Erase bBytes

    m_ZoomFactor = 1

End Sub

Private Sub Class_Terminate()

  Dim i As Long

    For i = 0 To 7
        DeleteObject m_hBrush(i)
    Next i
    Clear

End Sub

Public Sub Init(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

    Clear
    m_RegionMain = CreateRectRgn(x1, y1, x2, y2)

End Sub

Public Sub CreateRect(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Mode As CombineRgnMode, Optional ByVal Crop As Boolean = 0)

  Dim tmpRegion As Long

    If (m_Region <> 0) Then
        tmpRegion = CreateRectRgn(x1, y1, x2, y2)
        CombineRgn m_Region, m_Region, tmpRegion, Mode
        DeleteObject tmpRegion
      Else
        If (Mode <> RGN_Subtract) Then
            m_Region = CreateRectRgn(x1, y1, x2, y2)
        End If
    End If

    If (Crop) Then CombineRgn m_Region, m_Region, m_RegionMain, RGN_AND

    GetRgnData
    ScaleRgn m_ZoomFactor

End Sub

Public Sub CreateEllipse(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Mode As CombineRgnMode, Optional ByVal Crop As Boolean = 0)

  Dim tmpRegion As Long

    If (m_Region <> 0) Then
        tmpRegion = CreateEllipticRgn(x1, y1, x2, y2)
        CombineRgn m_Region, m_Region, tmpRegion, Mode
        DeleteObject tmpRegion
      Else
        If (Mode <> RGN_Subtract) Then
            m_Region = CreateEllipticRgn(x1, y1, x2, y2)
        End If
    End If

    If (Crop) Then CombineRgn m_Region, m_Region, m_RegionMain, RGN_AND

    GetRgnData
    ScaleRgn m_ZoomFactor

End Sub

Public Sub CreatePolygon(ByRef lPoints() As Long, ByVal Mode As CombineRgnMode, Optional ByVal Crop As Boolean = 0)

  Dim tmpRegion As Long

    If (UBound(lPoints) < 2) Then Exit Sub

    If (m_Region <> 0) Then
        tmpRegion = CreatePolygonRgn(lPoints(0), (UBound(lPoints) + 1) \ 2, WINDING)
        CombineRgn m_Region, m_Region, tmpRegion, Mode
        DeleteObject tmpRegion
      Else
        If (Mode <> RGN_Subtract) Then
            m_Region = CreatePolygonRgn(lPoints(0), (UBound(lPoints) + 1) \ 2, WINDING)
        End If
    End If

    If (Crop) Then CombineRgn m_Region, m_Region, m_RegionMain, RGN_AND

    GetRgnData
    ScaleRgn m_ZoomFactor

End Sub

Public Sub Invert()

  Dim rgnRct1 As RECT
  Dim rgnRct2 As RECT

    GetRgnBox m_Region, rgnRct1
    GetRgnBox m_RegionMain, rgnRct2

    If (IsRectEmpty(rgnRct1)) Then
        With rgnRct2
            m_Region = CreateRectRgn(.Left, .Top, .Right, .Bottom)
        End With
      Else
        CombineRgn m_Region, m_RegionMain, m_Region, RGN_XOR
    End If

    GetRgnData
    ScaleRgn m_ZoomFactor

End Sub

Public Sub Crop()

    CombineRgn m_Region, m_RegionMain, m_Region, RGN_AND

    GetRgnData
    ScaleRgn m_ZoomFactor

End Sub

Public Sub Offset(ByVal dx As Long, ByVal dy As Long)

    OffsetRgn m_Region, dx, dy

    GetRgnData
    ScaleRgn m_ZoomFactor

End Sub

Public Sub Clear()
    
    If (m_RegionMain <> 0) Then
        DeleteObject m_RegionMain
        m_RegionMain = 0
    End If
    If (m_Region <> 0) Then
        DeleteObject m_Region
        m_Region = 0
    End If
    If (m_RegionScaled <> 0) Then
        DeleteObject m_RegionScaled
        m_RegionScaled = 0
    End If

    Erase m_BufferData
    Erase m_BufferDataScaled

End Sub

Public Sub DrawToDC(ByVal hdc As Long, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0)

    OffsetRgn m_RegionScaled, xOffset, yOffset
    FrameRgn hdc, m_RegionScaled, m_hBrush(m_nBrush), 1, 1
    OffsetRgn m_RegionScaled, -xOffset, -yOffset

End Sub

Public Sub RotateBrush()
  
    m_nBrush = m_nBrush + 1
    If (m_nBrush = 8) Then
        m_nBrush = 0
    End If

End Sub

Public Function BoundRect(ByRef x1 As Long, ByRef y1 As Long, ByRef x2 As Long, ByRef y2 As Long) As Boolean

  Dim lType As Long
  Dim rgnRct As RECT

    lType = GetRgnBox(m_Region, rgnRct)

    If (IsRectEmpty(rgnRct)) Then '// ~ lType = NULLREGION ?
        BoundRect = 0
      Else
        With rgnRct
            x1 = .Left    '// = m_BufferData(4)
            y1 = .Top     '// = m_BufferData(5)
            x2 = .Right   '// = m_BufferData(6)
            y2 = .Bottom  '// = m_BufferData(7)
            BoundRect = -1
        End With
    End If

End Function

Public Function IsEmptyRegion() As Boolean

  Dim rgnRct As RECT

    GetRgnBox m_Region, rgnRct

    If (IsRectEmpty(rgnRct)) Then
        IsEmptyRegion = -1
      Else
        IsEmptyRegion = 0
    End If

End Function

Public Sub LoadFromFile(ByVal FileName As String)

  Dim ff As Integer

    ff = FreeFile
    Open FileName For Binary Access Read As #ff
    ReDim m_BufferData(0 To LOF(ff) \ 4 - 1) As Long
      Get #ff, , m_BufferData
    Close #ff
    ff = 0

    m_Region = ExtCreateRegion(ByVal 0&, UBound(m_BufferData) + 1, m_BufferData(0))
    ScaleRgn m_ZoomFactor

End Sub

Public Sub SaveToFile(ByVal FileName As String)

  Dim ff As Integer

    ff = FreeFile
    Open FileName For Binary Access Write As #ff
      Put #ff, , m_BufferData
    Close #ff
    ff = 0

End Sub

Public Function GetData() As Long()

    GetData = m_BufferData

End Function

Public Function GetDataCount() As Long

    On Error Resume Next
      GetDataCount = UBound(m_BufferData) + 1
    On Error GoTo 0
    
End Function

'Public Sub SetData(ByRef Data() As Long)
'    m_BufferData = Data
'    m_Region = ExtCreateRegion(ByVal 0&, UBound(m_BufferData) + 1, m_BufferData(0))
'    ScaleRgn m_ZoomFactor
'End Sub

Public Property Let ZoomFactor(ByVal Factor As Single)

    If (Factor < 0) Then
        m_ZoomFactor = 0.01
      Else
        m_ZoomFactor = Factor
    End If
    ScaleRgn m_ZoomFactor

End Property

Public Property Get ZoomFactor() As Single

    ZoomFactor = m_ZoomFactor

End Property

Public Property Get Region() As Long

    Region = m_Region

End Property

Public Property Let Region(ByVal lRgn As Long)

    GetExtRgnData lRgn
    ScaleRgn m_ZoomFactor

End Property

Public Property Get RegionScaled() As Long

    RegionScaled = m_RegionScaled

End Property

Public Property Get RegionMain() As Long

    RegionMain = m_RegionMain

End Property

Private Sub GetRgnData()

  Dim lCount As Long
    
    If (m_Region <> 0) Then
        
        lCount = GetRegionData(m_Region, 0, ByVal 0&)
        ReDim m_BufferData(lCount \ 4 - 1)
        GetRegionData m_Region, lCount, m_BufferData(0)
    
        ScaleRgn m_ZoomFactor
        
    End If

End Sub

Private Sub GetExtRgnData(ByVal lRgn As Long)

  Dim lCount As Long
    
    If (lRgn <> 0) Then
    
        lCount = GetRegionData(lRgn, 0, ByVal 0&)
        ReDim m_BufferData(lCount \ 4 - 1)
        GetRegionData lRgn, lCount, m_BufferData(0)
    
        m_Region = ExtCreateRegion(ByVal 0&, UBound(m_BufferData) + 1, m_BufferData(0))
        
    End If

End Sub

Private Sub ScaleRgn(ByVal Factor As Single)

  Dim i As Long
    
    If (m_Region <> 0) Then

        m_BufferDataScaled = m_BufferData

        If (Factor > 0) Then
            '// 0 - 3   : Header
            '// 4 - 7   : rcBound
            '// 8 - End : Region rects. collection (Buffer)
            For i = 4 To UBound(m_BufferData)
                m_BufferDataScaled(i) = m_BufferData(i) * Factor
            Next i
        End If
        
        If (m_RegionScaled <> 0) Then
            DeleteObject m_RegionScaled
            m_RegionScaled = 0
        End If
        m_RegionScaled = ExtCreateRegion(ByVal 0&, UBound(m_BufferDataScaled) + 1, m_BufferDataScaled(0))
    
    End If
    
End Sub

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 12:03:52) 140 + 367 = 507 Lines
