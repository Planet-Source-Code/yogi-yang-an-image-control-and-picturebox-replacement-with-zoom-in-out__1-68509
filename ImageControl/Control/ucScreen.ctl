VERSION 5.00
Begin VB.UserControl ucScreen 
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   236
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer tmrSelection 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4005
      Top             =   2940
   End
   Begin VB.PictureBox iSrc 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3840
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox iDst 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3210
      Left            =   150
      ScaleHeight     =   214
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   2
      Top             =   150
      Width           =   3345
      Begin VB.Line shpLine 
         BorderColor     =   &H00000000&
         DrawMode        =   6  'Mask Pen Not
         Index           =   0
         Visible         =   0   'False
         X1              =   57
         X2              =   165
         Y1              =   176
         Y2              =   96
      End
      Begin VB.Shape shpEllp 
         BorderColor     =   &H00000000&
         DrawMode        =   6  'Mask Pen Not
         Height          =   1035
         Left            =   1035
         Shape           =   2  'Oval
         Top             =   1500
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Shape shpRect 
         BorderColor     =   &H00000000&
         DrawMode        =   6  'Mask Pen Not
         Height          =   795
         Left            =   780
         Top             =   1050
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.HScrollBar hSB 
      Height          =   195
      Left            =   0
      Max             =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   -195
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.VScrollBar vSB 
      Height          =   2415
      Left            =   -195
      Max             =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image cGrab 
      Height          =   480
      Left            =   3945
      Picture         =   "ucScreen.ctx":0000
      Top             =   135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cRelease 
      Height          =   480
      Left            =   3945
      Picture         =   "ucScreen.ctx":030A
      Top             =   690
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape shp 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000F&
      Height          =   210
      Left            =   2010
      Top             =   2190
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "ucScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// ============================================
'// UC name: Screen [from cpvPicScroll OCX v3.0]
'// Author:  Carles P.V. - 2002
'// Date:    May 14, 2002
'// ============================================

'// Flood Fill function from EDais: http://edais.earlsoft.co.uk
'// Anti-alias algorithm for shapes by dafhi

Option Explicit

Public Enum BarsCts
    [sbAutomatic]
    [sbNone]
End Enum

Private Const pdf_BarsWidth = 13
Private Const pdf_MouseScroll = -1
Private Const pdf_Bars = [sbAutomatic]

Private pBarsWidth As Integer
Private pMouseScroll As Boolean
Private pBars As BarsCts

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Integer, y As Integer)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Integer, y As Integer)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Integer, y As Integer)
Public Event Scroll()

Private tBM As BITMAPINFO           '// Bitmap
Private BMExists As Boolean         '// Bitmap initialized flag
Private BMRect As RECT              '// Bitmap rectangle
Private tBMDone() As Boolean        '// Boolean bits map

Private Zm(1 To 15) As Integer      '// Zoom coeficients array
Private iz As Integer               '// Zoom index array
Private izInit As Integer

'// >Main View
Private Pt As POINTAPI              '// Cursor position
Private tmpPt As POINTAPI           '// Temp. cursor position (anchor point)
Private scrllMain As Boolean        '// Scrolling flag
'// >Scroll bars
Private OffSB As Boolean            '// Enabled/Disabled: scroll bars (Off/On)
Private lastH As Single             '// Last horizontal value
Private lastV As Single             '// Last vertical value
Private lastHM As Single            '// Last horizontal Max. value
Private lastVM As Single            '// Last vertical Max. value

Public BitsRegion As New cBitRegion '// Current region





'// Init/Read/Write properties
'========================================================================================

Private Sub UserControl_InitProperties()

    pBarsWidth = pdf_BarsWidth
    pMouseScroll = pdf_MouseScroll
    pBars = pdf_Bars

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        pBarsWidth = .ReadProperty("BarsWidth", pdf_BarsWidth)
        pMouseScroll = .ReadProperty("MouseScroll", pdf_MouseScroll)
        pBars = .ReadProperty("Bars", pdf_Bars)
        UserControl.BackColor = .ReadProperty("BackColor", &H8000000F)
        iDst.BackColor = .ReadProperty("BackColor", &H8000000F)
    End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "BackColor", iDst.BackColor, &H8000000F
        .WriteProperty "BarsWidth", pBarsWidth, pdf_BarsWidth
        .WriteProperty "MouseScroll", pMouseScroll, pdf_MouseScroll
        .WriteProperty "Bars", pBars, pdf_Bars
    End With

End Sub

'// UserControl Refresh/Resize
'========================================================================================

Private Sub UserControl_Initialize()

  '// Initialize zoom factors
  
  Dim i As Long

    For i = 1 To 15
        Zm(i) = i
    Next i
    iz = 1

End Sub

Private Sub iDst_Paint()

  Dim dstX As Long, dstY As Long
  Dim dstW As Long, dstH As Long
  Dim SrcX As Long, SrcY As Long
  Dim srcW As Long, srcH As Long
  Dim lRet As Long
    
    If (Extender.Visible = 0) Then Exit Sub
    If (Not BMExists) Then Exit Sub

    '// Get dimensions of source area to paint
    If (hSB.Max > 0) Then
        dstX = -hSB Mod Zm(iz)
        dstW = (iDst.Width \ Zm(iz)) * Zm(iz) + 2 * Zm(iz)
        SrcX = (hSB \ Zm(iz)) + 2
        srcW = (iDst.Width \ Zm(iz)) + 2
      Else
        dstX = 0
        dstW = iDst.Width
        SrcX = 2
        srcW = iSrc.Width
    End If

    If (vSB.Max > 0) Then
        dstY = -vSB Mod Zm(iz)
        dstH = ((iDst.Height - 1) \ Zm(iz)) * Zm(iz) + 2 * Zm(iz)
        SrcY = (vSB.Max \ Zm(iz)) - (vSB \ Zm(iz)) + 1
        srcH = ((iDst.Height - 1) \ Zm(iz)) + 2
      Else
        dstY = 0
        dstH = iDst.Height
        SrcY = 2
        srcH = iSrc.Height
    End If

    '// Draw it
    lRet = StretchDIBits(iDst.hdc, dstX, dstY, dstW, dstH, SrcX, SrcY, srcW, srcH, tBM.Bits(0, -2, -2), tBM, DIB_RGB_COLORS, SRCCOPY)
    If (lRet = 0) Then
       'Debug.Print "Unable to zoom " & (iz * 100) & "%"
        ZoomFactor = 1
    End If

    '// Draw region
    If (BitsRegion.Region And tmrSelection.Enabled) Then
        BitsRegion.DrawToDC iDst.hdc, -hSB, -vSB
    End If

    '// Refresh Panoramic View
    'If (fPanView.Visible) Then fPanView.UpdateClipRct

End Sub

Private Sub UserControl_Resize()

  Dim hExt As Boolean, vExt As Boolean
  Dim hVis As Boolean, vVis As Boolean
  Dim sW As Long, sH As Long
  Dim sbW As Long
  Dim i As Integer
  
    If (Not BMExists) Then Exit Sub

    sW = ScaleWidth
    sH = ScaleHeight
    sbW = pBarsWidth

    '// Check values don't exceed Max. integer (aprox.)
    If (iSrc.Width * Zm(iz) > 32767) Or (iSrc.Height * Zm(iz) > 32767) Then
       'Debug.Print "Unable to zoom " & (iz * 100) & "%"
        ZoomFactor = 1
        Exit Sub
    End If

    '// Hide dest. view / disable scroll bars
    iDst.Visible = 0
    Set iDst = Nothing
    OffSB = -1

    '// Resize
    '// Check minimum size
    If (sW < 2 * sbW) Then
        Width = (2 * sbW) * Screen.TwipsPerPixelX
        Exit Sub
    End If
    If (sH < 2 * sbW) Then
        Height = (2 * sbW) * Screen.TwipsPerPixelY
        Exit Sub
    End If

    '// Check if zoomed image exceeds visible area
    For i = 1 To 2
        If (iSrc.Width * Zm(iz) > sW - IIf(vVis, sbW, 0)) Then hExt = -1
        If (iSrc.Height * Zm(iz) > sH - IIf(hVis, sbW, 0)) Then vExt = -1
        '// Show/Hide scroll bars on case
        Select Case pBars
          Case 0
            '// Automatic
            If (hExt) Then
                hSB.Visible = -1
                hVis = -1
              Else
                hSB.Visible = 0
                hVis = 0
            End If
            If (vExt) Then
                vSB.Visible = -1
                vVis = -1
              Else
                vSB.Visible = 0
                vVis = 0
            End If
            If (hVis And vVis) Then shp.Visible = -1
          Case 1
            '// None
            hSB.Visible = 0
            hVis = 0
            vSB.Visible = 0
            vVis = 0
            shp.Visible = 0
        End Select
        hSB.Refresh
        vSB.Refresh
    Next i

    '// Relocate and resize scroll bars:
    hSB.Move 0, sH - sbW, sW - IIf(vVis, sbW, 0), sbW
    vSB.Move sW - sbW, 0, sbW, sH - IIf(hVis, sbW, 0)
    shp.Move hSB.Width, vSB.Height, sbW, sbW

    With iDst
        '// Readjust scroll picture area:
        If (vExt) Then
            .Height = sH - IIf(hVis, sbW, 0)
            .Top = 0
          Else
            .Height = iSrc.Height * Zm(iz)
            .Top = (sH - .Height - IIf(hVis, sbW, 0)) \ 2
        End If
        If (hExt) Then
            .Width = sW - IIf(vVis, sbW, 0)
            .Left = 0
          Else '
            .Width = iSrc.Width * Zm(iz)
            .Left = (sW - .Width - IIf(vVis, sbW, 0)) \ 2
        End If
        '// Readjust scroll bars values
        hSB.Max = iSrc.Width * Zm(iz) - .Width
        vSB.Max = iSrc.Height * Zm(iz) - .Height
        hSB.LargeChange = .Width
        vSB.LargeChange = .Height
        hSB.SmallChange = Zm(iz)
        vSB.SmallChange = Zm(iz)
        '// Zoom 'memory position'
        On Error Resume Next
          If (hExt And lastHM And izInit <> 0) Then
              hSB = lastH * hSB.Max / lastHM
            Else
              hSB = hSB.Max / 2
          End If
          If (vExt And lastVM And izInit <> 0) Then
              vSB = lastV * vSB.Max / lastVM
            Else
              vSB = vSB.Max / 2
          End If
          lastHM = hSB.Max
          lastVM = vSB.Max
          izInit = -1
          '// Set MousePointer
          If (pMouseScroll And (hExt Or vExt)) Then
              .MousePointer = vbCustom
              .MouseIcon = cRelease
            Else
              .MousePointer = vbDefault
          End If
      End With

      '// Refresh dest. view / Enable scroll bars
      iDst.Visible = -1
      OffSB = 0
      
      On Error GoTo 0
        
End Sub

Private Sub UserControl_Terminate()

    Set BitsRegion = Nothing

    Set iSrc = Nothing
    Set iDst = Nothing
    Erase tBM.Bits
    ClipCursor ByVal 0&

    tmrSelection.Enabled = 0

End Sub

'// Region / temp. selection shapes drawing
'========================================================================================

Private Sub tmrSelection_Timer()

    If (Not BitsRegion.IsEmptyRegion) Then
        BitsRegion.RotateBrush
        BitsRegion.DrawToDC iDst.hdc, -hSB, -vSB
    End If

End Sub

Public Sub ShowShape(ByVal shpType As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Optional ByVal NewLine As Boolean = 0)

    x1 = x1 * iz - hSB
    y1 = y1 * iz - vSB
    x2 = x2 * iz - hSB
    y2 = y2 * iz - vSB

    Select Case shpType
      Case 0  '// Rectangle
        shpRect.Visible = -1
        shpRect.Move x1, y1, x2 - x1, y2 - y1
      Case 1  '// Ellipse
        shpEllp.Visible = -1
        shpEllp.Move x1, y1, x2 - x1, y2 - y1
      Case 2  '// Polygon
        If (shpLine.Count > 0) Then
            shpLine(shpLine.Count - 1).Visible = -1
            With shpLine(shpLine.Count - 1)
                .x1 = x1
                .x2 = x2
                .y1 = y1
                .y2 = y2
            End With
            If (NewLine) Then Load shpLine(shpLine.Count)
        End If
    End Select

End Sub

Public Sub HideShape(ByVal shpType As Long)

  Dim i As Long
    
    Select Case shpType
      Case 0  '// Rectangle
        shpRect.Visible = 0
      Case 1  '// Ellipse
        shpEllp.Visible = 0
      Case 2  '// Polygon
        For i = 1 To shpLine.Count - 1
            Unload shpLine(i)
        Next i
        shpLine(0).Visible = 0
    End Select

End Sub

Public Sub StartTimer()

    tmrSelection_Timer
    tmrSelection.Enabled = -1

End Sub

Public Sub StopTimer()

    tmrSelection.Enabled = 0

End Sub

Public Sub SetAutoRedraw(ByVal Auto As Boolean)

    iDst.AutoRedraw = Auto
    If (Auto) Then
        UpdateImage
        iDst.Refresh
    End If

End Sub

'// Scroll bars
'========================================================================================

Public Sub Scroll(ByVal hV As Integer, vV As Integer)

    OffSB = -1
    hSB = hV
    vSB = vV
    OffSB = 0
    iDst_Paint

End Sub

Private Sub hSB_GotFocus()

    iDst.SetFocus

End Sub

Private Sub hSB_Change()

    lastH = hSB
    If (Not OffSB) Then
        RaiseEvent Scroll
        iDst_Paint
    End If

End Sub

Private Sub hSB_Scroll()

    hSB_Change

End Sub

Private Sub vSB_GotFocus()

    iDst.SetFocus

End Sub

Private Sub vSB_Change()

    lastV = vSB
    If (Not OffSB) Then
        RaiseEvent Scroll
        iDst_Paint
    End If

End Sub

Private Sub vSB_Scroll()

    vSB_Change

End Sub

'// Events/Scrolling
'========================================================================================
Private Sub iDst_Click()

    RaiseEvent Click

End Sub

Private Sub iDst_DblClick()

    RaiseEvent DblClick

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    x = x - iDst.Left
    y = y - iDst.Top
    
    SetCapture iDst.hwnd
    iDst_MouseDown Button, Shift, x, y
    
End Sub

Private Sub iDst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (Button = 1) Then
        iDst.MouseIcon = cGrab
        scrllMain = -1
        GetCursorPos Pt
        tmpPt = Pt
    End If
    RaiseEvent MouseDown(Button, Shift, (hSB + x) \ Zm(iz), (vSB + y) \ Zm(iz))

End Sub

Private Sub iDst_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  Static lstPt As POINTAPI
  Dim xInc As Long, yInc As Long
  Dim bMoved As Boolean

    If ((hSB + x) \ Zm(iz) = lstPt.x) And ((vSB + y) \ Zm(iz) = lstPt.y) Then
        bMoved = 0
      Else
        lstPt.x = (hSB + x) \ Zm(iz)
        lstPt.y = (vSB + y) \ Zm(iz)
        bMoved = -1
    End If

    If (Button <> 1 Or Not BMExists Or pMouseScroll = 0 Or scrllMain = 0) Then
        If (bMoved) Then RaiseEvent MouseMove(Button, Shift, (hSB + x) \ Zm(iz), (vSB + y) \ Zm(iz))
        Exit Sub
    End If


    OffSB = -1
    GetCursorPos Pt
    xInc = Pt.x - tmpPt.x
    yInc = Pt.y - tmpPt.y

    If (hSB.Max) Then
        If (xInc > 0) Then
            If (hSB - xInc > 0) Then
                hSB = hSB - xInc
              Else
                hSB = 0
            End If
          Else
            If (hSB - xInc < hSB.Max) Then
                hSB = hSB - xInc
              Else
                hSB = hSB.Max
            End If
        End If
    End If
    If (vSB.Max) Then
        If (yInc > 0) Then
            If (vSB - yInc > 0) Then
                vSB = vSB - yInc
              Else
                vSB = 0
            End If
          Else
            If (vSB - yInc < vSB.Max) Then
                vSB = vSB - yInc
              Else
                vSB = vSB.Max
            End If
        End If
    End If

    tmpPt = Pt
    OffSB = 0

    If (hSB.Max Or vSB.Max) Then
        iDst_Paint
        RaiseEvent Scroll
    End If
    If (bMoved) Then RaiseEvent MouseMove(Button, Shift, (hSB + x) \ Zm(iz), (vSB + y) \ Zm(iz))

End Sub

Private Sub iDst_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    iDst.MouseIcon = cRelease
    scrllMain = 0
    RaiseEvent MouseUp(Button, Shift, (hSB + x) \ Zm(iz), (vSB + y) \ Zm(iz))

End Sub

Private Sub iDst_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub iDst_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub iDst_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

'// Get scroll bar Values and Source/Dest. image dimensions
'// (Panoramic view)
'========================================================================================

Public Sub GetDimensions(ByRef sW As Integer, sH As Integer, dW As Integer, dH As Integer)

    sW = iSrc.Width
    sH = iSrc.Height
    dW = iDst.Width
    dH = iDst.Height

End Sub

Public Sub GetSBV(ByRef hV As Integer, vV As Integer)

    hV = hSB
    vV = vSB

End Sub

Public Sub GetSBM(ByRef hM As Integer, vM As Integer)

    hM = hSB.Max
    vM = vSB.Max

End Sub

'// Methods
'========================================================================================

Public Sub Clear()

  '// Clear main view and restore cursor

    iDst.Visible = 0
    iDst.MousePointer = vbDefault
    '// Reset scroll bars
    OffSB = -1
    hSB.Visible = 0
    vSB.Visible = 0
    hSB.Max = 0
    vSB.Max = 0
    shp.Visible = 0
    '// Delete source picture and Bits
    Set iSrc = Nothing
    Erase tBM.Bits
    BMExists = 0

End Sub

Public Sub UpdateImage()

    iDst_Paint

End Sub

Public Sub UpdateRegion()

    BitsRegion.DrawToDC iDst.hdc, -hSB, -vSB

End Sub

Public Sub ZoomIn()

    If (BMExists) Then
        If (iz < UBound(Zm)) Then
            iz = iz + 1
            BitsRegion.ZoomFactor = iz
            UserControl_Resize
        End If
    End If
    
End Sub

Public Sub ZoomOut()

    If (BMExists) Then
        If (iz > LBound(Zm)) Then
            iz = iz - 1
            BitsRegion.ZoomFactor = iz
            UserControl_Resize
        End If
    End If
End Sub

'// Paint functions
'========================================================================================

Public Sub CreateBlank(ByVal dstWidth As Long, ByVal dstHeight As Long)

    If (dstWidth < 1) Or (dstHeight < 1) Then Exit Sub

    With tBM.Header
        .biSize = 40
        .biWidth = dstWidth + 4
        .biHeight = -dstHeight - 4
        .biPlanes = 1
        .biBitCount = 32
        ReDim tBM.Bits(3, -2 To (dstWidth - 1) + 2, -2 To (dstHeight - 1) + 2)
    End With

    iSrc.Move 0, 0, dstWidth, dstHeight
    SetRect BMRect, 0, 0, dstWidth, dstHeight

    BMExists = -1
    UserControl_Resize

End Sub

Public Function BitsGetPixel(ByVal x As Long, ByVal y As Long) As Long

    If (BMExists) Then
        If (PtInRect(BMRect, x, y)) Then
            BitsGetPixel = RGB(tBM.Bits(2, x, y), tBM.Bits(1, x, y), tBM.Bits(0, x, y))
        End If
    Else
        BitsGetPixel = -1
    End If

End Function

Public Sub BitsPaint(ByVal x As Long, ByVal y As Long, ByRef mBits() As Byte, ByVal MaskColor As Long, ByVal Pressure As Long, Optional ByVal InRegion As Boolean = -1)

  Dim lRegion As Long

  Dim srcRct As RECT
  Dim mskRct As RECT
  Dim srcW As Long, srcH As Long
  Dim mskW As Long, mskH As Long
  Dim i As Long, j As Long
  Dim iIn As Long, jIn As Long
  Dim Rm As Long, Gm As Long, Bm As Long
  Dim cp1 As Single, cp2 As Single

    If (Not BMExists) Then Exit Sub
    
    srcW = UBound(tBM.Bits, 2) - 2
    srcH = UBound(tBM.Bits, 3) - 2
    mskW = UBound(mBits, 2) - 2
    mskH = UBound(mBits, 3) - 2
    If (srcW = 0 Or srcH = 0) Then Exit Sub
    If (mskW = 0 Or mskH = 0) Then Exit Sub

    SetRect srcRct, 0, 0, srcW, srcH
    SetRect mskRct, x, y, x + mskW, y + mskH
    IntersectRect srcRct, mskRct, srcRct
    If (IsRectEmpty(srcRct)) Then Exit Sub

    If (MaskColor <> -1) Then
        Rm = (MaskColor And &HFF&)
        Gm = (MaskColor And &HFF00&) \ 256
        Bm = (MaskColor And &HFF0000) \ 65536
      Else
        Rm = -1
    End If

    cp1 = Pressure / 100
    cp2 = 1 - cp1

    If (InRegion) Then
        lRegion = AdjustRegion
      Else
        lRegion = BitsRegion.RegionMain
    End If

    For i = srcRct.Left To srcRct.Right
        For j = srcRct.Top To srcRct.Bottom
            iIn = i - x
            jIn = j - y
            If (PtInRegion(lRegion, i, j)) Then
                If (mBits(2, iIn, jIn) = Rm And mBits(1, iIn, jIn) = Gm And mBits(0, iIn, jIn) = Bm) Then
                  Else
                    tBM.Bits(0, i, j) = cp1 * mBits(0, iIn, jIn) + cp2 * tBM.Bits(0, i, j)
                    tBM.Bits(1, i, j) = cp1 * mBits(1, iIn, jIn) + cp2 * tBM.Bits(1, i, j)
                    tBM.Bits(2, i, j) = cp1 * mBits(2, iIn, jIn) + cp2 * tBM.Bits(2, i, j)
                End If
            End If
        Next j
    Next i

End Sub

Public Sub BitsPaintPreMasked(ByVal x As Long, ByVal y As Long, ByRef mBits() As Byte, ByVal Pressure As Long, Optional ByVal InRegion As Boolean = -1)

  Dim lRegion As Long

  Dim srcRct As RECT
  Dim mskRct As RECT
  Dim srcW As Long, srcH As Long
  Dim mskW As Long, mskH As Long
  Dim i As Long, j As Long
  Dim iIn As Long, jIn As Long
  Dim cp1 As Single, cp2 As Single
    
    If (Not BMExists) Then Exit Sub

    srcW = UBound(tBM.Bits, 2) - 2
    srcH = UBound(tBM.Bits, 3) - 2
    mskW = UBound(mBits, 2) - 2
    mskH = UBound(mBits, 3) - 2
    If (srcW = 0 Or srcH = 0) Then Exit Sub
    If (mskW = 0 Or mskH = 0) Then Exit Sub

    SetRect srcRct, 0, 0, srcW, srcH
    SetRect mskRct, x, y, x + mskW, y + mskH
    IntersectRect srcRct, mskRct, srcRct
    If (IsRectEmpty(srcRct)) Then Exit Sub

    cp1 = Pressure / 100
    cp2 = 1 - cp1

    If (InRegion) Then
        lRegion = AdjustRegion
      Else
        lRegion = BitsRegion.RegionMain
    End If

    For i = srcRct.Left To srcRct.Right
        For j = srcRct.Top To srcRct.Bottom
            iIn = i - x
            jIn = j - y
            If (PtInRegion(lRegion, i, j)) Then
                If (mBits(3, iIn, jIn) < 255) Then
                    tBM.Bits(0, i, j) = cp1 * mBits(0, iIn, jIn) + cp2 * tBM.Bits(0, i, j)
                    tBM.Bits(1, i, j) = cp1 * mBits(1, iIn, jIn) + cp2 * tBM.Bits(1, i, j)
                    tBM.Bits(2, i, j) = cp1 * mBits(2, iIn, jIn) + cp2 * tBM.Bits(2, i, j)
                End If
            End If
        Next j
    Next i

End Sub

Public Sub BitsBWMask(ByVal x As Long, ByVal y As Long, ByRef mBits() As Byte, ByVal Color As Long, Optional ByVal Pressure As Long = 100)
'// Back transparent color: &HBEBEBE (grey 190)

  Dim lRegion As Long

  Dim srcRct As RECT
  Dim mskRct As RECT
  Dim srcW As Long, srcH As Long
  Dim mskW As Long, mskH As Long
  Dim i As Long, j As Long
  Dim iIn As Long, jIn As Long
  Dim R As Long, G As Long, B As Long
  Dim cP As Single
  Dim cText As Single, cBack As Single

    If (Not BMExists) Then Exit Sub

    srcW = UBound(tBM.Bits, 2) - 2
    srcH = UBound(tBM.Bits, 3) - 2
    mskW = UBound(mBits, 2) - 2
    mskH = UBound(mBits, 3) - 2
    If (srcW = 0 Or srcH = 0) Then Exit Sub
    If (mskW = 0 Or mskH = 0) Then Exit Sub

    SetRect srcRct, 0, 0, srcW, srcH
    SetRect mskRct, x, y, x + mskW, y + mskH
    IntersectRect srcRct, mskRct, srcRct

    srcRct.Right = srcRct.Right
    srcRct.Bottom = srcRct.Bottom

    If (IsRectEmpty(srcRct)) Then Exit Sub

    R = (Color And &HFF&)
    G = (Color And &HFF00&) \ 256
    B = (Color And &HFF0000) \ 65536

    cP = Pressure / 100

    lRegion = AdjustRegion

    For i = srcRct.Left To srcRct.Right
        For j = srcRct.Top To srcRct.Bottom
            If (PtInRegion(lRegion, i, j)) Then
                iIn = i - x
                jIn = j - y
                If (mBits(0, iIn, jIn) < 190) Then
                    cText = cP * (1 - (mBits(0, iIn, jIn) / 190))
                    cBack = 1 - cText
                    tBM.Bits(0, i, j) = cText * B + cBack * tBM.Bits(0, i, j)
                    tBM.Bits(1, i, j) = cText * G + cBack * tBM.Bits(1, i, j)
                    tBM.Bits(2, i, j) = cText * R + cBack * tBM.Bits(2, i, j)
                End If
            End If
        Next j
    Next i

End Sub

Public Sub BitsLine(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Color As Long, ByVal Pressure As Long)

  Dim lRegion As Long
  Dim R As Long, G As Long, B As Long

  Dim savX(1 To 4) As Long, savy(1 To 4) As Long
  Dim savAlpha(1 To 4) As Byte, pixelcount As Long
  Dim bytB As Byte, bytA As Byte
  Dim ax As Single, bx As Single, cx As Single, dx As Single
  Dim ay As Single, by As Single, cy As Single, dy As Single
  Dim l1 As Single, l2 As Single
  Dim l3 As Single, l4 As Single
  Dim RX1 As Long, RY1 As Long
  Dim RX2 As Long, RY2 As Long
  Dim xp5 As Single, yp5 As Single
  Dim r2 As Long, g2 As Long, b2 As Long
  Dim xL As Single, yL As Single
  Dim X4 As Long
  Dim Y4 As Long
  Dim P1 As Long
  Dim s2 As Single
  Dim Sng As Single
  Dim bgr As Long, cl As Long

  Dim Steps As Long
  Dim ddx As Long, ddy As Long
  Dim dix As Long, diy As Long
  Dim mci As Long
  Dim dx_step As Single, dy_step As Single
    
    If (Not BMExists) Then Exit Sub

    R = (Color And &HFF&)
    G = (Color And &HFF00&) \ 256
    B = (Color And &HFF0000) \ 65536

    ddx = x2 - x1
    ddy = y2 - y1
    dix = Abs(ddx)
    diy = Abs(ddy)
    If (dix > diy) Then
        mci = dix
      Else
        mci = diy
    End If

    Steps = mci + 1
    dx_step = ddx / Steps / 2
    dy_step = ddy / Steps / 2

    xL = x1
    yL = y1

    Pressure = Pressure * 2.55
    If (Pressure = 0) Then Pressure = 1

    lRegion = AdjustRegion

    Do Until Sng >= Steps

        Sng = Sng + 0.5

        '// Line
        xL = xL + dx_step
        yL = yL + dy_step

        '// Prevents error when vb rounds .5 down
        If xL = Int(xL) Then xL = xL + 0.001
        If yL = Int(yL) Then yL = yL + 0.001

        '// Anti-alias
        ax = xL - 0.5
        ay = yL - 0.5
        bx = ax + 1
        by = ay + 1
        RX1 = ax
        RX2 = RX1 + 1
        xp5 = RX1 + 0.5
        RY1 = ay
        RY2 = by
        l1 = RY1 + 0.5 - ay
        l2 = 256 * (xp5 - ax) - xp5 + ax
        l3 = 255 - l2
        l4 = by - RY2 + 0.5
        savX(1) = RX1
        savy(1) = RY1
        savX(2) = RX2
        savy(2) = RY1
        savy(3) = RY2
        savX(3) = RX1
        savy(4) = RY2
        savX(4) = RX2
        savAlpha(1) = l1 * l2
        savAlpha(2) = l1 * l3
        savAlpha(3) = l4 * l2
        savAlpha(4) = l4 * l3

        cl = 1
        Do Until cl = 5
            bytA = savAlpha(cl)
            X4 = savX(cl)
            Y4 = savy(cl)
            If (PtInRegion(lRegion, X4, Y4)) Then
                b2 = tBM.Bits(0, X4, Y4)
                g2 = tBM.Bits(1, X4, Y4)
                r2 = tBM.Bits(2, X4, Y4)
                s2 = (bytA / 255) * Pressure / 255
                tBM.Bits(0, X4, Y4) = b2 - s2 * (b2 - B)
                tBM.Bits(1, X4, Y4) = g2 - s2 * (g2 - G)
                tBM.Bits(2, X4, Y4) = r2 - s2 * (r2 - R)
            End If
            cl = cl + 1
        Loop
    Loop

End Sub

Public Sub BitsRectangle(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Color As Long, ByVal Pressure As Long)

    If (BMExists) Then
        BitsLine x1, y1, x2, y1, Color, Pressure
        BitsLine x2, y1, x2, y2, Color, Pressure
        BitsLine x2, y2, x1, y2, Color, Pressure
        BitsLine x1, y2, x1, y1, Color, Pressure
    End If

End Sub

Public Sub BitsPolygon(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Sides As Long, ByVal Color As Long, ByVal Pressure As Long)

  Dim d As Single, dStep As Single, degG As Single
  Dim R As Long
    
    If (Not BMExists) Then Exit Sub

    If (x2 - x1 <> 0) Then
        degG = Atn((y2 - y1) / (x2 - x1))
      Else
        degG = Sgn((y2 - y1)) * 1.5708
    End If

    If (x2 - x1 < 0) Then
        degG = degG + 3.1416
      Else
        If (y2 - y1 < 0) Then
            degG = degG + 6.2832
        End If
    End If

    dStep = 6.28318 / Sides
    R = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)

    For d = degG To degG + 6.28318 - dStep + 0.001 Step dStep
        BitsLine x1 + R * Cos(d), y1 + R * Sin(d), x1 + R * Cos(d + dStep), y1 + R * Sin(d + dStep), Color, Pressure
    Next d

End Sub

Public Sub BitsEllipse(ByVal xC As Long, ByVal yC As Long, ByVal xR As Long, ByVal yR As Long, ByVal Color As Long, ByVal Pressure As Long)
'// Needs readjust on ellipse case
'// ------------------------------

  Dim lRegion As Long
  Dim R As Long, G As Long, B As Long

  Dim savX(1 To 4) As Long, savy(1 To 4) As Long
  Dim savAlpha(1 To 4) As Byte, pixelcount As Long
  Dim bytB As Byte, bytA As Byte
  Dim ax As Single, bx As Single, cx As Single, dx As Single
  Dim ay As Single, by As Single, cy As Single, dy As Single
  Dim l1 As Single, l2 As Single
  Dim l3 As Single, l4 As Single
  Dim RX1 As Long, RY1 As Long
  Dim RX2 As Long, RY2 As Long
  Dim xp5 As Single, yp5 As Single
  Dim r2 As Long, g2 As Long, b2 As Long
  Dim x2 As Single, y2 As Single
  Dim X4 As Long
  Dim Y4 As Long
  Dim P1 As Long
  Dim sngPointSpacing As Single, sngLs As Single, maxR As Long
  Dim s2 As Single
  Dim Sng As Single
  Dim bgr As Long, cl As Long
  Const TwoPI As Single = 6.283185
    
    If (Not BMExists) Then Exit Sub
    
    R = (Color And &HFF&)
    G = (Color And &HFF00&) \ 256
    B = (Color And &HFF0000) \ 65536

    sngLs = TwoPI * 0.085

    Pressure = Pressure * 2.55
    If (Pressure = 0) Then Pressure = 1

    If (yR > xR) Then maxR = yR Else maxR = xR
    If (maxR < 0) Then
        sngPointSpacing = -sngLs / maxR
      ElseIf (maxR = 0) Then
        sngPointSpacing = sngLs
      Else
        sngPointSpacing = sngLs / maxR
    End If

    lRegion = AdjustRegion

    Do Until Sng >= TwoPI

        Sng = Sng + sngPointSpacing

        '// Circle formula -> Ellipse
        x2 = xC + xR * Cos(Sng)
        y2 = yC + yR * Sin(Sng)

        '// Prevents error when vb rounds .5 down
        If x2 = Int(x2) Then x2 = x2 + 0.001
        If y2 = Int(y2) Then y2 = y2 + 0.001

        '// Anti-alias
        ax = x2 - 0.5
        ay = y2 - 0.5
        bx = ax + 1
        by = ay + 1
        RX1 = ax
        RX2 = RX1 + 1
        xp5 = RX1 + 0.5
        RY1 = ay
        RY2 = by
        l1 = RY1 + 0.5 - ay
        l2 = 256 * (xp5 - ax) - xp5 + ax
        l3 = 255 - l2
        l4 = by - RY2 + 0.5
        savX(1) = RX1
        savy(1) = RY1
        savX(2) = RX2
        savy(2) = RY1
        savy(3) = RY2
        savX(3) = RX1
        savy(4) = RY2
        savX(4) = RX2
        savAlpha(1) = l1 * l2
        savAlpha(2) = l1 * l3
        savAlpha(3) = l4 * l2
        savAlpha(4) = l4 * l3

        cl = 1
        Do Until cl = 5
            bytA = savAlpha(cl)
            X4 = savX(cl)
            Y4 = savy(cl)
            If (PtInRegion(lRegion, X4, Y4)) Then
                b2 = tBM.Bits(0, X4, Y4)
                g2 = tBM.Bits(1, X4, Y4)
                r2 = tBM.Bits(2, X4, Y4)
                s2 = (bytA / 255) * Pressure / 255
                tBM.Bits(0, X4, Y4) = b2 - s2 * (b2 - B)
                tBM.Bits(1, X4, Y4) = g2 - s2 * (g2 - G)
                tBM.Bits(2, X4, Y4) = r2 - s2 * (r2 - R)
            End If
            cl = cl + 1
        Loop
    Loop

End Sub

Public Sub BitsBrush(ByVal x As Long, ByVal y As Long, ByVal d As Long, ByVal Pressure As Long, ByVal Color As Long)

  Dim lRegion As Long

  Dim i As Long, j As Long
  Dim iOut As Long, jOut As Long
  Dim bR As Long, bG As Long, Bb As Long
  Dim rBr As Long, rBrpow2 As Long
  Dim cp1 As Single, cp2 As Single
    
    If (Not BMExists) Then Exit Sub

    rBr = 0.5 * d
    rBrpow2 = rBr * rBr + 1

    bR = (Color And &HFF&)
    bG = (Color And &HFF00&) \ 256
    Bb = (Color And &HFF0000) \ 65536

    cp1 = Pressure / 100
    cp2 = 1 - cp1

    lRegion = AdjustRegion

    For j = -rBr To rBr
        jOut = j + y
        For i = -rBr To rBr
            iOut = i + x
            If (i * i + j * j < rBrpow2) Then
                If (PtInRegion(lRegion, iOut, jOut)) Then
                    If (tBMDone(iOut, jOut) = 0) Then
                        tBMDone(iOut, jOut) = -1
                        tBM.Bits(0, iOut, jOut) = cp1 * Bb + cp2 * tBM.Bits(0, iOut, jOut)
                        tBM.Bits(1, iOut, jOut) = cp1 * bG + cp2 * tBM.Bits(1, iOut, jOut)
                        tBM.Bits(2, iOut, jOut) = cp1 * bR + cp2 * tBM.Bits(2, iOut, jOut)
                    End If
                End If
            End If
        Next i
    Next j

End Sub

Public Sub BitsBrushLine(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal d As Long, ByVal Pressure As Long, ByVal Color As Long, ByVal BrushMode As Long)

  Dim i As Long
  Dim xd As Single, yd As Single
  Dim dx As Single, dy As Single
  Dim xInc As Integer, yInc As Integer
  Dim ix As Integer, iy As Integer
  Dim Steps As Integer
    
    If (Not BMExists) Then Exit Sub

    xInc = x2 - x1
    yInc = y2 - y1
    ix = Abs(xInc)
    iy = Abs(yInc)

    If (ix > iy) Then
        Steps = (ix + 1)
      Else
        Steps = (iy + 1)
    End If

    dx = (xInc / Steps) * (d / 10)
    dy = (yInc / Steps) * (d / 10)
    xd = x1
    yd = y1

    For i = 0 To Steps / (d / 10)
        Select Case BrushMode
          Case 0
            BitsBrush xd, yd, d, Pressure, Color
          Case 1
            BitsSoften xd, yd, d, Pressure
          Case 2
            BitsSoften xd, yd, d, Pressure, -1
          Case 3
            BitsSharpen xd, yd, d, Pressure
          Case 4
            BitsBrightness xd, yd, d, Pressure
          Case 5
            BitsBrightness xd, yd, d, Pressure, 0
        End Select
        xd = xd + dx
        yd = yd + dy
    Next i

End Sub

Public Sub BitsAirbrush(ByVal x As Long, ByVal y As Long, ByVal d As Long, ByVal Pressure As Long, ByVal Quantity As Long, ByVal Definition As Long, ByVal Color As Long)

  Dim lRegion As Long

  Dim i As Long, j As Long
  Dim iOut As Long, jOut As Long
  Dim bR As Long, bG As Long, Bb As Long
  Dim rBr As Long, rBrpow2 As Long
  Dim cQ As Single, cD As Single
  Dim cP As Single, cp1 As Double, cp2 As Double
    
    If (Not BMExists) Then Exit Sub

    rBr = 0.5 * d
    rBrpow2 = rBr * rBr + 1

    bR = (Color And &HFF&)
    bG = (Color And &HFF00&) \ 256
    Bb = (Color And &HFF0000) \ 65536

    cP = Pressure / 100
    cQ = Quantity / 100
    cD = 1 - (Definition / 100)

    lRegion = AdjustRegion

    For j = -rBr To rBr
        jOut = j + y
        For i = -rBr To rBr
            iOut = i + x
            If (i * i + j * j < rBrpow2) Then
                If (PtInRegion(lRegion, iOut, jOut) And Rnd <= cQ) Then
                    cp1 = (1 - cD * Sqr(i * i + j * j) / rBr) * cP
                    cp2 = (1 - cp1)
                    tBM.Bits(0, iOut, jOut) = cp1 * Bb + cp2 * tBM.Bits(0, iOut, jOut)
                    tBM.Bits(1, iOut, jOut) = cp1 * bG + cp2 * tBM.Bits(1, iOut, jOut)
                    tBM.Bits(2, iOut, jOut) = cp1 * bR + cp2 * tBM.Bits(2, iOut, jOut)
                End If
            End If
        Next i
    Next j

End Sub

Public Sub BitsBrightness(ByVal x As Long, ByVal y As Long, ByVal d As Long, ByVal Pressure As Long, Optional ByVal LightenMode As Boolean = -1)

  Dim lRegion As Long

  Dim Spd(255) As Long
  Dim i As Long, j As Long
  Dim iOut As Long, jOut As Long
  Dim rBr As Long, rBrpow2 As Long
  Dim cp1 As Single
    
    If (Not BMExists) Then Exit Sub

    rBr = 0.5 * d
    rBrpow2 = rBr * rBr + 1

    If (LightenMode) Then
        cp1 = 1 + Pressure / 100
        For i = 0 To 255
            Spd(i) = cp1 * i
            If Spd(i) > 255 Then Spd(i) = 255
        Next i
      Else
        cp1 = 1 - Pressure / 100
        For i = 0 To 255
            Spd(i) = cp1 * i
        Next i
    End If

    lRegion = AdjustRegion

    For j = -rBr To rBr
        jOut = j + y
        For i = -rBr To rBr
            iOut = i + x
            If (i * i + j * j < rBrpow2) Then
                If (PtInRegion(lRegion, iOut, jOut)) Then
                    If (tBMDone(iOut, jOut) = 0) Then
                        tBMDone(iOut, jOut) = -1
                        tBM.Bits(0, iOut, jOut) = Spd(tBM.Bits(0, iOut, jOut))
                        tBM.Bits(1, iOut, jOut) = Spd(tBM.Bits(1, iOut, jOut))
                        tBM.Bits(2, iOut, jOut) = Spd(tBM.Bits(2, iOut, jOut))
                    End If
                End If
            End If
        Next i
    Next j

End Sub

Public Sub BitsMasked(ByVal x As Long, ByVal y As Long, ByRef mBits() As Byte, ByVal d As Long, ByVal Pressure As Long)

  Dim lRegion As Long

  Dim mskW As Long, mskH As Long
  Dim i As Long, j As Long
  Dim iOut As Long, jOut As Long
  Dim iIn As Long, jIn As Long
  Dim rBr As Long, rBrpow2 As Long
  Dim cp1 As Single, cp2 As Single '
    
    If (Not BMExists) Then Exit Sub

    mskW = UBound(mBits, 2)
    mskH = UBound(mBits, 3)
    If (mskW = 0 Or mskH = 0) Then Exit Sub

    rBr = 0.5 * d
    rBrpow2 = rBr * rBr + 1

    lRegion = AdjustRegion

    cp1 = Pressure / 100
    cp2 = 1 - cp1

    For j = -rBr To rBr
        jOut = j + y
        jIn = jOut Mod mskH
        For i = -rBr To rBr
            iOut = i + x
            iIn = iOut Mod mskW
            If (i * i + j * j < rBrpow2) Then
                If (PtInRegion(lRegion, iOut, jOut)) Then
                    If (tBMDone(iOut, jOut) = 0) Then
                        tBMDone(iOut, jOut) = -1
                        tBM.Bits(0, iOut, jOut) = cp1 * mBits(0, iIn, jIn) + cp2 * tBM.Bits(0, iOut, jOut)
                        tBM.Bits(1, iOut, jOut) = cp1 * mBits(1, iIn, jIn) + cp2 * tBM.Bits(1, iOut, jOut)
                        tBM.Bits(2, iOut, jOut) = cp1 * mBits(2, iIn, jIn) + cp2 * tBM.Bits(2, iOut, jOut)
                    End If
                End If
            End If
        Next i
    Next j

End Sub

Public Sub BitsMaskedLine(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal d As Long, ByVal Pressure As Long, ByRef mBits() As Byte)

  Dim i As Long
  Dim xd As Single, yd As Single
  Dim dx As Single, dy As Single
  Dim xInc As Integer, yInc As Integer
  Dim ix As Integer, iy As Integer
  Dim Steps As Integer
    
    If (Not BMExists) Then Exit Sub

    xInc = x2 - x1
    yInc = y2 - y1
    ix = Abs(xInc)
    iy = Abs(yInc)

    If (ix > iy) Then
        Steps = (ix + 1)
      Else
        Steps = (iy + 1)
    End If

    dx = (xInc / Steps) * (d / 10)
    dy = (yInc / Steps) * (d / 10)
    xd = x1
    yd = y1

    For i = 0 To Steps / (d / 10)
        BitsMasked xd, yd, mBits, d, Pressure
        xd = xd + dx
        yd = yd + dy
    Next i

End Sub

Public Sub BitsSoften(ByVal x As Long, ByVal y As Long, ByVal d As Long, ByVal Pressure As Long, Optional ByVal Overlay As Boolean = 0)

  Dim lRegion As Long

  Dim i As Long, j As Long
  Dim iOut As Long, jOut As Long
  Dim rBr As Long, rBrpow2 As Long
  Dim cP As Long, cPW As Long
    
    If (Not BMExists) Then Exit Sub

    rBr = 0.5 * d
    rBrpow2 = rBr * rBr + 1

    cP = 10 - Pressure / 10
    cPW = cP + 4

    lRegion = AdjustRegion

    For j = -rBr To rBr
        jOut = j + y
        For i = -rBr To rBr
            iOut = i + x
            If (i * i + j * j < rBrpow2) Then
                If (PtInRegion(lRegion, iOut, jOut)) Then
                    If (tBMDone(iOut, jOut) = 0 Or Overlay) Then
                        tBMDone(iOut, jOut) = -1
                        tBM.Bits(0, iOut, jOut) = (CLng(tBM.Bits(0, iOut, jOut)) * cP + _
                                 tBM.Bits(0, iOut - 1, jOut) + tBM.Bits(0, iOut, jOut - 1) + _
                                 tBM.Bits(0, iOut, jOut + 1) + tBM.Bits(0, iOut + 1, jOut)) \ cPW
                        tBM.Bits(1, iOut, jOut) = (CLng(tBM.Bits(1, iOut, jOut)) * cP + _
                                 tBM.Bits(1, iOut - 1, jOut) + tBM.Bits(1, iOut, jOut - 1) + _
                                 tBM.Bits(1, iOut, jOut + 1) + tBM.Bits(1, iOut + 1, jOut)) \ cPW
                        tBM.Bits(2, iOut, jOut) = (CLng(tBM.Bits(2, iOut, jOut)) * cP + _
                                 tBM.Bits(2, iOut - 1, jOut) + tBM.Bits(2, iOut, jOut - 1) + _
                                 tBM.Bits(2, iOut, jOut + 1) + tBM.Bits(2, iOut + 1, jOut)) \ cPW
                    End If
                End If
            End If
        Next i
    Next j

End Sub

Public Sub BitsSharpen(ByVal x As Long, ByVal y As Long, ByVal d As Long, ByVal Pressure As Long)

  Dim lRegion As Long

  Dim i As Long, j As Long
  Dim B As Long, G As Long, R As Long
  Dim iOut As Long, jOut As Long
  Dim rBr As Long, rBrpow2 As Long
  Dim cP As Long, cPW As Long

    If (Not BMExists) Then Exit Sub
    
    rBr = 0.5 * d
    rBrpow2 = rBr * rBr + 1

    cP = 15 - Pressure / 15
    cPW = cP - 4

    lRegion = AdjustRegion

    For j = -rBr To rBr
        jOut = j + y
        For i = -rBr To rBr
            iOut = i + x
            If (i * i + j * j < rBrpow2) Then
                If (PtInRegion(lRegion, iOut, jOut)) Then
                    If (tBMDone(iOut, jOut) = 0) Then
                        tBMDone(iOut, jOut) = -1
                        B = (CLng(tBM.Bits(0, iOut, jOut)) * cP - _
                            tBM.Bits(0, iOut - 1, jOut) - tBM.Bits(0, iOut, jOut - 1) - _
                            tBM.Bits(0, iOut, jOut + 1) - tBM.Bits(0, iOut + 1, jOut)) \ cPW
                        G = (CLng(tBM.Bits(1, iOut, jOut)) * cP - _
                            tBM.Bits(1, iOut - 1, jOut) - tBM.Bits(1, iOut, jOut - 1) - _
                            tBM.Bits(1, iOut, jOut + 1) - tBM.Bits(1, iOut + 1, jOut)) \ cPW
                        R = (CLng(tBM.Bits(2, iOut, jOut)) * cP - _
                            tBM.Bits(2, iOut - 1, jOut) - tBM.Bits(2, iOut, jOut - 1) - _
                            tBM.Bits(2, iOut, jOut + 1) - tBM.Bits(2, iOut + 1, jOut)) \ cPW
                        If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                        If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                        If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                        tBM.Bits(0, iOut, jOut) = B
                        tBM.Bits(1, iOut, jOut) = G
                        tBM.Bits(2, iOut, jOut) = R
                    End If
                End If
            End If
        Next i
    Next j

End Sub

Public Sub BitsFill(ByVal xFill As Long, ByVal yFill As Long, ByVal Color As Long, ByVal Tolerance As Long, ByVal Pressure As Long)

  Dim lRegion As Long

  Dim spdR(255) As Long
  Dim spdG(255) As Long
  Dim spdB(255) As Long

  Dim Rt As Long, Gt As Long, Bt As Long
  Dim Rf As Long, Gf As Long, Bf As Long
  Dim xIn As Long, yIn As Long

  Dim BasePoint As POINTAPI
  Dim KeepChecking As Boolean
  Dim PointList() As POINTAPI
  Dim NumPoints As Long
  Dim DoPoint As Long
  Dim CheckPoints As Long
  Dim OffPix As POINTAPI
  Dim CheckRound As Long

  Dim distance As Long
  Dim tF As Long
  Dim cp1 As Single, cp2 As Single
    
    If (Not BMExists) Then Exit Sub

    If (PtInRect(BMRect, xFill, yFill) = 0) Then Exit Sub

    Rt = (Color And &HFF&)
    Gt = (Color And &HFF00&) \ 256
    Bt = (Color And &HFF0000) \ 65536
    Rf = tBM.Bits(2, xFill, yFill)
    Gf = tBM.Bits(1, xFill, yFill)
    Bf = tBM.Bits(0, xFill, yFill)

    tF = Tolerance
    cp1 = 1 - (Pressure / 100)
    cp2 = 1 - cp1

    lRegion = AdjustRegion

    If (PtInRegion(lRegion, xFill, yFill)) Then
        tBM.Bits(0, xFill, yFill) = cp1 * Bf + cp2 * Bt
        tBM.Bits(1, xFill, yFill) = cp1 * Gf + cp2 * Gt
        tBM.Bits(2, xFill, yFill) = cp1 * Rf + cp2 * Rt
    End If

    If (Abs(Rt - Rf) < 10 And Abs(Gt - Gf) < 10 And Abs(Bt - Bf) < 10) Then Exit Sub

    Rt = cp2 * Rt
    Gt = cp2 * Gt
    Bt = cp2 * Bt

    For xIn = 0 To 255
        spdR(xIn) = cp1 * xIn + Rt
        spdG(xIn) = cp1 * xIn + Gt
        spdB(xIn) = cp1 * xIn + Bt
    Next xIn

    BasePoint.x = xFill
    BasePoint.y = yFill

    ReDim PointList(0) As POINTAPI
    PointList(0) = BasePoint

    NumPoints = 0
    DoPoint = 0

    Do
        KeepChecking = 0

        For CheckPoints = DoPoint To NumPoints
            For CheckRound = 0 To 3
                Select Case CheckRound
                  Case 0
                    OffPix.x = -1
                    OffPix.y = 0
                  Case 1
                    OffPix.x = 1
                    OffPix.y = 0
                  Case 2
                    OffPix.x = 0
                    OffPix.y = -1
                  Case 3
                    OffPix.x = 0
                    OffPix.y = 1
                End Select

                xIn = PointList(CheckPoints).x + OffPix.x
                yIn = PointList(CheckPoints).y + OffPix.y

                If (PtInRegion(lRegion, xIn, yIn)) Then
                    If (tBMDone(xIn, yIn) = 0) Then
                        tBMDone(xIn, yIn) = -1

                        distance = Abs(tBM.Bits(0, xIn, yIn) - Bf) _
                                 + Abs(tBM.Bits(1, xIn, yIn) - Gf) _
                                 + Abs(tBM.Bits(2, xIn, yIn) - Rf)

                        If (distance <= tF) Then
                            tBM.Bits(0, xIn, yIn) = spdB(tBM.Bits(0, xIn, yIn))
                            tBM.Bits(1, xIn, yIn) = spdG(tBM.Bits(1, xIn, yIn))
                            tBM.Bits(2, xIn, yIn) = spdR(tBM.Bits(2, xIn, yIn))

                            NumPoints = NumPoints + 1
                            ReDim Preserve PointList(NumPoints)

                            PointList(NumPoints).x = xIn
                            PointList(NumPoints).y = yIn
                            KeepChecking = -1
                        End If
                    End If
                End If
            Next CheckRound
        Next CheckPoints

        DoPoint = CheckPoints
    Loop Until Not KeepChecking

End Sub

Public Sub BitsFillTexture(ByVal xFill As Long, ByVal yFill As Long, ByRef mBits() As Byte, ByVal Tolerance As Long, ByVal Pressure As Long)

  Dim lRegion As Long

  Dim spdR(255) As Long
  Dim spdG(255) As Long
  Dim spdB(255) As Long

  Dim spdM() As Byte

  Dim Rf As Long, Gf As Long, Bf As Long
  Dim xIn As Long, yIn As Long

  Dim BasePoint As POINTAPI
  Dim KeepChecking As Boolean
  Dim PointList() As POINTAPI
  Dim NumPoints As Long
  Dim DoPoint As Long
  Dim CheckPoints As Long
  Dim OffPix As POINTAPI
  Dim CheckRound As Long

  Dim distance As Long
  Dim tF As Long
  Dim cp1 As Single, cp2 As Single

  Dim mx As Long, my As Long
  Dim mw As Long, mh As Long
  Dim mxIn As Long, myIn As Long
    
    If (Not BMExists) Then Exit Sub

    If (PtInRect(BMRect, xFill, yFill) = 0) Then Exit Sub

    mw = UBound(mBits, 2) + 1
    mh = UBound(mBits, 3) + 1
    If (mw = 0 Or mh = 0) Then Exit Sub

    Rf = tBM.Bits(2, xFill, yFill)
    Gf = tBM.Bits(1, xFill, yFill)
    Bf = tBM.Bits(0, xFill, yFill)

    tF = Tolerance
    cp1 = 1 - (Pressure / 100)
    cp2 = 1 - cp1

    lRegion = AdjustRegion

    If (PtInRegion(lRegion, xFill, yFill)) Then
        tBM.Bits(0, xFill, yFill) = cp1 * Bf + cp2 * mBits(0, xFill Mod mw, yFill Mod mh)
        tBM.Bits(1, xFill, yFill) = cp1 * Gf + cp2 * mBits(1, xFill Mod mw, yFill Mod mh)
        tBM.Bits(2, xFill, yFill) = cp1 * Rf + cp2 * mBits(2, xFill Mod mw, yFill Mod mh)
    End If
    If (Abs(CLng(mBits(0, xFill Mod mw, yFill Mod mh)) - Bf) < 10 And _
        Abs(CLng(mBits(1, xFill Mod mw, yFill Mod mh)) - Gf) < 10 And _
        Abs(CLng(mBits(2, xFill Mod mw, yFill Mod mh)) - Rf) < 10) Then
        Exit Sub
    End If

    For xIn = 0 To 255
        spdR(xIn) = cp1 * xIn
        spdG(xIn) = cp1 * xIn
        spdB(xIn) = cp1 * xIn
    Next xIn

    spdM = mBits
    For xIn = 0 To mw - 1
        For yIn = 0 To mh - 1
            spdM(0, xIn, yIn) = cp2 * spdM(0, xIn, yIn)
            spdM(1, xIn, yIn) = cp2 * spdM(1, xIn, yIn)
            spdM(2, xIn, yIn) = cp2 * spdM(2, xIn, yIn)
        Next yIn
    Next xIn

    BasePoint.x = xFill
    BasePoint.y = yFill

    ReDim PointList(0) As POINTAPI
    PointList(0) = BasePoint

    NumPoints = 0
    DoPoint = 0

    Do
        KeepChecking = 0

        For CheckPoints = DoPoint To NumPoints
            For CheckRound = 0 To 3
                Select Case CheckRound
                  Case 0
                    OffPix.x = -1
                    OffPix.y = 0
                  Case 1
                    OffPix.x = 1
                    OffPix.y = 0
                  Case 2
                    OffPix.x = 0
                    OffPix.y = -1
                  Case 3
                    OffPix.x = 0
                    OffPix.y = 1
                End Select

                xIn = PointList(CheckPoints).x + OffPix.x
                yIn = PointList(CheckPoints).y + OffPix.y

                If (PtInRegion(lRegion, xIn, yIn)) Then
                    If (tBMDone(xIn, yIn) = 0) Then
                        tBMDone(xIn, yIn) = -1

                        distance = Abs(tBM.Bits(0, xIn, yIn) - Bf) _
                                   + Abs(tBM.Bits(1, xIn, yIn) - Gf) _
                                   + Abs(tBM.Bits(2, xIn, yIn) - Rf)

                        If (distance <= tF) Then
                            mxIn = xIn Mod mw
                            myIn = yIn Mod mh
                            tBM.Bits(0, xIn, yIn) = spdB(tBM.Bits(0, xIn, yIn)) + spdM(0, mxIn, myIn)
                            tBM.Bits(1, xIn, yIn) = spdG(tBM.Bits(1, xIn, yIn)) + spdM(1, mxIn, myIn)
                            tBM.Bits(2, xIn, yIn) = spdR(tBM.Bits(2, xIn, yIn)) + spdM(2, mxIn, myIn)

                            NumPoints = NumPoints + 1
                            ReDim Preserve PointList(NumPoints)

                            PointList(NumPoints).x = xIn
                            PointList(NumPoints).y = yIn
                            KeepChecking = -1
                        End If
                    End If
                End If
            Next CheckRound
        Next CheckPoints

        DoPoint = CheckPoints
    Loop Until Not KeepChecking

End Sub

Private Function AdjustRegion() As Long

    If (BitsRegion.IsEmptyRegion) Then
        AdjustRegion = BitsRegion.RegionMain
      Else
        AdjustRegion = BitsRegion.Region
    End If

End Function

Public Sub CreateDoneBitsTable()

    ReDim tBMDone(iSrc.Width + 2, iSrc.Height + 2)

End Sub

'// Clipboard
'========================================================================================

Public Sub BitsCopyToClipboard()

  Dim lhDC As Long
  Dim lhBmpOld As Long
  Dim lhObj As Long
    
    If (Not BMExists) Then Exit Sub

    lhDC = CreateCompatibleDC(iSrc.hdc)
    If (lhDC <> 0) Then
        lhObj = CreateCompatibleBitmap(iSrc.hdc, iSrc.Width, iSrc.Height)
        If (lhObj <> 0) Then

            lhBmpOld = SelectObject(lhDC, lhObj)
            SetDIBitsToDevice lhDC, 0, 0, iSrc.Width, iSrc.Height, 2, 2, 0, iSrc.Height + 2, tBM.Bits(0, -2, -2), tBM, DIB_RGB_COLORS
            SelectObject lhDC, lhBmpOld

            If (OpenClipboard(0) <> 0) Then
                EmptyClipboard
                SetClipboardData CF_BITMAP, lhObj
                CloseClipboard
            End If
        End If
        DeleteDC lhDC
    End If

End Sub

Public Sub PasteFromClipboard()

    If (Clipboard.GetFormat(vbCFBitmap)) Then
        Set Picture = Clipboard.GetData
    End If

End Sub

Public Sub Save(ByVal FileName As String)

    FileName = Left(FileName, InStrRev(FileName, ".")) & "bmp"

    If (Not BMExists) Then Exit Sub
    
    On Error Resume Next
      With iSrc
          Set .Picture = Nothing
          .AutoRedraw = -1
          StretchDIBits .hdc, 0, 0, .Width, .Height, 2, 2, .Width, .Height, tBM.Bits(0, -2, -2), tBM, DIB_RGB_COLORS, SRCCOPY
          SavePicture .Image, FileName
          .AutoRedraw = 0
      End With
    On Error GoTo 0

End Sub

'// Properties
'========================================================================================

Public Property Get Bits() As Byte()

  '// Get BITs table

    If (BMExists) Then
        Bits = tBM.Bits
    End If

End Property

Public Property Let Bits(ByRef srcBits() As Byte)

  '// Set BITs table

    If (BMExists) Then
        tBM.Bits = srcBits
    End If

End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal nData As OLE_COLOR)

    UserControl.BackColor = nData
    iDst.BackColor = nData
    iSrc.BackColor = nData
    UpdateImage

End Property

Public Property Get BarsWidth() As Integer

    BarsWidth = pBarsWidth

End Property

Public Property Let BarsWidth(ByVal nData As Integer)

    If (nData > 13) Then
        pBarsWidth = nData
      Else
        pBarsWidth = 13
    End If
    UserControl_Resize

End Property

Public Property Get MouseScroll() As Boolean

    MouseScroll = pMouseScroll

End Property

Public Property Let MouseScroll(ByVal nData As Boolean)

    If (nData) And (hSB.Max > 0 Or vSB.Max > 0) Then
        iDst.MousePointer = vbCustom
        iDst.MouseIcon = cRelease
      Else
        iDst.MousePointer = vbDefault
    End If
    pMouseScroll = nData

End Property

Public Property Set Picture(ByVal nData As StdPicture)

  Dim preBM As BITMAPINFO
  Dim ptrMem As Long, lenMem As Long

    hSB.Visible = 0
    vSB.Visible = 0

    iSrc = nData
    If (iSrc = 0) Then Clear: Exit Property

    With iSrc
        CreateBitmap preBM, preBM.Bits, .Width, .Height
        GetDIBits .hdc, .Image.Handle, 0, .Height, preBM.Bits(0, 0, 0), preBM, DIB_RGB_COLORS

        ReDim tBM.Bits(3, -2 To (.Width - 1) + 2, -2 To (.Height - 1) + 2)
        With tBM.Header
            .biSize = 40
            .biBitCount = 32
            .biPlanes = 1
            .biWidth = iSrc.Width + 4
            .biHeight = -iSrc.Height - 4
        End With

        lenMem = 4 * .Width
        For ptrMem = 0 To .Height - 1
            CopyMemory tBM.Bits(0, 0, ptrMem), preBM.Bits(0, 0, ptrMem), lenMem
        Next ptrMem

        Set .Picture = Nothing
        Erase preBM.Bits

        SetRect BMRect, 0, 0, .Width, .Height
        BitsRegion.Init 0, 0, .Width, .Height
    End With

    BMExists = -1
    UserControl_Resize

End Property

Public Property Get PictureExists() As Boolean

    PictureExists = BMExists

End Property

Public Property Get PictureWidth() As Long

    If (BMExists) Then
        PictureWidth = iSrc.Width
      Else
        PictureWidth = 0
    End If

End Property

Public Property Get PictureHeight() As Long

    If (BMExists) Then
        PictureHeight = iSrc.Height
      Else
        PictureHeight = 0
    End If

End Property

Public Property Get Bars() As BarsCts

    Bars = pBars

End Property

Public Property Let Bars(ByVal nData As BarsCts)

    pBars = nData
    UserControl_Resize

End Property

Public Property Get ZoomFactor() As Integer

    ZoomFactor = iz

End Property

Public Property Let ZoomFactor(ByVal nData As Integer)

    If (nData < 1) Then
        nData = 1
      ElseIf (nData > UBound(Zm)) Then
        nData = UBound(Zm)
    End If
    iz = nData

    BitsRegion.ZoomFactor = iz
    UserControl_Resize

End Property

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 12:39:54) 55 + 1929 = 1984 Lines
