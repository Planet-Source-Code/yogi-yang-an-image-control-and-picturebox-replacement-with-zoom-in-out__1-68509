VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//
'// cBitFilter Class
'// Carles P.V. - 2002
'//
'// Special (all!) thanks to Manuel Augusto Santos
'//
'// -------------------------------------------------
'//
'// Comments:
'//
'// -Filters work on 3D array (RGB+P, Width, Height)
'// -2 border pixel added for array filters (max.5x5)
'//  and scroll 'manipulation'.
'//
'// -Ex: Source image: 100x100 pixels
'//      Array dims. : (0 To 3, -2 To 101, -2 To 101)
'//
'// -------------------------------------------------

'//  ---------------------------------------------------------------------------------------------
'//  Blend                 <RadiusFactor, [CleanBack]>
'//  Brightness            <Factor>
'//  Colorize              <Hue, [Saturation], [Luminosity]>
'//  Contour               <Factor>
'//  Contrast              <Factor>
'//  Diffuse               <Factor>
'//  Dissolve 'Aqua'       -
'//  Emboss                <Factor>
'//  Engrave               <Factor>
'//  Equalize              -
'//  Flip                  <[fH], [fV]>
'//  Fold                  <wFactor, hFactor, [CleanBack]>
'//  GammaCorrection       <Factor>
'//  GradientCircular      <x1, y1, x2, y2, Color1, Color2, Pressure, Frequency>
'//  GradientLinear        <x1, y1, x2, y2, Color1, Color2, Pressure, Frequency>
'//  GreyScale             -
'//  Grid                  <HStep, VStep, ForceMargins, Color, [Opacity]>
'//  Halftone              <Radius, Definition, Intensity>
'//  Invert                -
'//  Isolate               <Color, Tolerance>
'//  KillChannel           <rChannel, gChannel, bChannel>
'//  Lens                  <xRel, yRel, dRel, Factor>
'//  Light                 <xRel, yRel, dRel, iFactor, dFactor>
'//  MaskFilter            <mskBits() 2D, Weight, [rFactor], [gFactor], [bFactor]>
'//  Noise                 <Factor, [Equalized]>
'//  PaletteApply          -
'//  Pixelize              <PixSize>
'//  RankFilterN4          <fltType>
'//  RankFilterND          <fltType>
'//  Reduction             <Factor>
'//  Replace               <ColorFrom, ColorTo, Tolerance>
'//  ReplaceHS             <H, S, ColorTo, HTolerance, STolerance>
'//  ReplaceL              <L, ColorTo, Tolerance>
'//  Rotate                <Degrees, [CleanBack]>
'//  RotateH               <HOffset, Saturation, HStep>
'//  RotateL               <LOffset, LStep>
'//  RotateSmooth          <Degrees>
'//  Scanlines             <Factor>
'//  Sharpen               <Factor>
'//  ShiftRGB              <rFactor, gFactor, bFactor>
'//  Shift                 <ColorFrom, Inc, Tolerance>
'//  Sketch                <Factor, [BW]>
'//  Soften                <Factor>
'//  Stretch               <wFactor, hFactor, [CleanBack]>
'//  StuckiBW              <PalWeight>
'//  SwapRGB               <rBank, gBank, bBank>
'//  Texturize             <mskBits()>
'//  Threshold             <Factor>
'//  TileMask              <mskBits(), [xOffset], [yOffset], [Tiled], [Transparency], [MaskColor]
'//  Tile                  <Factor>
'//  Twirl                 <Factor, [CleanBack]>
'//  Waves                 <fFactor, AFactor, Start>
'//  ---------------------------------------------------------------------------------------------
'//  BitsBlack             -
'//  BitsCreateExtBitmap   <BM 'BITMAPINFO', srcBits(), srcWidth, srcHeight>
'//  BitsCrop              <srcBits(), x1, y1, x2, y2>
'//  BitsResample          <srcBits(), newWidth, newHeight> (interp.)
'//  BitsResize            <srcBits(), newWidth, newHeight> (no interp.)
'//  BitsRotate90          <srcBits()>
'//  BitsTakeFromPicture   <srcPicture, srcWidth, srcHeight>
'//  ---------------------------------------------------------------------------------------------

Option Explicit

Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long

Private Const PROGRESS_DELAY As Long = 10
Private Const PI As Single = 3.141592
Private Const PIdiv180 As Single = PI / 180

Private mBits() As Byte
Private biW As Long
Private biH As Long
Private biRct As RECT
Private biRgn As Long

Private x As Long, xIn As Long, xOut As Long
Private y As Long, yIn As Long, yOut As Long
Private xStart As Long, xEnd As Long
Private yStart As Long, yEnd As Long
Private t As Long

Public Event Progress(P As Long)

'//
'// Properties:
'//

Public Property Get BitsWidth() As Long

  '// ~ScaleWidth

    BitsWidth = biW + 1

End Property

Public Property Get BitsHeight() As Long

  '// ~ScaleHeight

    BitsHeight = biH + 1

End Property

Public Property Get Bits() As Byte()

  '// Pass bits Array

    Bits = mBits

End Property

Public Property Let Bits(ByRef srcBits() As Byte)

  '// Get bits array

    mBits = srcBits
    biW = UBound(mBits, 2) - 2
    biH = UBound(mBits, 3) - 2
    SetRect biRct, 0, 0, biW + 1, biH + 1

End Property

Public Property Let BitsRegion(ByVal hRgn As Long)

  '// Effect region

    biRgn = hRgn

End Property

Private Sub Class_Terminate()

    Erase mBits

End Sub

'//
'// Subs: [no returned array]
'//

Public Sub BitsBlack()

  Dim blackBit(3) As Byte

    GetRgnBoundRect

    '// Paint black bits ('cut' bits) ... improve this
    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                CopyMemory mBits(0, x, y), blackBit(0), 4
            End If
        Next x
    Next y

End Sub

Friend Sub BitsCreateExtBitmap(ByRef BI As BITMAPINFO, ByRef srcBits() As Byte, ByVal srcWidth As Long, ByVal srcHeight As Long)

    ReDim BI.Bits(3, -2 To (srcWidth - 1) + 2, -2 To (srcHeight - 1) + 2) As Byte

    '// Set bitmap info
    With BI.Header
        .biSize = 40
        .biBitCount = 32
        .biPlanes = 1
        .biWidth = srcWidth + 4
        .biHeight = -srcHeight - 4
    End With
    BI.Bits = srcBits

End Sub

Public Sub BitsCrop(ByRef srcBits() As Byte, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

  Dim CropRect As RECT
  Dim lenM As Long

    '// Crop 1st. with main
    SetRect CropRect, x1, y1, x2, y2
    IntersectRect CropRect, biRct, CropRect

    '// Get margins (crop rect)
    biW = CropRect.Right - CropRect.Left - 1
    biH = CropRect.Bottom - CropRect.Top - 1
    SetRect biRct, 0, 0, biW + 1, biH + 1

    '// Redim. bitmap and move bits
    ReDim mBits(3, -2 To biW + 2, -2 To biH + 2) As Byte
    lenM = 4 * (biW + 1)
    For y = 0 To biH
        CopyMemory mBits(0, 0, y), srcBits(0, CropRect.Left, CropRect.Top + y), lenM
    Next y

    RaiseEvent Progress(100)

End Sub

Public Sub BitsCropMasked(ByRef srcBits() As Byte, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

  Dim CropRect As RECT

    '// Crop 1st. with main
    SetRect biRct, 0, 0, UBound(srcBits, 2) - 1, UBound(srcBits, 3) - 1
    SetRect CropRect, x1, y1, x2, y2
    IntersectRect CropRect, biRct, CropRect

    '// Get margins (crop rect)
    biW = CropRect.Right - CropRect.Left - 1
    biH = CropRect.Bottom - CropRect.Top - 1
    SetRect biRct, 0, 0, biW + 1, biH + 1

    '// Redim. bitmap and move bits
    ReDim mBits(3, -2 To biW + 2, -2 To biH + 2) As Byte
    For y = 0 To biH
        yIn = CropRect.Top + y
        For x = 0 To biW
            xIn = CropRect.Left + x
            If (PtInRegion(biRgn, xIn, yIn)) Then
                CopyMemory mBits(0, x, y), srcBits(0, xIn, yIn), 4
              Else
                mBits(3, x, y) = &HFF '// The first use reserved byte...
            End If
        Next x
    Next y

    RaiseEvent Progress(100)

End Sub

Public Sub BitsResample(ByRef srcBits() As Byte, ByVal newWidth As Long, ByVal newHeight As Long)

  Dim cx As Single, cy As Single
  Dim dx As Single, dy As Single
  Dim cdx As Single, cdy As Single
  Dim idx As Long, idy As Long

  Dim R As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single
  Dim G As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single
  Dim B As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single

  Dim ir1 As Long, ig1 As Long, ib1 As Long
  Dim ir2 As Long, ig2 As Long, ib2 As Long

    ReDim mBits(3, -2 To (newWidth - 1) + 2, -2 To (newHeight - 1) + 2) As Byte

    cx = (UBound(srcBits, 2) - 2) / newWidth
    cy = (UBound(srcBits, 3) - 2) / newHeight
    biW = newWidth - 1
    biH = newHeight - 1
    SetRect biRct, 0, 0, biW + 1, biH + 1

    For y = 0 To biH
        cdy = y * cy
        idy = Int(cdy)
        dy = cdy - idy
        For x = 0 To biW
            cdx = x * cx
            idx = Int(cdx)
            dx = cdx - idx
            '// Interpolate using the four nearest pixels in the source
            b1 = srcBits(0, idx, idy)
            g1 = srcBits(1, idx, idy)
            r1 = srcBits(2, idx, idy)
            b2 = srcBits(0, idx + 1, idy)
            g2 = srcBits(1, idx + 1, idy)
            r2 = srcBits(2, idx + 1, idy)
            b3 = srcBits(0, idx, idy + 1)
            g3 = srcBits(1, idx, idy + 1)
            r3 = srcBits(2, idx, idy + 1)
            b4 = srcBits(0, idx + 1, idy + 1)
            g4 = srcBits(1, idx + 1, idy + 1)
            r4 = srcBits(2, idx + 1, idy + 1)
            '// Interpolate in y direction:
            ib1 = b1 * (1 - dy) + b3 * dy
            ig1 = g1 * (1 - dy) + g3 * dy
            ir1 = r1 * (1 - dy) + r3 * dy
            ib2 = b2 * (1 - dy) + b4 * dy
            ig2 = g2 * (1 - dy) + g4 * dy
            ir2 = r2 * (1 - dy) + r4 * dy
            '// Interpolate in x direction:
            B = ib1 * (1 - dx) + ib2 * dx
            G = ig1 * (1 - dx) + ig2 * dx
            R = ir1 * (1 - dx) + ir2 * dx
            '// Check ranges
            If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
            If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
            If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
            '// Set output:
            mBits(0, x, y) = B
            mBits(1, x, y) = G
            mBits(2, x, y) = R
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y

    RaiseEvent Progress(100)

End Sub

Public Sub BitsResize(ByRef srcBits() As Byte, ByVal newWidth As Long, ByVal newHeight As Long)

  Dim cx As Single, cy As Single
  Dim idx As Long, idy As Long

    ReDim mBits(3, -2 To (newWidth - 1) + 2, -2 To (newHeight - 1) + 2) As Byte

    cx = (UBound(srcBits, 2) - 1) / newWidth
    cy = (UBound(srcBits, 3) - 1) / newHeight
    biW = newWidth - 1
    biH = newHeight - 1
    SetRect biRct, 0, 0, biW + 1, biH + 1

    For y = 0 To biH
        idy = Int(y * cy)
        For x = 0 To biW
            idx = Int(x * cx)
            mBits(0, x, y) = srcBits(0, idx, idy)
            mBits(1, x, y) = srcBits(1, idx, idy)
            mBits(2, x, y) = srcBits(2, idx, idy)
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y

    RaiseEvent Progress(100)

End Sub

Public Sub BitsRotate90(ByRef srcBits() As Byte)

  Dim spW As Long
  
    ReDim mBits(3, -2 To biH + 2, -2 To biW + 2) As Byte

    For y = 0 To biH
        For x = 0 To biW
            mBits(0, y, biW - x) = srcBits(0, x, y)
            mBits(1, y, biW - x) = srcBits(1, x, y)
            mBits(2, y, biW - x) = srcBits(2, x, y)
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y

    spW = biW
    biW = biH
    biH = spW
    SetRect biRct, 0, 0, biW + 1, biH + 1

    RaiseEvent Progress(100)

End Sub

Public Sub BitsTakeFromPicture(ByRef srcPicture As StdPicture, ByVal srcWidth As Long, ByVal srcHeight As Long)

  Dim lhDC As Long
  Dim lhBmpOld As Long
  Dim tBM As BITMAPINFO

  Dim preBM As BITMAPINFO
  Dim ptrMem As Long, lenMem As Long

    lhDC = CreateCompatibleDC(0)
    If (lhDC <> 0) Then
        lhBmpOld = SelectObject(lhDC, srcPicture.Handle)
        SelectObject lhDC, lhBmpOld

        CreateBitmap preBM, preBM.Bits, srcWidth, srcHeight
        GetDIBits lhDC, srcPicture.Handle, 0, srcHeight, preBM.Bits(0, 0, 0), preBM, DIB_RGB_COLORS
        DeleteObject lhDC

        With tBM.Header
            .biSize = 40
            .biWidth = srcWidth + 4
            .biHeight = -srcHeight - 4
            .biPlanes = 1
            .biBitCount = 32
            ReDim tBM.Bits(3, -2 To (srcWidth - 1) + 2, -2 To (srcHeight - 1) + 2)
        End With

        lenMem = 4 * srcWidth
        For ptrMem = 0 To srcHeight - 1
            CopyMemory tBM.Bits(0, 0, ptrMem), preBM.Bits(0, 0, ptrMem), lenMem
        Next ptrMem

        Bits = tBM.Bits
        Erase tBM.Bits
        Erase preBM.Bits
    End If

End Sub

'//
'// Functions: [returned array]
'//

Public Function Bend(ByVal RadiusFactor As Long, Optional ByVal CleanBack As Boolean = -1) As Byte()

  Dim tBits() As Byte
  Dim ox As Long, oy As Long
  Dim zx As Long, zy As Long
  Dim zAngle As Single, zR As Single
  
  Dim cR As Single
  
  Const TwoPI As Single = 2 * PI

    tBits = mBits

    ox = biW \ 2
    oy = biH \ 2
    
    cR = 2 + (RadiusFactor / 100)

    GetRgnBoundRect

    For y = yStart To yEnd
        zy = y - oy
        For x = xStart To xEnd
            zx = x - ox
            zR = cR * Sqr(zx * zx + zy * zy)
            zAngle = -ArcTan(zy, zx)
            If (zy < 0 Or zx > 0) Then zAngle = zAngle + TwoPI
            xIn = biW * zAngle / TwoPI
            yIn = biH - zR
            If (PtInRegion(biRgn, xIn, yIn)) Then
                tBits(0, x, y) = mBits(0, xIn, yIn)
                tBits(1, x, y) = mBits(1, xIn, yIn)
                tBits(2, x, y) = mBits(2, xIn, yIn)
              Else
                If (CleanBack) Then
                    tBits(0, x, y) = 0
                    tBits(1, x, y) = 0
                    tBits(2, x, y) = 0
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Bend = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Brightness(ByVal Factor As Long) As Byte()
  '             = ShiftRGB(RFactor=GFactor=BFactor)

  Dim tBits() As Byte
  Dim Spd(255) As Long

    tBits = mBits

    For x = 0 To 255
        Spd(x) = x + Factor
        If (Spd(x) > 255) Then Spd(x) = 255
        If (Spd(x) < 0) Then Spd(x) = 0
    Next x

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                tBits(0, x, y) = Spd(mBits(0, x, y))
                tBits(1, x, y) = Spd(mBits(1, x, y))
                tBits(2, x, y) = Spd(mBits(2, x, y))
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Brightness = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Colorize(ByVal Hue As Single, Optional ByVal Saturation As Single = 0.5, Optional ByVal Luminosity As Single = 1) As Byte()

  '// Need mHSLRGB module

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim H As Single, S As Single, L As Single

    tBits = mBits

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                L = (CLng(mBits(0, x, y)) + mBits(1, x, y) + mBits(2, x, y)) / 765
                HSLtoRGB Hue, Saturation, L * Luminosity, R, G, B
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Colorize = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Contour(ByVal Factor As Long, Optional ByVal BackColor As Long = &HFFFFFF) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim rB As Long, Gb As Long, Bb As Long
  Dim cF As Long

    tBits = Bits

    rB = (BackColor And &HFF&)
    Gb = (BackColor And &HFF00&) \ 256
    Bb = (BackColor And &HFF0000) \ 65536
    cF = Factor

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                B = CLng(mBits(0, x - 1, y - 1)) + mBits(0, x - 1, y) + _
                         mBits(0, x - 1, y + 1) + mBits(0, x, y - 1) + _
                         mBits(0, x, y + 1) + mBits(0, x + 1, y - 1) + _
                         mBits(0, x + 1, y) + mBits(0, x + 1, y + 1)
                G = CLng(mBits(1, x - 1, y - 1)) + mBits(1, x - 1, y) + _
                         mBits(1, x - 1, y + 1) + mBits(1, x, y - 1) + _
                         mBits(1, x, y + 1) + mBits(1, x + 1, y - 1) + _
                         mBits(1, x + 1, y) + mBits(1, x + 1, y + 1)
                R = CLng(mBits(2, x - 1, y - 1)) + mBits(2, x - 1, y) + _
                         mBits(2, x - 1, y + 1) + mBits(2, x, y - 1) + _
                         mBits(2, x, y + 1) + mBits(2, x + 1, y - 1) + _
                         mBits(2, x + 1, y) + mBits(2, x + 1, y + 1)
                B = 8 * mBits(0, x, y) - B + 255
                G = 8 * mBits(1, x, y) - G + 255
                R = 8 * mBits(2, x, y) - R + 255
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                If ((B + G + R) \ 3 >= cF) Then
                    B = Bb
                    G = Gb
                    R = rB
                End If
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Contour = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Contrast(ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim Spd(765) As Long
  Dim R As Long, G As Long, B As Long
  Dim cF As Single
  Dim mCol As Long, nCol As Long

    tBits = mBits

    For x = 0 To 765
        Spd(x) = x \ 3
    Next x

    mCol = 0
    For y = 0 To biH
        For x = 0 To biW
            B = mBits(0, x, y)
            G = mBits(1, x, y)
            R = mBits(2, x, y)
            mCol = mCol + Spd(R + G + B)
            nCol = nCol + 1
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    RaiseEvent Progress(100)

    mCol = mCol \ nCol
    cF = (Factor + 100) / 100
    For x = 0 To 255
        Spd(x) = (x - mCol) * cF + mCol
    Next x

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                B = Spd(mBits(0, x, y))
                G = Spd(mBits(1, x, y))
                R = Spd(mBits(2, x, y))
                Do While (B < 0) Or (B > 255) Or (G < 0) Or (G > 255) Or (R < 0) Or (R > 255)
                    If (B <= 0) And (G <= 0) And (R <= 0) Then
                        B = 0
                        G = 0
                        R = 0
                    End If
                    If (B >= 255) And (G >= 255) And (R >= 255) Then
                        B = 255
                        G = 255
                        R = 255
                    End If
                    If (B < 0) Then
                        G = G + B \ 2
                        R = R + B \ 2
                        B = 0
                    End If
                    If (B > 255) Then
                        G = G + (B - 255) \ 2
                        R = R + (B - 255) \ 2
                        B = 255
                    End If
                    If (G < 0) Then
                        B = B + G \ 2
                        R = R + G \ 2
                        G = 0
                    End If
                    If (G > 255) Then
                        B = B + (G - 255) \ 2
                        R = R + (G - 255) \ 2
                        G = 255
                    End If
                    If (R < 0) Then
                        G = G + R \ 2
                        B = B + R \ 2
                        R = 0
                    End If
                    If (R > 255) Then
                        G = G + (R - 255) \ 2
                        B = B + (R - 255) \ 2
                        R = 255
                    End If
                Loop
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Contrast = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Diffuse(ByVal wFactor As Long, ByVal hFactor As Long) As Byte()

  Dim tBits() As Byte
  Dim cW As Single
  Dim cH As Single

    tBits = mBits

    cW = wFactor / 2
    cH = hFactor / 2

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            xOut = x + Rnd * wFactor - cW
            yOut = y + Rnd * hFactor - cH
            If (PtInRegion(biRgn, xOut, yOut)) Then
                tBits(0, x, y) = mBits(0, xOut, yOut)
                tBits(1, x, y) = mBits(1, xOut, yOut)
                tBits(2, x, y) = mBits(2, xOut, yOut)
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Diffuse = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Dissolve() As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim Med(1 To 4) As Long
  Dim Dev(1 To 4) As Long
  Dim i As Long, j As Long
  Dim sDev As Long, vDev As Long

    tBits = mBits

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                For i = 0 To 2
                    Med(1) = CLng(mBits(i, x - 2, y - 2)) + mBits(i, x - 1, y - 2) + mBits(i, x, y - 2) + _
                                  mBits(i, x - 2, y - 1) + mBits(i, x - 1, y - 1) + mBits(i, x, y - 1) + _
                                  mBits(i, x - 2, y) + mBits(i, x - 1, y) + mBits(i, x, y)
                    Med(2) = CLng(mBits(i, x + 2, y - 2)) + mBits(i, x + 1, y - 2) + mBits(i, x, y - 2) + _
                                  mBits(i, x + 2, y - 1) + mBits(i, x + 1, y - 1) + mBits(i, x, y - 1) + _
                                  mBits(i, x + 2, y) + mBits(i, x + 1, y) + mBits(i, x, y)
                    Med(3) = CLng(mBits(i, x - 2, y + 2)) + mBits(i, x - 1, y + 2) + mBits(i, x, y + 2) + _
                                  mBits(i, x - 2, y + 1) + mBits(i, x - 1, y + 1) + mBits(i, x, y + 1) + _
                                  mBits(i, x - 2, y) + mBits(i, x - 1, y) + mBits(i, x, y)
                    Med(4) = CLng(mBits(i, x + 2, y + 2)) + mBits(i, x + 1, y + 2) + mBits(i, x, y + 2) + _
                                  mBits(i, x + 2, y + 1) + mBits(i, x + 1, y + 1) + mBits(i, x, y + 1) + _
                                  mBits(i, x + 2, y) + mBits(i, x + 1, y) + mBits(i, x, y)
                    Med(1) = Med(1) \ 9
                    Med(2) = Med(2) \ 9
                    Med(3) = Med(3) \ 9
                    Med(4) = Med(4) \ 9
                    Dev(1) = Abs(mBits(i, x - 2, y - 2) - Med(1)) + Abs(mBits(i, x - 1, y - 2) - Med(1)) + Abs(mBits(i, x, y - 2) - Med(1)) + _
                             Abs(mBits(i, x - 2, y - 1) - Med(1)) + Abs(mBits(i, x - 1, y - 1) - Med(1)) + Abs(mBits(i, x, y - 1) - Med(1)) + _
                             Abs(mBits(i, x - 2, y) - Med(1)) + Abs(mBits(i, x - 1, y) - Med(1)) + Abs(mBits(i, x, y)) - Med(1)
                    Dev(2) = Abs(mBits(i, x + 2, y - 2) - Med(2)) + Abs(mBits(i, x + 1, y - 2) - Med(2)) + Abs(mBits(i, x, y - 2)) - Med(2) + _
                             Abs(mBits(i, x + 2, y - 1) - Med(2)) + Abs(mBits(i, x + 1, y - 1) - Med(2)) + Abs(mBits(i, x, y - 1)) - Med(2) + _
                             Abs(mBits(i, x + 2, y) - Med(2)) + Abs(mBits(i, x + 1, y) - Med(2)) + Abs(mBits(i, x, y)) - Med(2)
                    Dev(3) = Abs(mBits(i, x - 2, y + 2) - Med(3)) + Abs(mBits(i, x - 1, y + 2) - Med(3)) + Abs(mBits(i, x, y + 2) - Med(3)) + _
                             Abs(mBits(i, x - 2, y + 1) - Med(3)) + Abs(mBits(i, x - 1, y + 1) - Med(3)) + Abs(mBits(i, x, y + 1) - Med(3)) + _
                             Abs(mBits(i, x - 2, y) - Med(3)) + Abs(mBits(i, x - 1, y) - Med(3)) + Abs(mBits(i, x, y)) - Med(3)
                    Dev(4) = Abs(mBits(i, x + 2, y + 2) - Med(4)) + Abs(mBits(i, x + 1, y + 2) - Med(4)) + Abs(mBits(i, x, y + 2) - Med(4)) + _
                             Abs(mBits(i, x + 2, y + 1) - Med(4)) + Abs(mBits(i, x + 1, y + 1) - Med(4)) + Abs(mBits(i, x, y + 1) - Med(4)) + _
                             Abs(mBits(i, x + 2, y) - Med(4)) + Abs(mBits(i, x + 1, y) - Med(4)) + Abs(mBits(i, x, y) - Med(4))
                    vDev = 999999
                    sDev = 0
                    For j = 1 To 4
                        If (Dev(j) < vDev) Then
                            vDev = Dev(j)
                            sDev = j
                        End If
                    Next j
                    tBits(i, x, y) = Med(sDev)
                Next i
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Dissolve = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Emboss(ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim gOut As Long
  Dim cF As Single

    tBits = mBits

    cF = Factor / 10

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                If (x = biW) Or (y = biH) Then
                    If (x = biW) Then
                        tBits(0, x, y) = tBits(0, biW - 1, y)
                        tBits(1, x, y) = tBits(1, biW - 1, y)
                        tBits(2, x, y) = tBits(2, biW - 1, y)
                      Else
                        tBits(0, x, y) = tBits(0, x, biH - 1)
                        tBits(1, x, y) = tBits(1, x, biH - 1)
                        tBits(2, x, y) = tBits(2, x, biH - 1)
                    End If
                  Else
                    B = Abs((CLng(mBits(0, x, y)) - mBits(0, x + 1, y + 1)) * cF + 128)
                    G = Abs((CLng(mBits(1, x, y)) - mBits(1, x + 1, y + 1)) * cF + 128)
                    R = Abs((CLng(mBits(2, x, y)) - mBits(2, x + 1, y + 1)) * cF + 128)
                    gOut = (B + G + R) \ 3
                    If (gOut < 0) Then gOut = 0 Else If (gOut > 255) Then gOut = 255
                    tBits(0, x, y) = gOut
                    tBits(1, x, y) = gOut
                    tBits(2, x, y) = gOut
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Emboss = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Engrave(ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim gOut As Long
  Dim cF As Single

    tBits = mBits

    cF = Factor / 10

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                If (x = biW) Or (y = biH) Then
                    If (x = biW) Then
                        tBits(0, x, y) = tBits(0, biW - 1, y)
                        tBits(1, x, y) = tBits(1, biW - 1, y)
                        tBits(2, x, y) = tBits(2, biW - 1, y)
                      Else
                        tBits(0, x, y) = tBits(0, x, biH - 1)
                        tBits(1, x, y) = tBits(1, x, biH - 1)
                        tBits(2, x, y) = tBits(2, x, biH - 1)
                    End If
                  Else
                    B = Abs((CLng(mBits(0, x + 1, y + 1)) - mBits(0, x, y)) * cF + 128)
                    G = Abs((CLng(mBits(1, x + 1, y + 1)) - mBits(1, x, y)) * cF + 128)
                    R = Abs((CLng(mBits(2, x + 1, y + 1)) - mBits(2, x, y)) * cF + 128)
                    gOut = (B + G + R) \ 3
                    If (gOut < 0) Then gOut = 0 Else If (gOut > 255) Then gOut = 255
                    tBits(0, x, y) = gOut
                    tBits(1, x, y) = gOut
                    tBits(2, x, y) = gOut
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Engrave = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Equalize() As Byte()

  Dim H(255) As Long, L As Long
  Dim tBits() As Byte
  Dim Avg As Long
  Dim Sum As Long, Cum As Long
  Dim UF As Single, Lf As Single
  Dim i As Long

    tBits = mBits

    '// Get Histogram
    For y = 0 To biH
        For x = 0 To biW
            L = 0.113 * tBits(0, x, y) + _
                0.586 * tBits(1, x, y) + _
                0.301 * tBits(2, x, y)
            H(L) = H(L) + 1
        Next x
    Next y

    '// Find Avg, range factors
    Sum = 0
    Cum = ((biW + 1) * (biH + 1)) \ 2
    For x = 0 To 255
        Sum = Sum + H(x)
        If (Sum > Cum) Then Exit For
    Next x
    If (H(x) \ 2 > Sum - Cum) Then
        Avg = x
      Else
        Avg = x - 1
    End If
    UF = 127 / (255 - Avg)
    Lf = 127 / (Avg - 1)

    '// Equalize
    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                For i = 0 To 2
                    If (mBits(i, x, y) < Avg) Then
                        tBits(i, x, y) = (mBits(i, x, y)) * Lf
                      Else
                        tBits(i, x, y) = (mBits(i, x, y) - Avg) * UF + 128
                    End If
                Next i
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Equalize = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Fold(ByVal wFactor As Long, ByVal hFactor As Long, Optional ByVal CleanBack As Boolean = -1) As Byte()

  Dim tBits() As Byte
  Dim cW As Single
  Dim cH As Single
  Dim ox As Long, oy As Long

    tBits = mBits

    cW = (1 + wFactor / 100)
    cH = (1 + hFactor / 100)
    ox = biW \ 2
    oy = biH \ 2

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            xIn = (x - ox) * (1 + ((y - oy) / 100 * hFactor) / biH) + ox
            yIn = (y - oy) * (1 + ((x - ox) / 100 * wFactor) / biW) + oy
            If (PtInRegion(biRgn, xIn, yIn)) Then
                tBits(0, x, y) = mBits(0, xIn, yIn)
                tBits(1, x, y) = mBits(1, xIn, yIn)
                tBits(2, x, y) = mBits(2, xIn, yIn)
              Else
                If (CleanBack) Then
                    tBits(0, x, y) = 0
                    tBits(1, x, y) = 0
                    tBits(2, x, y) = 0
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Fold = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Flip(Optional ByVal fH As Boolean = 0, Optional ByVal fV As Boolean = 0) As Byte()

  Dim tBits() As Byte

    tBits = mBits

    For y = 0 To biH
        If (fV) Then yIn = biH - y Else yIn = y
        For x = 0 To biW
            If (fH) Then xIn = biW - x Else xIn = x
            '// Whole image / no region limited
            tBits(0, x, y) = mBits(0, xIn, yIn)
            tBits(1, x, y) = mBits(1, xIn, yIn)
            tBits(2, x, y) = mBits(2, xIn, yIn)
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Flip = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function GammaCorrection(ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim dB As Double, dG As Double, dR As Double
  Dim cF As Single

    tBits = mBits

    cF = (Factor + 100) / 100

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                '// Get data
                B = mBits(0, x, y)
                G = mBits(1, x, y)
                R = mBits(2, x, y)
                If (cF > 0) Then
                    dB = B / 255
                    dG = G / 255
                    dR = R / 255
                    '// Correct gamma
                    dB = dB ^ (1 / cF)
                    dG = dG ^ (1 / cF)
                    dR = dR ^ (1 / cF)
                    '// Set data
                    B = dB * 255
                    G = dG * 255
                    R = dR * 255
                  Else
                    B = 0
                    G = 0
                    R = 0
                End If
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    GammaCorrection = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function GradientCircular(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Pressure As Long, ByVal Frequency As Long) As Byte()

  Dim tBits() As Byte

  Dim cR1 As Long, cG1 As Long, cB1 As Long
  Dim cR2 As Long, cG2 As Long, cB2 As Long

  Dim rG As Long
  Dim yAbs As Long
  Dim c1 As Single, c2 As Single
  Dim cp1 As Single, cp2 As Single
  Dim cF As Single

    If (Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2) = 0) Then
        GradientCircular = mBits
        Exit Function
    End If

    tBits = mBits

    cR1 = (Color1 And &HFF&)
    cG1 = (Color1 And &HFF00&) \ 256
    cB1 = (Color1 And &HFF0000) \ 65536
    cR2 = (Color2 And &HFF&)
    cG2 = (Color2 And &HFF00&) \ 256
    cB2 = (Color2 And &HFF0000) \ 65536

    rG = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)

    cp1 = Pressure / 100
    cp2 = 1 - cp1

    cF = (Frequency - 1) * PI
    If (cF > rG) Then cF = rG

    xStart = -x1
    xEnd = biW - x1
    yStart = -y1
    yEnd = biH - y1
    yAbs = yEnd - yStart

    For y = yStart To yEnd
        yOut = y + y1
        For x = xStart To xEnd
            xOut = x + x1
            If (PtInRegion(biRgn, xOut, yOut)) Then
                c1 = Sqr(x * x + y * y) / rG
                If (cF > 0) Then c1 = Sin(c1 * cF): c1 = c1 * c1
                If (c1 > 1) Then c1 = 1
                c2 = 1 - c1
                If (Color1 = -1) Then
                    cB1 = mBits(0, xOut, yOut)
                    cG1 = mBits(1, xOut, yOut)
                    cR1 = mBits(2, xOut, yOut)
                End If
                If (Color2 = -1) Then
                    cB2 = mBits(0, xOut, yOut)
                    cG2 = mBits(1, xOut, yOut)
                    cR2 = mBits(2, xOut, yOut)
                End If
                tBits(0, xOut, yOut) = cp1 * (c1 * cB2 + c2 * cB1) + cp2 * mBits(0, xOut, yOut)
                tBits(1, xOut, yOut) = cp1 * (c1 * cG2 + c2 * cG1) + cp2 * mBits(1, xOut, yOut)
                tBits(2, xOut, yOut) = cp1 * (c1 * cR2 + c2 * cR1) + cp2 * mBits(2, xOut, yOut)
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress((y - yStart) * 100 \ yAbs)
        End If
    Next y
    GradientCircular = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function GradientLinear(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Pressure As Long, ByVal Frequency As Long) As Byte()

  Dim tBits() As Byte

  Dim cR1 As Long, cG1 As Long, cB1 As Long
  Dim cR2 As Long, cG2 As Long, cB2 As Long

  Dim dG As Long, degG As Single
  Dim yAbs As Long
  Dim xC As Single, yC As Single
  Dim c1 As Single, c2 As Single
  Dim cp1 As Single, cp2 As Single
  Dim cF As Single

    If (Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2) = 0) Then
        GradientLinear = mBits
        Exit Function
    End If

    tBits = mBits

    cR1 = (Color1 And &HFF&)
    cG1 = (Color1 And &HFF00&) \ 256
    cB1 = (Color1 And &HFF0000) \ 65536
    cR2 = (Color2 And &HFF&)
    cG2 = (Color2 And &HFF00&) \ 256
    cB2 = (Color2 And &HFF0000) \ 65536

    dG = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)

    If (x2 - x1 <> 0) Then
        degG = Atn((y2 - y1) / (x2 - x1))
      Else
        degG = Sgn((y2 - y1)) * 90 * PIdiv180
    End If

    If (x2 - x1 < 0) Then
        degG = degG + 180 * PIdiv180
      Else
        If (y2 - y1 < 0) Then
            degG = degG + 360 * PIdiv180
        End If
    End If

    cp1 = Pressure / 100
    cp2 = 1 - cp1

    cF = (Frequency - 1) * PI
    If (cF > dG) Then cF = dG

    xStart = -x1
    xEnd = biW - x1
    yStart = -y1
    yEnd = biH - y1
    yAbs = yEnd - yStart

    xC = Cos(degG) / dG
    yC = Sin(degG) / dG

    For y = yStart To yEnd
        yOut = y + y1
        For x = xStart To xEnd
            xOut = x + x1
            If (PtInRegion(biRgn, xOut, yOut)) Then
                c1 = x * xC + y * yC
                If (cF > 0) Then c1 = Sin(c1 * cF): c1 = c1 * c1
                If (c1 > 1) Then c1 = 1
                If (c1 < 0) Then c1 = 0
                c2 = 1 - c1
                If (Color1 = -1) Then
                    cB1 = mBits(0, xOut, yOut)
                    cG1 = mBits(1, xOut, yOut)
                    cR1 = mBits(2, xOut, yOut)
                End If
                If (Color2 = -1) Then
                    cB2 = mBits(0, xOut, yOut)
                    cG2 = mBits(1, xOut, yOut)
                    cR2 = mBits(2, xOut, yOut)
                End If
                tBits(0, xOut, yOut) = cp1 * (c1 * cB2 + c2 * cB1) + cp2 * mBits(0, xOut, yOut)
                tBits(1, xOut, yOut) = cp1 * (c1 * cG2 + c2 * cG1) + cp2 * mBits(1, xOut, yOut)
                tBits(2, xOut, yOut) = cp1 * (c1 * cR2 + c2 * cR1) + cp2 * mBits(2, xOut, yOut)
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress((y - yStart) * 100 \ yAbs)
        End If
    Next y
    GradientLinear = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function GreyScale() As Byte()

  Dim tBits() As Byte
  Dim gOut As Long

    tBits = mBits

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
            '// gOut = 0.114 * mBits(0, x, y) + 0.587 * mBits(1, x, y) + 0.299 * mBits(2, x, y)
                gOut = 0.071 * mBits(0, x, y) + 0.707 * mBits(1, x, y) + 0.222 * mBits(2, x, y)
                tBits(0, x, y) = gOut
                tBits(1, x, y) = gOut
                tBits(2, x, y) = gOut
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    GreyScale = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Grid(ByVal HStep As Long, ByVal VStep As Long, ByVal ForceMargins As Boolean, ByVal Color As Long, Optional ByVal Opacity As Long = 255) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim cO As Single, cOBack As Single

    tBits = mBits

    R = (Color And &HFF&)
    G = (Color And &HFF00&) \ 256
    B = (Color And &HFF0000) \ 65536
    
    tBits = mBits

    cO = Opacity / 255
    cOBack = 1 - cO

    GetRgnBoundRect
    
    For y = yStart To yEnd Step VStep
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                tBits(0, x, y) = cO * B + cOBack * mBits(0, x, y)
                tBits(1, x, y) = cO * G + cOBack * mBits(1, x, y)
                tBits(2, x, y) = cO * R + cOBack * mBits(2, x, y)
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    If (ForceMargins) Then
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, biH)) Then
                tBits(0, x, biH) = cO * B + cOBack * mBits(0, x, biH)
                tBits(1, x, biH) = cO * G + cOBack * mBits(1, x, biH)
                tBits(2, x, biH) = cO * R + cOBack * mBits(2, x, biH)
            End If
        Next x
    End If
    For x = xStart To xEnd Step HStep
        For y = xStart To yEnd
            If (PtInRegion(biRgn, x, y)) Then
                tBits(0, x, y) = cO * B + cOBack * mBits(0, x, y)
                tBits(1, x, y) = cO * G + cOBack * mBits(1, x, y)
                tBits(2, x, y) = cO * R + cOBack * mBits(2, x, y)
            End If
        Next y
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(x * 100 \ biW)
        End If
    Next x
    If (ForceMargins) Then
        For y = yStart To yEnd
            If (PtInRegion(biRgn, biW, y)) Then
                tBits(0, biW, y) = cO * B + cOBack * mBits(0, biW, y)
                tBits(1, biW, y) = cO * G + cOBack * mBits(1, biW, y)
                tBits(2, biW, y) = cO * R + cOBack * mBits(2, biW, y)
            End If
        Next y
    End If
    Grid = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Halftone(ByVal Radius As Long, Optional ByVal Definition As Long = 0, Optional ByVal Intensity As Long = 0) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim rH As Long, dH As Long, rHpow2 As Long
  Dim cD As Single, cI As Single
  Dim mx As Long, my As Long
  Dim mC As Long
  Dim ci1 As Single, ci2 As Single

    tBits = mBits

    rH = Radius
    dH = 2 * Radius

    cD = (Definition + rH) / rH
    cI = (Intensity + 100) / 100
    rHpow2 = cD * Radius * Radius

    GetRgnBoundRect

    For y = yStart To yEnd Step dH
        For x = xStart To xEnd Step dH
            B = 0
            G = 0
            R = 0
            mC = 0
            For my = y To y + dH
                For mx = x To x + dH
                    If (PtInRegion(biRgn, mx, my)) Then
                        B = B + mBits(0, mx, my)
                        G = G + mBits(1, mx, my)
                        R = R + mBits(2, mx, my)
                        mC = mC + 1
                    End If
                Next mx
            Next my
            mC = mC / cI
            If (mC > 0) Then
                B = B / mC
                G = G / mC
                R = R / mC
                If (B > 255) Then B = 255
                If (G > 255) Then G = 255
                If (R > 255) Then R = 255
            End If

            For my = -rH To rH
                yOut = my + y + rH
                For mx = -rH To rH
                    xOut = mx + x + rH
                    If (PtInRegion(biRgn, xOut, yOut)) Then
                        If (mx * mx + my * my <= rHpow2) Then
                            ci1 = 1 - (mx * mx + my * my) / rHpow2
                            tBits(0, xOut, yOut) = ci1 * B
                            tBits(1, xOut, yOut) = ci1 * G
                            tBits(2, xOut, yOut) = ci1 * R
                          Else
                            tBits(0, xOut, yOut) = 0
                            tBits(1, xOut, yOut) = 0
                            tBits(2, xOut, yOut) = 0
                        End If
                    End If
                Next mx
            Next my
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Halftone = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Invert() As Byte()

  Dim tBits() As Byte
  Dim Spd(255) As Long

    tBits = mBits

    For x = 0 To 255
        Spd(x) = 255 - x
    Next x

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                tBits(0, x, y) = Spd(mBits(0, x, y))
                tBits(1, x, y) = Spd(mBits(1, x, y))
                tBits(2, x, y) = Spd(mBits(2, x, y))
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Invert = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Isolate(ByVal Color As Long, ByVal EraseColor As Long, ByVal Tolerance As Long) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim Re As Long, Ge As Long, Be As Long
  Dim tF As Long
  Dim distance As Long

    tBits = mBits

    R = (Color And &HFF&)
    G = (Color And &HFF00&) \ 256
    B = (Color And &HFF0000) \ 65536
    Re = (EraseColor And &HFF&)
    Ge = (EraseColor And &HFF00&) \ 256
    Be = (EraseColor And &HFF0000) \ 65536
    tF = (127.5 * Tolerance / 100) ^ 2 + 1

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                distance = (mBits(0, x, y) - B) * (mBits(0, x, y) - B) _
                         + (mBits(1, x, y) - G) * (mBits(1, x, y) - G) _
                         + (mBits(2, x, y) - R) * (mBits(2, x, y) - R)
                If (distance >= tF) Then
                    tBits(0, x, y) = Be
                    tBits(1, x, y) = Ge
                    tBits(2, x, y) = Re
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Isolate = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function KillChannel(ByVal rChannel As Boolean, ByVal gChannel As Boolean, ByVal bChannel As Boolean) As Byte()

  Dim tBits() As Byte

    tBits = mBits

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                If (bChannel) Then tBits(0, x, y) = 0
                If (gChannel) Then tBits(1, x, y) = 0
                If (rChannel) Then tBits(2, x, y) = 0
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    KillChannel = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Lens(ByVal xRel As Single, yRel As Single, dRel As Single, ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim xOffset As Long, yOffset As Long
  Dim xL As Long, yL As Long
  Dim cF As Single
  Dim rL As Long, sL As Single, zL As Long
  Dim rLpow2 As Long, sLpow2 As Long

    tBits = mBits

    xOffset = xRel * biW
    yOffset = yRel * biH

    If (biW > biH) Then
        rL = dRel * biW
      Else
        rL = dRel * biH
    End If
    cF = rL * (150 - Factor) / 150
    sL = Sqr(rL * rL - cF * cF)
    rLpow2 = rL * rL
    sLpow2 = sL * sL

    For y = -rL To rL
        yOut = y + yOffset
        For x = -rL To rL
            If (x * x + y * y < sLpow2) Then
                xOut = x + xOffset
                If (PtInRegion(biRgn, xOut, yOut)) Then
                    zL = Sqr(rLpow2 - x * x - y * y)
                    xL = (x * cF) \ zL + xOffset
                    yL = (y * cF) \ zL + yOffset
                    tBits(0, xOut, yOut) = mBits(0, xL, yL)
                    tBits(1, xOut, yOut) = mBits(1, xL, yL)
                    tBits(2, xOut, yOut) = mBits(2, xL, yL)
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress((y + rL) * 100 \ (2 * rL))
        End If
    Next y
    Lens = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Light(ByVal xRel As Single, ByVal yRel As Single, ByVal dRel As Single, ByVal iFactor As Long, ByVal dFactor As Long) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim xOffset As Long, yOffset As Long
  Dim rL As Long, rLpow2 As Long
  Dim incL As Long
  Dim defL As Single

    tBits = mBits

    xOffset = xRel * biW
    yOffset = yRel * biH

    If (biW > biH) Then
        rL = dRel * biW
      Else
        rL = dRel * biH
    End If
    rLpow2 = rL * rL
    defL = 1 - (dFactor / 100)

    For y = -rL To rL
        yOut = y + yOffset
        For x = -rL To rL
            xOut = x + xOffset
            If (x * x + y * y < rLpow2) Then
                If (PtInRegion(biRgn, xOut, yOut)) Then
                    incL = (1 - defL * Sqr(x * x + y * y) / rL) * iFactor
                    B = mBits(0, xOut, yOut) + incL
                    G = mBits(1, xOut, yOut) + incL
                    R = mBits(2, xOut, yOut) + incL
                    If (B > 255) Then B = 255
                    If (G > 255) Then G = 255
                    If (R > 255) Then R = 255
                    tBits(0, xOut, yOut) = B
                    tBits(1, xOut, yOut) = G
                    tBits(2, xOut, yOut) = R
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress((y + rL) * 100 \ (2 * rL))
        End If
    Next y
    Light = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function MaskFilter(ByRef mskBits() As Long, ByVal Weight, Optional ByVal rFactor As Long, Optional ByVal gFactor As Long, Optional ByVal bFactor As Long, Optional ByVal KeepMargins As Boolean = 0) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim mx As Long, my As Long
  Dim mSize As Long, mSizediv2 As Long

    tBits = mBits

    mSize = UBound(mskBits)
    mSizediv2 = mSize \ 2

    GetRgnBoundRect
    xStart = xStart - mSizediv2
    yStart = yStart - mSizediv2
    xEnd = xEnd - mSizediv2
    yEnd = yEnd - mSizediv2

    If (KeepMargins) Then
        If (xStart < 0) Then xStart = 0
        If (yStart < 0) Then yStart = 0
        If (xEnd > biW - mSize) Then xEnd = biW - mSize
        If (yEnd > biH - mSize) Then yEnd = biH - mSize
    End If

    For y = yStart To yEnd
        For x = xStart To xEnd
            B = 0
            G = 0
            R = 0
            For my = 0 To mSize
                For mx = 0 To mSize
                    xIn = x + my '!
                    yIn = y + mx '!
                    B = B + mskBits(mx, my) * mBits(0, xIn, yIn)
                    G = G + mskBits(mx, my) * mBits(1, xIn, yIn)
                    R = R + mskBits(mx, my) * mBits(2, xIn, yIn)
                Next mx
            Next my
            xOut = x + mSizediv2
            yOut = y + mSizediv2
            If (PtInRegion(biRgn, xOut, yOut)) Then
                B = B \ Weight + bFactor
                G = G \ Weight + gFactor
                R = R \ Weight + rFactor
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                tBits(0, xOut, yOut) = B
                tBits(1, xOut, yOut) = G
                tBits(2, xOut, yOut) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    MaskFilter = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Noise(ByVal Factor As Long, Optional Equalized As Boolean = 0) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim Rf As Single, Gf As Single, Bf As Single
  Dim cf1 As Long, cf2 As Long

    tBits = mBits

    cf1 = Factor
    cf2 = Factor * 2 + 1

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                If (Equalized) Then
                    Bf = Rnd
                    Gf = Bf
                    Rf = Bf
                  Else
                    Bf = Rnd
                    Gf = Rnd
                    Rf = Rnd
                End If
                B = CLng(mBits(0, x, y)) + (cf2 * Rf - cf1)
                G = CLng(mBits(1, x, y)) + (cf2 * Gf - cf1)
                R = CLng(mBits(2, x, y)) + (cf2 * Bf - cf1)
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Noise = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function PaletteApply() As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim errR As Long, errG As Long, errB As Long
  Dim newR As Long, newG As Long, newB As Long
  Dim i As Long, iCoeff As Long

    tBits = mBits

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                ClosestColor mBits(2, x, y), mBits(1, x, y), mBits(0, x, y), R, G, B
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R

                errB = -(CLng(tBits(0, x, y)) - mBits(0, x, y))
                errG = -(CLng(tBits(1, x, y)) - mBits(1, x, y))
                errR = -(CLng(tBits(2, x, y)) - mBits(2, x, y))

                If (Abs(errB) + Abs(errG) + Abs(errR) > 3) Then
                    If (x < biW) Then
                        xIn = x + 1
                        newB = mBits(0, xIn, y) + (errB * 7) \ 16
                        newG = mBits(1, xIn, y) + (errG * 7) \ 16
                        newR = mBits(2, xIn, y) + (errR * 7) \ 16
                        If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                        If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                        If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                        mBits(0, xIn, y) = newB
                        mBits(1, xIn, y) = newG
                        mBits(2, xIn, y) = newR
                    End If
                    If (y < biH) Then
                        For i = -1 To 1
                            xIn = x + i
                            yIn = y + 1
                            If (xIn > 0) And (xIn < biW) Then
                                If (i = 0) Then
                                    iCoeff = 4
                                  Else
                                    iCoeff = 0
                                End If
                                newB = mBits(0, xIn, yIn) + (errB * iCoeff) \ 16
                                newG = mBits(1, xIn, yIn) + (errG * iCoeff) \ 16
                                newR = mBits(2, xIn, yIn) + (errR * iCoeff) \ 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                mBits(0, xIn, yIn) = newB
                                mBits(1, xIn, yIn) = newG
                                mBits(2, xIn, yIn) = newR
                            End If
                        Next i
                    End If
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    PaletteApply = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Pixelize(ByVal PixSize As Long) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim px As Long, py As Long
  Dim sx As Long, sy As Long
  Dim mC As Long

    tBits = mBits

    GetRgnBoundRect

    B = 0
    G = 0
    R = 0
    For y = yStart To yEnd
        sy = (y \ PixSize) * PixSize
        For x = xStart To xEnd
            sx = (x \ PixSize) * PixSize
            If (PtInRegion(biRgn, x, y)) Then
                If (x Mod PixSize) = 0 Then
                    B = 0
                    G = 0
                    R = 0
                    mC = 0
                    For px = sx To sx + (PixSize - 1)
                        For py = sy To sy + (PixSize - 1)
                            If (px <= biW) And (py <= biH) Then
                                B = B + mBits(0, px, py)
                                G = G + mBits(1, px, py)
                                R = R + mBits(2, px, py)
                                mC = mC + 1
                            End If
                        Next py
                    Next px
                    If (mC > 0) Then
                        B = B \ mC
                        G = G \ mC
                        R = R \ mC
                    End If
                End If
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Pixelize = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function RankFilterN4(ByVal fltType As Long) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim lRgn As Long
  Dim i As Long, j As Long
  Dim rx As Long, ry As Long

    tBits = mBits

    GetRgnBoundRect 1 - fltType

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                If (fltType = 0) Then
                    B = 255
                    G = 255
                    R = 255
                  Else
                    B = 0
                    G = 0
                    R = 0
                End If
                For ry = -1 To 1
                    yIn = y + ry
                    For rx = -1 To 1
                        '// N4(p):
                        If (Abs(rx) <> Abs(ry) Or (rx = 0 And ry = 0)) Then
                            xIn = x + rx
                            Select Case fltType
                              Case 0  '// Minimum (enhance black)
                                If (mBits(2, xIn, yIn) < B) Then
                                    B = mBits(2, xIn, yIn)
                                End If
                                If (mBits(1, xIn, yIn) < G) Then
                                    G = mBits(1, xIn, yIn)
                                End If
                                If (mBits(0, xIn, yIn) < R) Then
                                    R = mBits(0, xIn, yIn)
                                End If
                              Case 1  '// Maximum (enhance white)
                                If (mBits(2, xIn, yIn) > B) Then
                                    B = mBits(2, xIn, yIn)
                                End If
                                If (mBits(1, xIn, yIn) > G) Then
                                    G = mBits(1, xIn, yIn)
                                End If
                                If (mBits(0, xIn, yIn) > R) Then
                                    R = mBits(0, xIn, yIn)
                                End If
                            End Select
                        End If
                    Next rx
                Next ry
                tBits(2, x, y) = B
                tBits(1, x, y) = G
                tBits(0, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    RankFilterN4 = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function RankFilterND(ByVal fltType As Long) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim lRgn As Long
  Dim i As Long, j As Long
  Dim rx As Long, ry As Long

    tBits = mBits

    GetRgnBoundRect 1 - fltType

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                If (fltType = 0) Then
                    B = 255
                    G = 255
                    R = 255
                  Else
                    B = 0
                    G = 0
                    R = 0
                End If
                For ry = -1 To 1
                    yIn = y + ry
                    For rx = -1 To 1
                        '// ND(p):
                        If (Abs(rx) = Abs(ry)) Then
                            xIn = x + rx
                            Select Case fltType
                              Case 0  '// Minimum (enhance black)
                                If (mBits(2, xIn, yIn) < B) Then
                                    B = mBits(2, xIn, yIn)
                                End If
                                If (mBits(1, xIn, yIn) < G) Then
                                    G = mBits(1, xIn, yIn)
                                End If
                                If (mBits(0, xIn, yIn) < R) Then
                                    R = mBits(0, xIn, yIn)
                                End If
                              Case 1  '// Maximum (enhance white)
                                If (mBits(2, xIn, yIn) > B) Then
                                    B = mBits(2, xIn, yIn)
                                End If
                                If (mBits(1, xIn, yIn) > G) Then
                                    G = mBits(1, xIn, yIn)
                                End If
                                If (mBits(0, xIn, yIn) > R) Then
                                    R = mBits(0, xIn, yIn)
                                End If
                            End Select
                        End If
                    Next rx
                Next ry
                tBits(2, x, y) = B
                tBits(1, x, y) = G
                tBits(0, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    RankFilterND = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Reduction(ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim Spd(255) As Long
  Dim cF As Long

    tBits = mBits

    cF = Factor + 1

    For x = 0 To 255
        Spd(x) = Round(x / cF) * cF
        If Spd(x) > 255 Then Spd(x) = 255
    Next x

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                tBits(0, x, y) = Spd(mBits(0, x, y))
                tBits(1, x, y) = Spd(mBits(1, x, y))
                tBits(2, x, y) = Spd(mBits(2, x, y))
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Reduction = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Replace(ByVal ColorFrom As Long, ByVal ColorTo As Long, ByVal Tolerance As Long) As Byte()

  Dim tBits() As Byte
  Dim Rf As Long, Gf As Long, Bf As Long
  Dim Rt As Long, Gt As Long, Bt As Long
  Dim tF As Long
  Dim distance As Long, ratio As Single

    tBits = mBits

    Rf = (ColorFrom And &HFF&)
    Gf = (ColorFrom And &HFF00&) \ 256
    Bf = (ColorFrom And &HFF0000) \ 65536
    Rt = (ColorTo And &HFF&)
    Gt = (ColorTo And &HFF00&) \ 256
    Bt = (ColorTo And &HFF0000) \ 65536
    tF = (127.5 * Tolerance / 100) ^ 2 + 1

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                distance = (mBits(0, x, y) - Bf) * (mBits(0, x, y) - Bf) _
                           + (mBits(1, x, y) - Gf) * (mBits(1, x, y) - Gf) _
                           + (mBits(2, x, y) - Rf) * (mBits(2, x, y) - Rf)
                If (distance <= tF) Then
                    ratio = 1 - distance / tF
                    tBits(0, x, y) = mBits(0, x, y) + (Bt - mBits(0, x, y)) * ratio
                    tBits(1, x, y) = mBits(1, x, y) + (Gt - mBits(1, x, y)) * ratio
                    tBits(2, x, y) = mBits(2, x, y) + (Rt - mBits(2, x, y)) * ratio
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Replace = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function ReplaceHS(ByVal ColorFrom As Long, ByVal ColorTo As Long, ByVal HTolerance As Long, ByVal STolerance As Long, ByVal Tolerance As Long) As Byte()

  Dim tBits() As Byte
  Dim Hf As Single, Sf As Single, Lf As Single
  Dim H As Single, S As Single, L As Single
  Dim Rf As Long, Gf As Long, Bf As Long
  Dim Rt As Long, Gt As Long, Bt As Long
  Dim distance As Long, ratio As Single
  Dim tF As Long
  Dim cTH As Single, cTS As Single

    tBits = mBits

    Rf = (ColorFrom And &HFF&)
    Gf = (ColorFrom And &HFF00&) \ 256
    Bf = (ColorFrom And &HFF0000) \ 65536
    Rt = (ColorTo And &HFF&)
    Gt = (ColorTo And &HFF00&) \ 256
    Bt = (ColorTo And &HFF0000) \ 65536

    RGBtoHSL Rf, Gf, Bf, Hf, Sf, Lf

    tF = (127.5 * Tolerance / 100) ^ 2 + 1
    cTH = HTolerance / 100
    cTS = STolerance / 100

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                RGBtoHSL mBits(2, x, y), mBits(1, x, y), mBits(0, x, y), H, S, L
                If (Abs(H - Hf) <= cTH And Abs(S - Sf) <= cTS) Then
                    distance = (mBits(0, x, y) - Bf) * (mBits(0, x, y) - Bf) _
                             + (mBits(1, x, y) - Gf) * (mBits(1, x, y) - Gf) _
                             + (mBits(2, x, y) - Rf) * (mBits(2, x, y) - Rf)
                    ratio = 1 - distance / tF
                    If (distance <= tF) Then
                        tBits(0, x, y) = mBits(0, x, y) + (Bt - mBits(0, x, y)) * ratio
                        tBits(1, x, y) = mBits(1, x, y) + (Gt - mBits(1, x, y)) * ratio
                        tBits(2, x, y) = mBits(2, x, y) + (Rt - mBits(2, x, y)) * ratio
                    End If
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    ReplaceHS = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function ReplaceL(ByVal ColorFrom As Long, ByVal ColorTo As Long, ByVal LTolerance As Long, ByVal Tolerance As Long) As Byte()

  Dim tBits() As Byte
  Dim Rf As Long, Gf As Long, Bf As Long
  Dim Rt As Long, Gt As Long, Bt As Long
  Dim L As Long
  Dim distance As Long, ratio As Single
  Dim tF As Long
  Dim cTL As Long

    tBits = mBits

    Rf = (ColorFrom And &HFF&)
    Gf = (ColorFrom And &HFF00&) \ 256
    Bf = (ColorFrom And &HFF0000) \ 65536
    Rt = (ColorTo And &HFF&)
    Gt = (ColorTo And &HFF00&) \ 256
    Bt = (ColorTo And &HFF0000) \ 65536

    tF = (127.5 * Tolerance / 100) ^ 2 + 1
    cTL = LTolerance * 2.55
    L = (Rf + Gf + Bf) \ 3

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                If (Abs((CLng(mBits(0, x, y)) + mBits(1, x, y) + mBits(2, x, y)) \ 3 - L) <= LTolerance) Then
                    distance = (mBits(0, x, y) - Bf) * (mBits(0, x, y) - Bf) _
                             + (mBits(1, x, y) - Gf) * (mBits(1, x, y) - Gf) _
                             + (mBits(2, x, y) - Rf) * (mBits(2, x, y) - Rf)
                    ratio = 1 - distance / tF
                    If (distance <= tF) Then
                        tBits(0, x, y) = mBits(0, x, y) + (Bt - mBits(0, x, y)) * ratio
                        tBits(1, x, y) = mBits(1, x, y) + (Gt - mBits(1, x, y)) * ratio
                        tBits(2, x, y) = mBits(2, x, y) + (Rt - mBits(2, x, y)) * ratio
                    End If
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    ReplaceL = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Rotate(ByVal Degrees As Long, Optional ByVal CleanBack As Boolean = -1) As Byte()

  Dim tBits() As Byte
  Dim cosSpd As Single
  Dim sinSpd As Single
  Dim ox As Long, oy As Long
  Dim rx As Long, ry As Long

    tBits = mBits

    ox = biW \ 2
    oy = biH \ 2
    cosSpd = Cos(Degrees * PIdiv180)
    sinSpd = Sin(Degrees * PIdiv180)

    GetRgnBoundRect

    For y = yStart To yEnd
        yIn = y - oy
        For x = xStart To xEnd
            xIn = x - ox
            rx = xIn * cosSpd - yIn * sinSpd + ox
            ry = xIn * sinSpd + yIn * cosSpd + oy
            If (PtInRegion(biRgn, rx, ry)) Then
                tBits(0, x, y) = mBits(0, rx, ry)
                tBits(1, x, y) = mBits(1, rx, ry)
                tBits(2, x, y) = mBits(2, rx, ry)
              Else
                If (CleanBack) Then
                    tBits(0, x, y) = 0
                    tBits(1, x, y) = 0
                    tBits(2, x, y) = 0
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Rotate = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function RotateH(ByVal HOffset As Long, ByVal Saturation As Long)

  '// Need mHSLRGB module

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim H As Single, S As Single, L As Single
  Dim cHOf As Single
  Dim cSat As Single
  
    tBits = mBits

    cHOf = HOffset / 16.67
    cSat = Saturation / 100

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                RGBtoHSL mBits(2, x, y), mBits(1, x, y), mBits(0, x, y), H, S, L
                H = H + cHOf
                If (H > 5) Then H = H - 6
                HSLtoRGB H, S * cSat, L, R, G, B
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    RotateH = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function RotateL(ByVal LOffset As Long, ByVal LStep As Long)

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim H As Single, S As Single, L As Single
  
  Dim cOff As Single
  Dim cStp As Single

    tBits = mBits

    cOff = LOffset / 10
    cStp = LStep / 2.5

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                L = (0.071 * mBits(0, x, y) + 0.707 * mBits(1, x, y) + 0.222 * mBits(2, x, y)) / 255
                L = 127.5 * Sin(cStp * L + cOff) + 127.5
                If (L > 255) Then L = L - 255 Else If (L < 0) Then L = L + 255
                tBits(0, x, y) = L
                tBits(1, x, y) = L
                tBits(2, x, y) = L
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    RotateL = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function RotateSmooth(ByVal Degrees As Long) As Byte()

  Dim tBits() As Byte
  Dim cosSpd As Single
  Dim sinSpd As Single

  Dim R As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single
  Dim G As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single
  Dim B As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single

  Dim ir1 As Long, ig1 As Long, ib1 As Long
  Dim ir2 As Long, ig2 As Long, ib2 As Long

  Dim ox As Long, oy As Long
  Dim rx As Single, ry As Single
  Dim drx As Single, dry As Single
  Dim irx As Long, iry As Long

    tBits = mBits

    ox = biW \ 2
    oy = biH \ 2
    cosSpd = Cos(Degrees * PIdiv180)
    sinSpd = Sin(Degrees * PIdiv180)

    GetRgnBoundRect

    For y = yStart To yEnd
        yIn = y - oy
        For x = xStart To xEnd
            xIn = x - ox
            rx = xIn * cosSpd - yIn * sinSpd + ox
            ry = xIn * sinSpd + yIn * cosSpd + oy
            irx = Int(rx)
            iry = Int(ry)
            '// Whole image
            If (PtInRect(biRct, irx, iry)) Then
                drx = rx - irx
                dry = ry - iry
                '// Interpolate using the four nearest pixels in the source
                b1 = mBits(0, irx, iry)
                g1 = mBits(1, irx, iry)
                r1 = mBits(2, irx, iry)
                b2 = mBits(0, irx + 1, iry)
                g2 = mBits(1, irx + 1, iry)
                r2 = mBits(2, irx + 1, iry)
                b3 = mBits(0, irx, iry + 1)
                g3 = mBits(1, irx, iry + 1)
                r3 = mBits(2, irx, iry + 1)
                b4 = mBits(0, irx + 1, iry + 1)
                g4 = mBits(1, irx + 1, iry + 1)
                r4 = mBits(2, irx + 1, iry + 1)
                '// Interpolate in y direction:
                ib1 = b1 * (1 - dry) + b3 * dry
                ig1 = g1 * (1 - dry) + g3 * dry
                ir1 = r1 * (1 - dry) + r3 * dry
                ib2 = b2 * (1 - dry) + b4 * dry
                ig2 = g2 * (1 - dry) + g4 * dry
                ir2 = r2 * (1 - dry) + r4 * dry
                '// Interpolate in x direction:
                B = ib1 * (1 - drx) + ib2 * drx
                G = ig1 * (1 - drx) + ig2 * drx
                R = ir1 * (1 - drx) + ir2 * drx
                '// Check ranges
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                '// Set output:
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
              Else
                tBits(0, x, y) = 0
                tBits(1, x, y) = 0
                tBits(2, x, y) = 0
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    RotateSmooth = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Saturation(ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim Spd(765) As Long
  Dim R As Long, G As Long, B As Long
  Dim V As Long
  Dim cF As Single

    tBits = mBits

    For x = 0 To 765
        Spd(x) = x \ 3
    Next x

    cF = Factor / 100

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                B = mBits(0, x, y)
                G = mBits(1, x, y)
                R = mBits(2, x, y)
                V = Spd(B + G + R)
                B = B + cF * (B - V)
                G = G + cF * (G - V)
                R = R + cF * (R - V)
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Saturation = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Scanlines(ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim Spd(255) As Long
  Dim cF As Long

    tBits = mBits

    cF = Factor
    For x = 0 To 255
        Spd(x) = x - cF
        If (Spd(x) < 0) Then Spd(x) = 0
    Next x

    GetRgnBoundRect

    For y = yStart To yEnd Step 2
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                tBits(0, x, y) = Spd(mBits(0, x, y))
                tBits(1, x, y) = Spd(mBits(1, x, y))
                tBits(2, x, y) = Spd(mBits(2, x, y))
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Scanlines = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Sharpen(ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim mF As Long, dF As Long

    tBits = mBits

    mF = 24 + (100 - Factor)
    dF = 8 + (100 - Factor)

    GetRgnBoundRect 1

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                B = CLng(mBits(0, x, y - 1)) + mBits(0, x - 1, y) + _
                         mBits(0, x + 1, y) + mBits(0, x, y + 1) + _
                         mBits(0, x + 1, y + 1) + mBits(0, x - 1, y + 1) + _
                         mBits(0, x + 1, y - 1) + mBits(0, x - 1, y - 1)
                B = (mF * (mBits(0, x, y)) - 2 * B) \ dF
                G = CLng(mBits(1, x, y - 1)) + mBits(1, x - 1, y) + _
                         mBits(1, x + 1, y) + mBits(1, x, y + 1) + _
                         mBits(1, x + 1, y + 1) + mBits(1, x - 1, y + 1) + _
                         mBits(1, x + 1, y - 1) + mBits(1, x - 1, y - 1)
                G = (mF * (mBits(1, x, y)) - 2 * G) \ dF
                R = CLng(mBits(2, x, y - 1)) + mBits(2, x - 1, y) + _
                         mBits(2, x + 1, y) + mBits(2, x, y + 1) + _
                         mBits(2, x + 1, y + 1) + mBits(2, x - 1, y + 1) + _
                         mBits(2, x + 1, y - 1) + mBits(2, x - 1, y - 1)
                R = (mF * (mBits(2, x, y)) - 2 * R) \ dF
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Sharpen = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function ShiftRGB(ByVal rFactor As Long, ByVal gFactor As Long, ByVal bFactor As Long) As Byte()

  Dim tBits() As Byte
  Dim rSpd(255) As Long
  Dim gSpd(255) As Long
  Dim bSpd(255) As Long

    tBits = mBits

    For x = 0 To 255
        bSpd(x) = x + bFactor
        If (bSpd(x) > 255) Then bSpd(x) = 255
        If (bSpd(x) < 0) Then bSpd(x) = 0
        gSpd(x) = x + gFactor
        If (gSpd(x) > 255) Then gSpd(x) = 255
        If (gSpd(x) < 0) Then gSpd(x) = 0
        rSpd(x) = x + rFactor
        If (rSpd(x) > 255) Then rSpd(x) = 255
        If (rSpd(x) < 0) Then rSpd(x) = 0
    Next x

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                tBits(0, x, y) = bSpd(mBits(0, x, y))
                tBits(1, x, y) = gSpd(mBits(1, x, y))
                tBits(2, x, y) = rSpd(mBits(2, x, y))
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    ShiftRGB = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Shift(ByVal ColorFrom As Long, ByVal Inc As Long, ByVal Tolerance As Long) As Byte()

  Dim tBits() As Byte
  Dim Spd(255) As Long
  Dim Rf As Long, Gf As Long, Bf As Long
  Dim tF As Long
  Dim distance As Long, ratio As Single

    tBits = mBits

    Rf = (ColorFrom And &HFF&)
    Gf = (ColorFrom And &HFF00&) \ 256
    Bf = (ColorFrom And &HFF0000) \ 65536
    tF = (127.5 * Tolerance / 100) ^ 2 + 1

    For x = 0 To 255
        Spd(x) = x + Inc
        If Spd(x) < 0 Then Spd(x) = 0
        If Spd(x) > 255 Then Spd(x) = 255
    Next x

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                distance = (mBits(0, x, y) - Bf) * (mBits(0, x, y) - Bf) _
                         + (mBits(1, x, y) - Gf) * (mBits(1, x, y) - Gf) _
                         + (mBits(2, x, y) - Rf) * (mBits(2, x, y) - Rf)
                If (distance < tF) Then
                    ratio = 1 - distance / tF
                    tBits(0, x, y) = mBits(0, x, y) + (Spd(mBits(0, x, y)) - mBits(0, x, y)) * ratio
                    tBits(1, x, y) = mBits(1, x, y) + (Spd(mBits(1, x, y)) - mBits(1, x, y)) * ratio
                    tBits(2, x, y) = mBits(2, x, y) + (Spd(mBits(2, x, y)) - mBits(2, x, y)) * ratio
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Shift = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Sketch(ByVal Factor As Long, Optional ByVal BW As Boolean = 0, Optional BackColor As Long = &HFFFFFF) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim rB As Long, Gb As Long, Bb As Long
  Dim cF As Long
  Dim bwOut As Long

    tBits = mBits

    rB = (BackColor And &HFF&)
    Gb = (BackColor And &HFF00&) \ 256
    Bb = (BackColor And &HFF0000) \ 65536
    cF = Factor

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                B = CLng(mBits(0, x - 1, y - 1)) + mBits(0, x - 1, y) + _
                         mBits(0, x - 1, y + 1) + mBits(0, x, y - 1) + _
                         mBits(0, x, y + 1) + mBits(0, x + 1, y - 1) + _
                         mBits(0, x + 1, y) + mBits(0, x + 1, y + 1)
                G = CLng(mBits(1, x - 1, y - 1)) + mBits(1, x - 1, y) + _
                         mBits(1, x - 1, y + 1) + mBits(1, x, y - 1) + _
                         mBits(1, x, y + 1) + mBits(1, x + 1, y - 1) + _
                         mBits(1, x + 1, y) + mBits(1, x + 1, y + 1)
                R = CLng(mBits(2, x - 1, y - 1)) + mBits(2, x - 1, y) + _
                         mBits(2, x - 1, y + 1) + mBits(2, x, y - 1) + _
                         mBits(2, x, y + 1) + mBits(2, x + 1, y - 1) + _
                         mBits(2, x + 1, y) + mBits(2, x + 1, y + 1)
                B = 8 * mBits(0, x, y) - B + 255
                G = 8 * mBits(1, x, y) - G + 255
                R = 8 * mBits(2, x, y) - R + 255
                bwOut = (CLng(B) + G + R) \ 3
                If (bwOut > cF) Then bwOut = 255 Else If (bwOut <= cF) Then bwOut = 0
                If (BW) Then
                    tBits(0, x, y) = bwOut
                    tBits(1, x, y) = bwOut
                    tBits(2, x, y) = bwOut
                  ElseIf (bwOut = 255) Then
                    tBits(0, x, y) = Bb
                    tBits(1, x, y) = Gb
                    tBits(2, x, y) = rB
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Sketch = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Soften(ByVal Factor As Long, Optional ByVal IsText As Boolean = 0) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim cF As Long, cW As Long

    tBits = mBits

    cF = 101 - Factor
    cW = cF + 4

    GetRgnBoundRect 1

    If (IsText) Then xEnd = xEnd - 2: yEnd = yEnd - 2

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                B = (CLng(cF * mBits(0, x, y)) + _
                    mBits(0, x - 1, y) + mBits(0, x, y - 1) + _
                    mBits(0, x, y + 1) + mBits(0, x + 1, y)) \ cW
                G = (CLng(cF * mBits(1, x, y)) + _
                    mBits(1, x - 1, y) + mBits(1, x, y - 1) + _
                    mBits(1, x, y + 1) + mBits(1, x + 1, y)) \ cW
                R = (CLng(cF * mBits(2, x, y)) + _
                    mBits(2, x - 1, y) + mBits(2, x, y - 1) + _
                    mBits(2, x, y + 1) + mBits(2, x + 1, y)) \ cW
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Soften = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Stretch(ByVal wFactor As Long, ByVal hFactor As Long, Optional ByVal CleanBack As Boolean = -1) As Byte()

  Dim tBits() As Byte
  Dim wSpd() As Long
  Dim hSpd() As Long
  Dim cW As Single, cH As Single
  Dim ox As Long, oy As Long

    tBits = mBits

    ReDim wSpd(biW) As Long
    ReDim hSpd(biH) As Long
    cW = (1 + wFactor / 100)
    cH = (1 + hFactor / 100)

    ox = biW \ 2
    oy = biH \ 2

    For x = 0 To biW
        wSpd(x) = (x - ox) * cW + ox
    Next x
    For y = 0 To biH
        hSpd(y) = (y - oy) * cH + oy
    Next y

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, wSpd(x), hSpd(y))) Then
                tBits(0, x, y) = mBits(0, wSpd(x), hSpd(y))
                tBits(1, x, y) = mBits(1, wSpd(x), hSpd(y))
                tBits(2, x, y) = mBits(2, wSpd(x), hSpd(y))
              Else
                If (CleanBack) Then
                    tBits(0, x, y) = 0
                    tBits(1, x, y) = 0
                    tBits(2, x, y) = 0
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Stretch = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function StretchH(ByVal HOffset As Long, ByVal Saturation As Long, ByVal HStep As Long)

  '// Need mHSLRGB module

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim H As Single, S As Single, L As Single
  Dim cSat As Single
  Dim cStp As Single
  Dim cOff As Single
  
    tBits = mBits

    cSat = Saturation / 100
    cStp = HStep / 10
    cOff = HOffset / 16.67

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                L = (CLng(mBits(0, x, y)) + mBits(1, x, y) + mBits(2, x, y)) / 765
                H = 3 * Sin(cStp * L) + cOff
                If (H < -1) Then H = H + 6 Else If (H > 5) Then H = H - 6
                HSLtoRGB H, cSat, L, R, G, B
                tBits(0, x, y) = B
                tBits(1, x, y) = G
                tBits(2, x, y) = R
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    StretchH = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function StuckiBW(ByVal PalWeight As Long) As Byte()

  Dim tBits() As Byte
  Dim Spd(765) As Long
  Dim R As Long, G As Long, B As Long
  Dim Erro As Long
  Dim VecErro() As Long
  Dim nCol As Long, mCol As Long
  Dim PartErr(1 To 12, -255 To 255) As Long

    tBits = mBits

    For x = 0 To 765
        Spd(x) = x \ 3
    Next x

    For x = -255 To 255
        PartErr(1, x) = (8 * x) \ 42
        PartErr(2, x) = (4 * x) \ 42
        PartErr(3, x) = (2 * x) \ 42
        PartErr(4, x) = (4 * x) \ 42
        PartErr(5, x) = (8 * x) \ 42
        PartErr(6, x) = (4 * x) \ 42
        PartErr(7, x) = (2 * x) \ 42
        PartErr(8, x) = (1 * x) \ 42
        PartErr(9, x) = (2 * x) \ 42
        PartErr(10, x) = (4 * x) \ 42
        PartErr(11, x) = (2 * x) \ 42
        PartErr(12, x) = (1 * x) \ 42
    Next x

    Erro = 0
    ReDim VecErro(1 To 3, biW) As Long
    For x = 0 To biW
        VecErro(1, x) = 0
        VecErro(2, x) = 0
        VecErro(3, x) = 0
    Next x
    For y = 0 To biH
        For x = 0 To biW
            B = mBits(0, x, y)
            G = mBits(1, x, y)
            R = mBits(2, x, y)
            B = Spd(R + G + B)
            mCol = mCol + B
            nCol = nCol + 1
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    RaiseEvent Progress(100)

    mCol = mCol \ nCol

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                B = mBits(0, x, y)
                G = mBits(1, x, y)
                R = mBits(2, x, y)
                B = Spd(R + G + B)
                B = B + (VecErro(1, x) * 10) \ PalWeight
                If (B < 0) Then B = 0
                If (B > 255) Then B = 255
                If (B < mCol) Then nCol = 0 Else nCol = 255
                tBits(0, x, y) = nCol
                tBits(1, x, y) = nCol
                tBits(2, x, y) = nCol
                Erro = (B - nCol)
                If (x < biW - 1) Then VecErro(1, x + 1) = VecErro(1, x + 1) + PartErr(1, Erro)
                If (x < biW - 2) Then VecErro(1, x + 2) = VecErro(1, x + 2) + PartErr(2, Erro)
                If (y < biH - 1) Then
                    If (x > 1) Then VecErro(2, x - 2) = VecErro(2, x - 2) + PartErr(3, Erro)
                    If (x > 0) Then VecErro(2, x - 1) = VecErro(2, x - 1) + PartErr(4, Erro)
                    VecErro(2, x) = VecErro(2, x) + PartErr(5, Erro)
                    If (x < biW - 1) Then VecErro(2, x + 1) = VecErro(2, x + 1) + PartErr(6, Erro)
                    If (x < biW - 2) Then VecErro(2, x + 2) = VecErro(2, x + 2) + PartErr(7, Erro)
                End If
                If (y < biH - 2) Then
                    If (x > 1) Then VecErro(3, x - 2) = VecErro(3, x - 2) + PartErr(8, Erro)
                    If (x > 0) Then VecErro(3, x - 1) = VecErro(3, x - 1) + PartErr(9, Erro)
                    VecErro(3, x) = VecErro(3, x) + PartErr(10, Erro)
                    If (x < biW - 1) Then VecErro(3, x + 1) = VecErro(3, x + 1) + PartErr(11, Erro)
                    If (x < biW - 2) Then VecErro(3, x + 2) = VecErro(3, x + 2) + PartErr(12, Erro)
                End If
            End If
        Next x
        For x = 0 To biW
            VecErro(1, x) = VecErro(2, x)
            VecErro(2, x) = VecErro(3, x)
            VecErro(3, x) = 0
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    StuckiBW = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function SwapRGB(ByVal rBank As Long, ByVal gBank As Long, ByVal bBank As Long, ByVal rInv As Boolean, ByVal gInv As Boolean, ByVal bInv As Boolean, ByVal rKill As Boolean, ByVal gKill As Boolean, ByVal bKill As Boolean) As Byte()

  Dim tBits() As Byte

    tBits = mBits

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                If (bInv) Then
                    tBits(0, x, y) = (255 - mBits(bBank, x, y)) And (Not bKill)
                  Else
                    tBits(0, x, y) = mBits(bBank, x, y) And (Not bKill)
                End If
                If (gInv) Then
                    tBits(1, x, y) = (255 - mBits(gBank, x, y)) And (Not gKill)
                  Else
                    tBits(1, x, y) = mBits(gBank, x, y) And (Not gKill)
                End If
                If (rInv) Then
                    tBits(2, x, y) = (255 - mBits(rBank, x, y)) And (Not rKill)
                  Else
                    tBits(2, x, y) = mBits(rBank, x, y) And (Not rKill)
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    SwapRGB = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Texturize(ByRef mskBits() As Byte) As Byte()

  Dim tBits() As Byte
  Dim R As Long, G As Long, B As Long
  Dim mx As Long, my As Long
  Dim mWSize As Long, mHSize As Long
  Dim cB As Single, cP As Single

    tBits = mBits

    mWSize = UBound(mskBits, 2) - 2
    mHSize = UBound(mskBits, 3) - 2

    GetRgnBoundRect

    For y = yStart To yEnd Step mHSize
        For x = xStart To xEnd Step mWSize
            For my = 0 To mHSize
                yIn = y + my
                For mx = 0 To mWSize
                    xIn = x + mx
                    If (PtInRegion(biRgn, xIn, yIn)) Then
                        '// R=B=G (based on embossed/engraved 'map')
                        cB = mskBits(0, mx, my) / 128
                        B = cB * mBits(0, xIn, yIn)
                        G = cB * mBits(1, xIn, yIn)
                        R = cB * mBits(2, xIn, yIn)
                        If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                        If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                        If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                        tBits(0, xIn, yIn) = B
                        tBits(1, xIn, yIn) = G
                        tBits(2, xIn, yIn) = R
                    End If
                Next mx
            Next my
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Texturize = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Threshold(ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim cF As Long

    tBits = mBits

    cF = Factor

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                If ((CLng(mBits(0, x, y)) + mBits(1, x, y) + mBits(2, x, y)) \ 3 <= cF) Then
                    tBits(0, x, y) = 0
                    tBits(1, x, y) = 0
                    tBits(2, x, y) = 0
                  Else
                    tBits(0, x, y) = 255
                    tBits(1, x, y) = 255
                    tBits(2, x, y) = 255
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Threshold = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Tile(ByVal Factor As Long) As Byte()

  Dim tBits() As Byte
  Dim cW As Long, cH As Long
  Dim sW As Single, sH As Single
  Dim i As Long, j As Long

    tBits = mBits

    cW = biW / Factor
    cH = biH / Factor
    If (cW < 1) Then cW = 1
    If (cH < 1) Then cH = 1
    sW = biW / cW
    sH = biH / cH

    GetRgnBoundRect

    For y = yStart To yEnd Step cH
        For x = xStart To xEnd Step cW
            For j = y To y + cH
                yIn = (j - y) * sH
                For i = x To x + cW
                    xIn = (i - x) * sW
                    If (PtInRegion(biRgn, i, j)) Then
                        tBits(0, i, j) = mBits(0, xIn, yIn)
                        tBits(1, i, j) = mBits(1, xIn, yIn)
                        tBits(2, i, j) = mBits(2, xIn, yIn)
                    End If
                Next i
            Next j
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Tile = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function TileMask(ByRef mskBits() As Byte, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0, Optional ByVal Tiled As Boolean = 0, Optional ByVal Transparency As Long = 50, Optional ByVal MaskColor As Long = -1) As Byte()

  Dim tBits() As Byte
  Dim mw As Long, mh As Long
  Dim mx As Long, my As Long
  Dim sx As Long, sy As Long
  Dim Rm As Long, Gm As Long, Bm As Long
  Dim ci1 As Single, ci2 As Single

    tBits = mBits

    mw = UBound(mskBits, 2) - 2
    mh = UBound(mskBits, 3) - 2

    If (MaskColor <> -1) Then
        Rm = (MaskColor And &HFF&)
        Gm = (MaskColor And &HFF00&) \ 256
        Bm = (MaskColor And &HFF0000) \ 65536
      Else
        Rm = -1
    End If

    If (Tiled) Then
        sx = mw + 1
        sy = mh + 1
      Else
        sx = biW - xOffset + 1
        sy = biH - yOffset + 1
    End If

    ci1 = Transparency / 100
    ci2 = 1 - ci1

    For y = yOffset To biH Step sy
        For x = xOffset To biW Step sx
            For my = y To y + mh
                yIn = my - y
                For mx = x To x + mw
                    xIn = mx - x
                    If (PtInRegion(biRgn, mx, my)) Then
                        If (mskBits(2, xIn, yIn) = Rm And mskBits(1, xIn, yIn) = Gm And mskBits(0, xIn, yIn) = Bm) Then
                          Else
                            tBits(0, mx, my) = ci1 * mBits(0, mx, my) + ci2 * mskBits(0, xIn, yIn)
                            tBits(1, mx, my) = ci1 * mBits(1, mx, my) + ci2 * mskBits(1, xIn, yIn)
                            tBits(2, mx, my) = ci1 * mBits(2, mx, my) + ci2 * mskBits(2, xIn, yIn)
                        End If
                    End If
                Next mx
            Next my
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress((y - yOffset) * 100 \ (biH - yOffset))
        End If
    Next y
    TileMask = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Twirl(ByVal Factor As Long, ByVal OffsetDeg As Long, Optional ByVal CleanBack As Boolean = -1) As Byte()

  Dim tBits() As Byte
  Dim d As Long, theta As Single, cF As Single, cOD As Single
  Dim yInPow2 As Long
  Dim ox As Long, oy As Long
  Dim rx As Long, ry As Long

    tBits = mBits

    ox = biW \ 2
    oy = biH \ 2

    cF = (Factor / 10) / ox
    cOD = OffsetDeg * PIdiv180

    GetRgnBoundRect

    For y = yStart To yEnd
        yIn = y - oy
        yInPow2 = yIn * yIn
        For x = xStart To xEnd
            xIn = x - ox
            d = Sqr(xIn * xIn + yInPow2)
            If (d = 0) Then
                rx = ox
                ry = oy
              Else
                theta = (ArcTan(xIn, yIn)) - d * cF
                rx = d * Cos(theta + cOD) + ox
                ry = d * Sin(theta + cOD) + oy
            End If
            If (PtInRegion(biRgn, rx, ry)) Then
                tBits(0, x, y) = mBits(0, rx, ry)
                tBits(1, x, y) = mBits(1, rx, ry)
                tBits(2, x, y) = mBits(2, rx, ry)
              Else
                If (CleanBack) Then
                    tBits(0, x, y) = 0
                    tBits(1, x, y) = 0
                    tBits(2, x, y) = 0
                End If
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Twirl = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

Public Function Waves(ByVal fFactor As Long, ByVal AFactor As Long, Optional ByVal Start As Long = 0) As Byte()

  Dim tBits() As Byte
  Dim cA As Single
  Dim cF As Single
  Dim cS As Long

    ReDim Spd(biH) As Long

    tBits = mBits

    cF = fFactor * 1
    cA = AFactor * 0.25
    cS = Start

    For y = 0 To biH
        Spd(y) = y + cA * Sin((cS + cF * y) * PIdiv180)
        If (Spd(y) < 0) Then Spd(y) = 0
        If (Spd(y) > biH) Then Spd(y) = biH
    Next y

    GetRgnBoundRect

    For y = yStart To yEnd
        For x = xStart To xEnd
            If (PtInRegion(biRgn, x, y)) Then
                tBits(0, x, y) = mBits(0, x, Spd(y))
                tBits(1, x, y) = mBits(1, x, Spd(y))
                tBits(2, x, y) = mBits(2, x, Spd(y))
            End If
        Next x
        If (timeGetTime - t > PROGRESS_DELAY) Then
            t = timeGetTime
            RaiseEvent Progress(y * 100 \ biH)
        End If
    Next y
    Waves = tBits
    Erase tBits

    RaiseEvent Progress(100)

End Function

'//
'// Functions 2: [L min/med/max]
'//

Public Function GetLMax() As Long

  Dim Lmax As Long
  Dim LCur As Long

    Lmax = 0
    For y = 0 To biH
        For x = 0 To biW
            LCur = (CLng(mBits(0, x, y)) + mBits(1, x, y) + mBits(2, x, y)) \ 3
            If (LCur > Lmax) Then Lmax = LCur
        Next x
    Next y
    GetLMax = Lmax

End Function

Public Function GetLMin() As Long

  Dim Lmin As Long
  Dim LCur As Long

    Lmin = 255
    For y = 0 To biH
        For x = 0 To biW
            LCur = (CLng(mBits(0, x, y)) + mBits(1, x, y) + mBits(2, x, y)) \ 3
            If (LCur < Lmin) Then Lmin = LCur
        Next x
    Next y
    GetLMin = Lmin

End Function

Public Function GetLMed() As Long

  Dim Spd(765) As Long
  Dim R As Long, G As Long, B As Long
  Dim mCol As Long, nCol As Long

    For x = 0 To 765
        Spd(x) = x \ 3
    Next x

    mCol = 0
    For y = 0 To biH
        For x = 0 To biW
            B = mBits(0, x, y)
            G = mBits(1, x, y)
            R = mBits(2, x, y)
            mCol = mCol + Spd(R + G + B)
            nCol = nCol + 1
        Next x
    Next y
    mCol = mCol \ nCol

    GetLMed = mCol

End Function

'//
'// ArcTan...
'//

Private Function ArcTan(ByVal xt As Long, ByVal yt As Long) As Single

    If (xt = 0) Then
        If (yt > 0) Then
            ArcTan = PI / 2
          Else
            ArcTan = -(PI / 2)
        End If
      Else
        If (xt < 0) Then
            ArcTan = PI + Atn(yt / xt)
          Else
            ArcTan = Atn(yt / xt)
        End If
    End If

End Function

'//
'// Get minimum rect. region
'//

Private Sub GetRgnBoundRect(Optional ByVal FilterMargin As Long = 0)

  Dim BRct As RECT

    GetRgnBox biRgn, BRct

    With BRct
        xStart = .Left + FilterMargin
        xEnd = .Right - FilterMargin - 1
        yStart = .Top + FilterMargin
        yEnd = .Bottom - FilterMargin - 1
    End With

End Sub

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 11:51:29) 99 + 3065 = 3164 Lines



