Attribute VB_Name = "mSetup"
'//
'// mSetup module
'// (To do... -> INI format) ?
'//

Option Explicit

Private Type tMainInfo
    Width As Integer
    Height As Integer
    iPMaximized As Boolean
    ImagesMaximized As Boolean
    ImagesPath As String * 256
    PreserveLastFolder As Boolean
    AskToSave As Boolean
    StartZoom As Integer
    Recent(3) As String * 256
    ShowToolbar As Boolean
    ShowPanoramicView As Boolean
    ShowPalette As Boolean
    ShowTools As Boolean
    ShowToolOptions As Boolean
    ShowPasteBox As Boolean
End Type

Private Type tPaletteInfo
    Hue As Single
    ColorFore As Long
    ColorBack As Long
    Colors(15) As Long
    LinkCurrent As Boolean
End Type

Private Type tToolInfo
    Key As String * 25
End Type

Private Type tSelection
    Mode As Integer
    ShapeType As Integer
    Pressure As Integer
    MoveSelection As Boolean
End Type

Private Type tBrushInfo
    Mode As Integer
    Size As Integer
    Pressure As Integer
    Continuous As Boolean
End Type

Private Type tAirbrushInfo
    Size As Integer
    Pressure As Integer
    Quantity As Integer
    Definition As Integer
End Type

Private Type tFillInfo
    Mode As Integer
    Tolerance As Integer
    Pressure As Integer
End Type

Private Type tShapeInfo
    Shape As Integer
    Pressure As Integer
    Sides As Integer
End Type

Private Type tStampInfo
    MaskColor As Long
    UseMaskColor As Boolean
    StampMask As Boolean
    ClearBack As Boolean
    Pressure As Integer
End Type

Private Type tTextInfo
    Text As String * 256
    FontName As String * 256
    FontSize As Integer
    FontBold As Boolean
    FontItalic As Boolean
    FontUnderline As Boolean
    Alignment As Integer
    Pressure As Integer
    Rotate90 As Boolean
    FlipH As Boolean
    FlipV As Boolean
    SmoothFactor As Integer
End Type

Private Type tGradInfo
    Mode As Long
    Colors(1) As Long
    Transparent(1) As Boolean
    Pressure As Integer
    Frequency As Integer
End Type

Private Type tEffectInfo
    Effects(11) As Boolean
    Amount As Integer
End Type

Public Type tiPSetup
    Main As tMainInfo
    Palette As tPaletteInfo
    Tool As tToolInfo
    Selection As tSelection
    Brush As tBrushInfo
    Airbrush As tAirbrushInfo
    Fill As tFillInfo
    Shape As tShapeInfo
    Stamp As tStampInfo
    Text As tTextInfo
    Gradient As tGradInfo
    Effect As tEffectInfo
End Type

Public iPStp As tiPSetup

'// Load Setup data
Public Sub LoadSetup()

  Dim i As Integer
  Dim ff As Integer
  Dim firstGo As Boolean

    If (Not FileFound(App.Path & "\iP.dat")) Then
        firstGo = -1
    Else
        If (Len(App.Path & "\iP.dat") = 0) Then
            MsgBox "Configuration data file is corrupted" & vbCrLf & "Values will be initilized", vbExclamation, "iP"
            Kill App.Path & "\iP.dat"
            firstGo = -1
        End If
    End If
    
    ff = FreeFile
    Open App.Path & "\iP.dat" For Binary As #ff
      Get #ff, , iPStp
    Close #ff
    ff = 0

    If (firstGo) Then
        '// Main window/images/boxes
        With iPStp.Main
            .Width = 640 * Screen.TwipsPerPixelX
            .Height = 480 * Screen.TwipsPerPixelY
            .iPMaximized = 0
            .ImagesMaximized = 0
            .StartZoom = 1
            .ImagesPath = App.Path & "\Images\"
            .ShowToolbar = 0
            .ShowPanoramicView = 0
            .ShowPalette = 0
            .ShowTools = 0
            .ShowToolOptions = 0
            .ShowPasteBox = 0
            .PreserveLastFolder = -1
            .AskToSave = -1
        End With
        '// Palette
        With iPStp.Palette
            .Hue = 0
            .ColorBack = &H0
            .ColorFore = &HFFFFFF
            For i = 0 To 15
                .Colors(i) = RGB(i * 17, i * 17, i * 17)
            Next i
            .LinkCurrent = -1
        End With
        '// Tools
        With iPStp.Selection
            .Mode = 2
            .ShapeType = 0
            .Pressure = 50
            .MoveSelection = -1
        End With
        With iPStp.Brush
            .Mode = 0
            .Size = 25
            .Pressure = 25
            .Continuous = -1
        End With
        With iPStp.Airbrush
            .Size = 75
            .Pressure = 25
            .Quantity = 100
            .Definition = 0
        End With
        With iPStp.Fill
            .Mode = 0
            .Tolerance = 25
            .Pressure = 100
        End With
        With iPStp.Shape
            .Shape = 0
            .Pressure = 100
            .Sides = 3
        End With
        With iPStp.Stamp
            .MaskColor = &H0
            .UseMaskColor = 0
            .StampMask = -1
            .ClearBack = 0
            .Pressure = 50
        End With
        With iPStp.Text
            .Text = "iP 1.0"
            .FontName = "Arial"
            .FontSize = 20
            .FontBold = -1
            .FontItalic = -1
            .FontUnderline = 0
            .Alignment = 2
            .Pressure = 100
            .Rotate90 = 0
            .FlipH = 0
            .FlipV = 0
            .SmoothFactor = 0
        End With
        With iPStp.Gradient
            .Mode = 0
            .Colors(0) = &HFFFFFF
            .Colors(1) = &H0
            .Transparent(0) = 0
            .Transparent(1) = 0
            .Pressure = 50
            .Frequency = 1
        End With
        With iPStp.Effect
            .Amount = 10
            For i = 0 To 7
                .Effects(i) = -1
            Next i
        End With
    End If
    
    '// Flag (Get current bits on image change)
    fIDLast = -1
    
    '// Fill tool values
   '// Load fToolOptions

End Sub

'// Save Setup data
Public Sub SaveSetup()

  Dim ff As Integer

    ff = FreeFile
    Open App.Path & "\iP.dat" For Binary As #ff
      Put #ff, , iPStp
    Close #ff
    ff = 0

End Sub

'// Add to recent files
Public Sub AddRecent(ByVal FileName As String)

  Dim i As Integer

    With iPStp.Main
        For i = 0 To UBound(.Recent)
            If (FileName = RTrim(.Recent(i))) Then
                Exit Sub
            End If
            If (.Recent(i) = Space(256)) Then
                .Recent(i) = FileName
                Exit Sub
            End If
        Next i
        For i = UBound(.Recent) To 1 Step -1
            .Recent(i) = .Recent(i - 1)
        Next i
        .Recent(0) = FileName
    End With

End Sub

Public Function CompactPath(ByVal hdc As Long, ByVal FullPath As String, ByVal Width As Long) As String

  '// from
  '//  KPD-Team 2000
  '//  URL: http://www.allapi.net/
  '//  E-Mail: KPDTeam@Allapi.net

  Dim ZeroPos As Long

    '// Compact
    PathCompactPath hdc, FullPath, Width

    '// Remove all trailing Chr$(0)'s
    ZeroPos = InStr(1, FullPath, Chr$(0))
    If (ZeroPos > 0) Then
        CompactPath = Left$(FullPath, ZeroPos - 1)
      Else
        CompactPath = FullPath
    End If

End Function

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 11:44:33) 119 + 169 = 288 Lines
