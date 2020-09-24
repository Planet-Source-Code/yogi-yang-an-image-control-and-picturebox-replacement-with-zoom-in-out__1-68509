VERSION 5.00
Begin VB.UserControl epCmDlg 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   375
   Begin VB.Image imgLogo 
      Height          =   240
      Left            =   75
      Picture         =   "epCmDlg.ctx":0000
      Stretch         =   -1  'True
      Top             =   75
      Width           =   240
   End
End
Attribute VB_Name = "epCmDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' Win32 Declarations for the Common Dialog
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDlg) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long

' Win32 Declarations for the ShowFont function
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const LF_FACESIZE = 32
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long
        hDC As Long
        lpLogFont As Long
        iPointSize As Long
        flags As Long
        rgbColors As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        hInstance As Long
        lpszStyle As String
        nFontType As Integer
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long
        nSizeMax As Long
End Type

Private Type ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Type PRINTDLG_TYPE
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hDC As Long
        flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
        extra As String * 100
End Type

Private Type DEVMODE_TYPE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type PageSetupDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        flags As Long
        ptPaperSize As POINTAPI
        rtMinMargin As RECT
        rtMargin As RECT
        hInstance As Long
        lCustData As Long
        lpfnPageSetupHook As Long
        lpfnPagePaintHook As Long
        lpPageSetupTemplateName As String
        hPageSetupTemplate As Long
End Type

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 31
End Type

' Constants for the common dialog
Private Const OFN_ALLOWMULTISELECT = &H200  'Allow multi select (Open Dialog)
Private Const OFN_EXPLORER = &H80000        'Set windows style explorer
Private Const OFN_FILEMUSTEXIST = &H1000    'File must exist
Private Const OFN_HIDEREADONLY = &H4        'Hide read-only check box (Open Dialog)
Private Const OFN_OVERWRITEPROMPT = &H2     'Promt beafore overwritning file (Save Dialog)
Private Const OFN_PATHMUSTEXIST = &H800     'Path must exist
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const DEFAULT_CHARSET = 1
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_QUALITY = 0
Private Const FW_BOLD = 700
Private Const FF_ROMAN = 16      '  Variable stroke width, serifed.
Private Const FW_NORMAL = 400
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const REGULAR_FONTTYPE = &H400
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&

' Constants for the GlobalAllocate
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40

Private Const MAX_PATH = 260 'Constant for maximum path

Public cFileName As Collection   'Filename collection
Public cFileTitle As Collection  'Filetitle collection

' Default Property Values:
Const m_def_CancelError = 0
Const m_def_Filename = ""
Const m_def_DialogTitle = ""
Const m_def_InitialDir = ""
Const m_def_Filter = ""
Const m_def_FilterIndex = 1
Const m_def_MultiSelect = 0
Const m_def_FontName = "Arial"
Const m_def_FontSize = 10
Const m_def_FontColor = 0
Const m_def_FontBold = 0
Const m_def_FontItalic = 0
Const m_def_FontUnderline = 0
Const m_def_FontStrikeThru = 0

' Property Variables:
Dim m_CancelError As Boolean
Dim m_Filename As String
Dim m_DialogTitle As String
Dim m_InitialDir As String
Dim m_Filter As String
Dim m_FilterIndex As Integer
Dim m_MultiSelect As Boolean
Dim m_FontName As String
Dim m_FontSize As Integer
Dim m_FontColor As Long
Dim m_FontBold As Boolean
Dim m_FontItalic As Boolean
Dim m_FontUnderline As Boolean
Dim m_FontStrikeThru As Boolean

'***** CANCEL ERROR
Public Property Get CancelError() As Boolean
    CancelError = m_CancelError
End Property
Public Property Let CancelError(ByVal New_CancelError As Boolean)
    m_CancelError = New_CancelError
    PropertyChanged "CancelError"
End Property
'***** MULTI SELECT
Public Property Get MultiSelect() As Boolean
    MultiSelect = m_MultiSelect
End Property
Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    m_MultiSelect = New_MultiSelect
    PropertyChanged "MultiSelect"
End Property
'***** DEFAULT FILENAME
Public Property Get DefaultFilename() As String
Attribute DefaultFilename.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    DefaultFilename = m_Filename
End Property
Public Property Let DefaultFilename(ByVal New_Filename As String)
    m_Filename = New_Filename
    PropertyChanged "DefaultFilename"
End Property
'***** DIALOG TITLE
Public Property Get DialogTitle() As String
Attribute DialogTitle.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    DialogTitle = m_DialogTitle
End Property
Public Property Let DialogTitle(ByVal New_DialogTitle As String)
    m_DialogTitle = New_DialogTitle
    PropertyChanged "DialogTitle"
End Property
'***** INITIAL DIRECTORY
Public Property Get InitialDir() As String
    InitialDir = m_InitialDir
End Property
Public Property Let InitialDir(ByVal New_InitialDir As String)
    m_InitialDir = New_InitialDir
    PropertyChanged "InitialDir"
End Property
'***** FILTER
Public Property Get Filter() As String
    Filter = m_Filter
End Property
Public Property Let Filter(ByVal New_Filter As String)
    m_Filter = New_Filter
    PropertyChanged "Filter"
End Property
'***** FILTER INDEX
Public Property Get FilterIndex() As Integer
    FilterIndex = m_FilterIndex
End Property
Public Property Let FilterIndex(ByVal New_FilterIndex As Integer)
    m_FilterIndex = New_FilterIndex
    PropertyChanged "FilterIndex"
End Property
'***** FONT NAME
Public Property Get FontName() As String
    FontName = m_FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
End Property
'***** FONT SIZE
Public Property Get FontSize() As Integer
    FontSize = m_FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Integer)
    m_FontSize = New_FontSize
End Property
'***** FONT COLOR
Public Property Get FontColor() As Long
    FontColor = m_FontColor
End Property
Public Property Let FontColor(ByVal New_FontColor As Long)
    m_FontColor = New_FontColor
End Property
'***** FONT BOLD
Public Property Get FontBold() As Boolean
    FontBold = m_FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    m_FontBold = New_FontBold
End Property
'***** FONT ITALIC
Public Property Get FontItalic() As Boolean
    FontItalic = m_FontItalic
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    m_FontItalic = New_FontItalic
End Property
'***** FONT UNDERLINE
Public Property Get FontUnderline() As Boolean
    FontUnderline = m_FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    m_FontUnderline = New_FontUnderline
End Property
'***** FONT STRIKETHRU
Public Property Get FontStrikeThru() As Boolean
    FontStrikeThru = m_FontStrikeThru
End Property
Public Property Let FontStrikeThru(ByVal New_FontStrikeThru As Boolean)
    m_FontStrikeThru = New_FontStrikeThru
End Property

Private Sub UserControl_Initialize()
    UserControl_Resize
End Sub
' Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_CancelError = m_def_CancelError
    m_Filename = m_def_Filename
    m_DialogTitle = m_def_DialogTitle
    m_InitialDir = m_def_InitialDir
    m_Filter = m_def_Filter
    m_FilterIndex = m_def_FilterIndex
    m_MultiSelect = m_def_MultiSelect
    m_FontName = m_def_FontName
    m_FontSize = m_def_FontSize
    m_FontColor = m_def_FontColor
    m_FontBold = m_def_FontBold
    m_FontItalic = m_def_FontItalic
    m_FontUnderline = m_def_FontUnderline
    m_FontStrikeThru = m_def_FontStrikeThru
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_CancelError = PropBag.ReadProperty("CancelError", m_def_CancelError)
    m_Filename = PropBag.ReadProperty("DefaultFilename", m_def_Filename)
    m_DialogTitle = PropBag.ReadProperty("DialogTitle", m_def_DialogTitle)
    m_InitialDir = PropBag.ReadProperty("InitialDir", m_def_InitialDir)
    m_Filter = PropBag.ReadProperty("Filter", m_def_Filter)
    m_FilterIndex = PropBag.ReadProperty("FilterIndex", m_def_FilterIndex)
    m_MultiSelect = PropBag.ReadProperty("MultiSelect", m_def_MultiSelect)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 360
    UserControl.Width = 375
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CancelError", m_CancelError, m_def_CancelError)
    Call PropBag.WriteProperty("DefaultFilename", m_Filename, m_def_Filename)
    Call PropBag.WriteProperty("DialogTitle", m_DialogTitle, m_def_DialogTitle)
    Call PropBag.WriteProperty("InitialDir", m_InitialDir, m_def_InitialDir)
    Call PropBag.WriteProperty("Filter", m_Filter, m_def_Filter)
    Call PropBag.WriteProperty("FilterIndex", m_FilterIndex, m_def_FilterIndex)
    Call PropBag.WriteProperty("MultiSelect", m_MultiSelect, m_def_MultiSelect)
End Sub

Public Function ShowOpen()
    '** Description:
    '** Calls open dialog without OCX
    Dim epOFN As OPENFILENAME
    Dim lngRet As Long
    With epOFN
    
        If MultiSelect Then 'If Multi Select then
            .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
            .lpstrFile = DefaultFilename & Space(9999 - Len(DefaultFilename)) & vbNullChar
            .lpstrFileTitle = Space(9999) & vbNullChar
        Else
            .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
            .lpstrFile = DefaultFilename & String(MAX_PATH - Len(DefaultFilename), 0) & vbNullChar
            .lpstrFileTitle = String(MAX_PATH, 0) & vbNullChar
        End If

        .hwndOwner = UserControl.ContainerHwnd 'Handle to window
        .lpstrFilter = SetFilter(Filter) & vbNullChar 'File filter
        .lpstrInitialDir = InitialDir & vbNullChar 'Initial directory
        .lpstrTitle = DialogTitle & vbNullChar 'Dialog title
        .lStructSize = Len(epOFN) 'Structure size in bytes
        .nFilterIndex = FilterIndex 'Filter index
        .nMaxFile = Len(.lpstrFile) 'Maximum file length
        .nMaxFileTitle = Len(.lpstrFileTitle) 'Maximum file title length
    End With
    
    lngRet = GetOpenFileName(epOFN) 'Call open dialog
    
    If lngRet <> 0 Then 'If there are no errors continue with opening file
        ParseFileName epOFN.lpstrFile
    Else
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If
End Function

Public Function ShowSave()
    '** Description:
    '** Calls save dialog without OCX
    Dim epOFN As OPENFILENAME
    Dim lngRet As Long
    With epOFN
        .hwndOwner = UserControl.ContainerHwnd 'Handle to window
        .flags = OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        .lpstrFile = DefaultFilename & String(MAX_PATH - Len(DefaultFilename), 0) & vbNullChar
        .lpstrFileTitle = String(MAX_PATH, 0) & vbNullChar
        .lpstrFilter = SetFilter(Filter) & vbNullChar 'File filter
        .lpstrInitialDir = InitialDir & vbNullChar 'Initial directory
        .lpstrTitle = DialogTitle & vbNullChar 'Dialog title
        .lStructSize = Len(epOFN) 'Structure size in bytes
        .nFilterIndex = FilterIndex 'Filter index
        .nMaxFile = Len(.lpstrFile) 'Maximum file length
        .nMaxFileTitle = Len(.lpstrFileTitle) 'Maximum file title length
    End With
    
    lngRet = GetSaveFileName(epOFN) 'Call save dialog
    
    If lngRet <> 0 Then 'If there are no errors continue with saving file
        ParseFileName epOFN.lpstrFile
    Else
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If
End Function

Public Function ShowFont()
    '** Description:
    '** Call font dialog without OCX
    Dim CF As CHOOSEFONT
    Dim LF As LOGFONT
    Dim lMemHandle As Long
    Dim lLogFont As Long
    Dim lngRet As Long
    
    With LF
        .lfCharSet = DEFAULT_CHARSET 'Default character set
        .lfClipPrecision = CLIP_DEFAULT_PRECIS 'Clipping precision
        .lfFaceName = "Arial" & vbNullChar 'Font name
        .lfHeight = 13 'Height
        .lfOutPrecision = OUT_DEFAULT_PRECIS 'Precision mapping
        .lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN 'Default pitch
        .lfQuality = DEFAULT_QUALITY 'Default quality
        .lfWeight = FW_NORMAL 'Regular font type
    End With
    
    ' Create the memory block
    lMemHandle = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(LF))
    lLogFont = GlobalLock(lMemHandle)
    CopyMemory ByVal lLogFont, LF, Len(LF)
        
    With CF
        .flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
        .hDC = Printer.hDC 'Device context of default printer
        .hwndOwner = UserControl.ContainerHwnd 'Handle to window
        .iPointSize = 120 'Set font size to 12 size
        .lpLogFont = lLogFont 'Log font
        .lStructSize = Len(CF) 'Size of structure in bytes
        .nFontType = REGULAR_FONTTYPE 'Regular font type
        .nSizeMax = 72 'Maximum font size
        .nSizeMin = 10 'Minimum font size
        .rgbColors = RGB(0, 0, 0) 'Font color
    End With
    
    lngRet = CHOOSEFONT(CF) 'Call font dialog
    If lngRet <> 0 Then 'If there are no errors continue with font
        CopyMemory LF, ByVal lLogFont, Len(LF)

        FontName = Left(LF.lfFaceName, InStr(LF.lfFaceName, vbNullChar) - 1)
        FontSize = CF.iPointSize / 10
        FontColor = CF.rgbColors
        If LF.lfWeight = FW_NORMAL Then
            FontBold = False
            FontItalic = False
            FontUnderline = False
            FontStrikeThru = False
        Else
            If LF.lfWeight = FW_BOLD Then FontBold = True
            If LF.lfItalic <> 0 Then FontItalic = True
            If LF.lfUnderline <> 0 Then FontUnderline = True
            If LF.lfStrikeOut <> 0 Then FontStrikeThru = True
        End If
    Else
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If
    
    ' Unlock and free the memory block
    ' Note this must be done
    GlobalUnlock lMemHandle
    GlobalFree lMemHandle
End Function

Public Function ShowColor()
    '** Description:
    '** Call color dialog without OCX
    Dim epCC As ChooseColor
    Dim lngRet As Long
    Dim CusCol(0 To 16) As Long
    Dim I As Integer
    
    ' Fills custom colors with white
    For I = 0 To 15
        CusCol(I) = vbWhite
    Next
    
    With epCC
        .hwndOwner = UserControl.ContainerHwnd 'Handle to window
        .lStructSize = Len(epCC) 'Structure size in bytes
        .lpCustColors = VarPtr(CusCol(0)) 'Custom colors
        .rgbResult = 0 'RGB result
    End With
    
    lngRet = ChooseColor(epCC) 'Call color dialog
    If lngRet <> 0 Then 'If there are no errors continue with color
        ShowColor = epCC.rgbResult
    Else
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If
End Function

Public Function ShowPageSetup()
    '** Description:
    '** Call page setup dialog without OCX
    Dim epPSD As PageSetupDlg
    Dim lngRet As Long
    
    epPSD.lStructSize = Len(epPSD) 'Structure size in bytes
    epPSD.hwndOwner = UserControl.ContainerHwnd 'Handle to window
    
    lngRet = PageSetupDlg(epPSD) 'Call page setup dialog
    If lngRet <> 0 Then 'If there are no errors continue
        '
    Else
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If
End Function

Public Function ShowPrinter()
    '** Description:
    '** Call printer dialog without OCX
    '**
    '** Note:
    '** This is not my function it's from KPD-Team 1998 URL: http://www.allapi.net
    '** and i have modified it a little
    '-> Code by Donald Grover
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures

    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hwndOwner = UserControl.ContainerHwnd 'Handle to window

    On Error Resume Next
    'Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) <> 0 Then

        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                    'set printer toolbar name at this point
                End If
            Next
        End If

        On Error Resume Next
        'Set printer object properties according to selections made
        'by user
        Printer.Copies = DevMode.dmCopies
        Printer.Duplex = DevMode.dmDuplex
        Printer.Orientation = DevMode.dmOrientation
        Printer.PaperSize = DevMode.dmPaperSize
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource
        On Error GoTo 0
    Else
        If CancelError Then
            ' For this to work you must check in Tools\Options\General
            ' Break on Unhandled errors if it isn't already checked
            Err.Raise 32755, App.EXEName, "Cancel was selected.", "cmdlg98.chm", 32755
        End If
    End If
End Function

Private Function ParseFileName(sFileName As String)
    '** Description:
    '** Remove null chars from filename and parse multi filename
    '**
    '** Syntax:
    '** szFilename = ParseFileName(strFilename)
    '**
    '** Example:
    '** szFilename = ParseFileName("C:\Autoexec.bat||")
    Dim I As Long
    Dim sPath As String
    Dim sFiles() As String
    Dim Pos As Integer
    Dim sFile As String
    Dim sFileTitle As String
    
    ' Create new collections
    Set cFileName = New Collection
    Set cFileTitle = New Collection
    ' Found position of two last null chars
    Pos = InStr(sFileName, vbNullChar & vbNullChar)
    ' Remove from filename last two chars
    sFile = Left(sFileName, Pos - 1)
    
    ' Check to see if filename is single or multi
    If InStr(1, sFile, vbNullChar) <> 0 Then
    ' Multi file
        sFile = Left(sFileName, Pos) & vbNullChar 'Add null char at end of filename
        sPath = Left(sFileName, InStr(1, sFileName, Chr(0)) - 1) 'Get file path
        sFiles = Split(sFile, Chr(0)) 'Split file where is nullchar
        
        ' Add all filenames to collection
        For I = LBound(sFiles) To UBound(sFiles) - 2
            ' If path doesent contain separator then add it
            If Right(sPath, 1) = "\" Then
                cFileName.Add sPath & sFiles(I)
            Else
                cFileName.Add sPath & "\" & sFiles(I)
            End If
            ' Add file title
            cFileTitle.Add sFiles(I)
            ' Remove first item from collections
            If I = 1 Then cFileName.Remove 1: cFileTitle.Remove 1
        Next
    Else ' Single file
        'Add file name to collection
        cFileName.Add sFile
        ' Add file title
        cFileTitle.Add Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
    End If
End Function

Private Function SetFilter(sFlt As String) As String
    '** Description:
    '** Replace "|" with Null Character
    '**
    '** Syntax:
    '** szFilter = SetFilter(strFilter)
    '**
    '** Example:
    '** szFilter = SetFilter("Text Files (*.txt)|*.txt|All Files |*.*|")
    Dim sLen As Long
    Dim Pos As Long

    sLen = Len(sFlt) 'Get filter length
    Pos = InStr(1, sFlt, "|") 'Find first position of "|"

    ' Loop while Pos > 0
    While Pos > 0
        ' Replace "|" with null char
        sFlt = Left(sFlt, Pos - 1) & vbNullChar & Mid(sFlt, Pos + 1, sLen - Pos)
        ' Find next position of "|"
        Pos = InStr(Pos + 1, sFlt, "|")
    Wend
    SetFilter = sFlt ' Set filter
End Function
