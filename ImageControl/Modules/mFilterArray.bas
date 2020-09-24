Attribute VB_Name = "mFilterArray"
'//
'// Filter array control module
'// Special thanks vbaccelerator.com
'//

Option Explicit

Public Enum MaskFilterID
    [mfltBlur]
    [mfltBlurMore]
    [mfltSoften]
    [mfltSoftenMore]
    [mfltSharpen]
    [mfltSharpenMore]
    [mfltBump]
    [mfltBumpMore]
    [mfltContour]
    [mfltContourMore]
    [mfltContourSoft]
    [mfltContourMoreSoft]
    [mfltEmboss]
    [mfltEngrave]
End Enum

Public Type MASK_FILTER
    mfltArray() As Long
    mfltWeight As Long
    mfltRFct As Long
    mfltGFct As Long
    mfltBFct As Long
End Type



Public Sub GetArrayInfo(ByRef mfltMask As MASK_FILTER, ByVal mfltID As MaskFilterID)

  Dim tMskFlt() As Long
  Dim tV() As String
  Dim tF() As String
  Dim tV_UB As Long
  Dim i As Long, j As Long

    Select Case mfltID
      Case [mfltBlur]
        tV() = Split("1|1|1|1|1|1|1|1|1", "|")
        tF() = Split("9|0|0|0", "|")
      Case [mfltBlurMore]
        tV() = Split("1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1", "|")
        tF() = Split("25|0|0|0", "|")
      Case [mfltSoften]
        tV() = Split("0|1|0|1|3|1|0|1|0", "|")
        tF() = Split("7|0|0|0", "|")
      Case [mfltSoftenMore]
        tV() = Split("1|3|1|3|9|3|1|3|1", "|")
        tF() = Split("25|0|0|0", "|")
      Case [mfltSharpen]
        tV() = Split("-1|-1|-1|-1|15|-1|-1|-1|-1", "|")
        tF() = Split("7|0|0|0", "|")
      Case [mfltSharpenMore]
        tV() = Split("0|-1|0|-1|5|-1|0|-1|0", "|")
        tF() = Split("1|0|0|0", "|")
      Case [mfltBump]
        tV() = Split("-1|-1|0|-1|4|1|0|1|1", "|")
        tF() = Split("4|0|0|0", "|")
      Case [mfltBumpMore]
        tV() = Split("-1|-1|0|-1|2|1|0|1|1", "|")
        tF() = Split("2|0|0|0", "|")
      Case [mfltContour]
        tV() = Split("-1|-1|-1|-1|8|-1|-1|-1|-1", "|")
        tF() = Split("1|255|255|255", "|")
      Case [mfltContourMore]
        tV() = Split("-2|-2|-2|-2|16|-2|-2|-2|-2", "|")
        tF() = Split("1|255|255|255", "|")
      Case [mfltContourSoft]
        tV() = Split("0|-1|0|-1|4|-1|0|-1|0", "|")
        tF() = Split("1|255|255|255", "|")
      Case [mfltContourMoreSoft]
        tV() = Split("0|-2|0|-2|8|-2|0|-2|0", "|")
        tF() = Split("1|255|255|255", "|")
      Case [mfltEmboss]
        tV() = Split("0|0|0|0|1|0|0|0|-1", "|")
        tF() = Split("1|128|128|128", "|")
      Case [mfltEngrave]
        tV() = Split("0|0|0|0|-1|0|0|0|1", "|")
        tF() = Split("1|128|128|128", "|")
    End Select

    tV_UB = Sqr(UBound(tV) + 1) - 1
    ReDim tMskFlt(tV_UB, tV_UB) As Long

    For i = 0 To tV_UB
        For j = 0 To tV_UB
            tMskFlt(i, j) = tV((tV_UB + 1) * i + j)
        Next j
    Next i
    mfltMask.mfltArray = tMskFlt
    mfltMask.mfltWeight = tF(0)
    mfltMask.mfltRFct = tF(1)
    mfltMask.mfltGFct = tF(2)
    mfltMask.mfltBFct = tF(3)

End Sub

':) Ulli's VB Code Formatter V2.13.2 (16/07/02 11:42:19) 31 + 71 = 102 Lines
