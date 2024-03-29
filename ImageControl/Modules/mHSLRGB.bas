Attribute VB_Name = "mHSLRGB"
'//
'// mHSLRGB module
'// Special thanks vbaccelerator.com
'//

Option Explicit



Public Sub RGBtoHSL(ByVal R As Long, ByVal G As Long, ByVal B As Long, _
                    ByRef H As Single, ByRef S As Single, ByRef L As Single)
                    
Dim Max As Single
Dim Min As Single
Dim delta As Single
Dim rR As Single, rG As Single, rB As Single

   '// Given:   RGB each in [0,1]
   '// Desired: H in [0,360] and S in [0,1], except if S = 0, then H = UNDEFINED
    rR = R / 255: rG = G / 255: rB = B / 255
   
    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    L = (Max + Min) / 2
   
   '// Calculate saturation:
   '// Acrhomatic case
    If (Max = Min) Then
        S = 0
        H = 0
    Else
   '// Chromatic case
       '// First calculate the saturation
        If (L <= 0.5) Then
            S = (Max - Min) / (Max + Min)
          Else
            S = (Max - Min) / (2 - Max - Min)
        End If
       '// Next calculate the hue
        delta = Max - Min
        If (rR = Max) Then
            H = (rG - rB) / delta     '// Resulting color is between yellow and magenta
          ElseIf (rG = Max) Then
            H = 2 + (rB - rR) / delta '// Resulting color is between cyan and yellow
          ElseIf (rB = Max) Then
            H = 4 + (rR - rG) / delta '// Resulting color is between magenta and cyan
        End If
    End If
End Sub

Public Sub HSLtoRGB(ByVal H As Single, ByVal S As Single, ByVal L As Single, _
                    ByRef R As Long, ByRef G As Long, ByRef B As Long)
      
Dim rR As Single, rG As Single, rB As Single
Dim Min As Single, Max As Single

   '// Achromatic case:
    If (S = 0) Then
        rR = L: rG = L: rB = L
    Else
   '// Chromatic case:
        If (L <= 0.5) Then
           '// s = (Max - Min) / (Max + Min)
            Min = L * (1 - S)
          Else
           '// s = (Max - Min) / (2 - Max - Min)
            Min = L - S * (1 - L)
        End If
        Max = 2 * L - Min
      
       '// Now depending on sector we can evaluate the H,L,S:
        If (H < 1) Then
            rR = Max
            If (H < 0) Then
                rG = Min
                rB = rG - H * (Max - Min)
              Else
                rB = Min
                rG = H * (Max - Min) + rB
            End If
          ElseIf (H < 3) Then
            rG = Max
            If (H < 2) Then
                rB = Min
                rR = rB - (H - 2) * (Max - Min)
              Else
                rR = Min
                rB = (H - 2) * (Max - Min) + rR
            End If
          Else
            rB = Max
            If (H < 4) Then
                rR = Min
                rG = rR - (H - 4) * (Max - Min)
              Else
                rG = Min
                rR = (H - 4) * (Max - Min) + rG
            End If
        End If
   End If
   R = rR * 255: G = rG * 255: B = rB * 255
End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
    If (rR > rG) Then
        If (rR > rB) Then Maximum = rR Else Maximum = rB
      Else
        If (rB > rG) Then Maximum = rB Else Maximum = rG
    End If
End Function

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
    If (rR < rG) Then
        If (rR < rB) Then Minimum = rR Else Minimum = rB
      Else
        If (rB < rG) Then Minimum = rB Else Minimum = rG
    End If
End Function



Public Function RotateH40(ByVal H As Long) As Long
   '// Rotate Hue ->[Red...Red]
    If (H > 200) Then RotateH40 = H - 240 Else RotateH40 = H
End Function



