VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin Project1.ucScreen ucPic 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10821
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Image"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin Project1.epCmDlg cd 
      Left            =   8760
      Top             =   240
      _ExtentX        =   661
      _ExtentY        =   635
   End
   Begin VB.Label lblZoom 
      Height          =   255
      Left            =   8760
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents tfilter As cBitFilter
Attribute tfilter.VB_VarHelpID = -1

Private Sub Command1_Click()
  cd.CancelError = False
  cd.DialogTitle = "Open Image"
  cd.Filter = "All Supported Images|*.jpg;*.gif;*.bmp;*.wmf;*.dib"
  cd.FilterIndex = 1
  cd.ShowOpen
  If cd.cFileName(1) <> "" Then
    Set ucPic.Picture = LoadPicture(cd.cFileName(1))
    lblZoom.Caption = "Zoom : 1:1"
    ucPic.ZoomFactor = 1
    ucPic.SetFocus
  End If
End Sub

Private Sub ucPic_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = vbCtrlMask Then
    If KeyCode = vbKeySubtract Then
      ucPic.ZoomOut
      lblZoom = "Zoom : " & ucPic.ZoomFactor & ":1"
    End If
    
    If KeyCode = vbKeyAdd Then
      ucPic.ZoomIn
      lblZoom = "Zoom : " & ucPic.ZoomFactor & ":1"
    End If
  End If
End Sub
