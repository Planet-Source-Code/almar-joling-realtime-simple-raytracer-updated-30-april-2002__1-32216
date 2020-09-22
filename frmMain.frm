VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RayTrace"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   459
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar scrlLightVer 
      Height          =   6570
      Left            =   7140
      Max             =   600
      Min             =   -600
      TabIndex        =   2
      Top             =   60
      Width           =   210
   End
   Begin VB.HScrollBar scrlLightHor 
      Height          =   210
      Left            =   30
      Max             =   600
      Min             =   -600
      TabIndex        =   1
      Top             =   6660
      Width           =   7065
   End
   Begin VB.PictureBox picRay 
      BorderStyle     =   0  'None
      Height          =   6525
      Left            =   45
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   435
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   464
      TabIndex        =   0
      Top             =   15
      Width           =   6960
      Begin VB.Timer tmrFPS 
         Interval        =   1000
         Left            =   3960
         Top             =   5160
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//Realtime raytracer
'//Original (c++) version and other nice
'//Raytrace versions (with shadows, cilinders, etc)
'//Can be found at http://www.2tothex.com/
'//VB port by Almar Joling / quadrantwars@quadrantwars.com
'//Websites: http://www.quadrantwars.com (my game)
'//          http://vbfibre.digitalrice.com (Many VB speed tricks with benchmarks)

'//This code is highly optimized. If you manage to gain some more FPS
'//I'm always interested =-)

'//Finished @ 01/03/2002
'//Feel free to post this code anywhere, but please leave the above info
'//and author info intact. Thank you.

Private Sub Form_Load()
    Me.Show
    DoEvents
    mdlMain.Main
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub

Private Sub scrlLightHor_Change()
    '//Change light position
    LightLoc.X = scrlLightHor.Value
End Sub

Private Sub scrlLightVer_Change()
    '//Change light position (inversed)
    LightLoc.Y = -scrlLightVer.Value
End Sub

Private Sub tmrFPS_Timer()
Me.Caption = "RayTrace :: " & CStr(iFPS) & "fps"
iFPS = 0
End Sub
