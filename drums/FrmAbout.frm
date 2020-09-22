VERSION 5.00
Begin VB.Form FrmAbout 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   2520
   ClientLeft      =   1875
   ClientTop       =   1530
   ClientWidth     =   5340
   ControlBox      =   0   'False
   Icon            =   "FrmAbout.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Author:        Stuart Pennington.
' Project:       Midi Percussion Sequencer (Drum Machine).
' Test Platform: Windows 98SE
' Processor:     P2 300MHz.

Private Sub Form_Load()
  PositionForm Me
End Sub

Private Sub Form_MouseUp(Button%, Shift%, X!, Y!)
  If Button = 1 Then Unload Me
End Sub

Private Sub Form_Paint()
  Dim Rct As RECT

  SetRect Rct, 0, 0, ScaleWidth, ScaleHeight
  DrawEdge hDC, Rct, 5, 15

  DrawIconEx hDC, 11, 11, FrmMicroKit.Icon, 32, 32, 0, 0, 3

  FontBold = True
  FontSize = 12
  TextOut hDC, 50, 27, "MicroKit", 8

  FontBold = False
  FontSize = 8
  TextOut hDC, 123, 32, "Version 1.01", 12

  FontSize = 10
  TextOut hDC, 51, 49, "Rhythms For Windows", 19
  FontSize = 8
  TextOut hDC, 51, 78, "Programmer:", 11

  TextOut hDC, 117, 78, "Stuart Pennington / Modified by Aldo Vargas", 43

  TextOut hDC, 51, 92, "Language:", 9

  TextOut hDC, 117, 92, "Microsoft Visual Basic 5.00 Professional", 40

  FontBold = True
  TextOut hDC, 51, 128, "© 2000 CyberVision", 18
End Sub
