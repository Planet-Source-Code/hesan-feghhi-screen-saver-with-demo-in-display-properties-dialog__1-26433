VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   720
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DirX, DirY, DirR, DirB, r, b

Private Sub Form_Load()
 DirX = 1
 DirY = 1
 DirR = 1
 DirB = -1
 b = 255
End Sub

Private Sub Timer1_Timer()
 DirX = DirX + 1
 DirY = DirY + 1
 r = r + DirR
 If r = 255 Or r = 0 Then DirR = -DirR
 b = b + DirB
 If b = 255 Or b = 0 Then DirB = -DirB
 Me.Line -(5 * DirX + 1000 * Cos(DirX), 5 * DirY + 1000 * Sin(DirY)), RGB(r, 100, b)
 If DirX > 500 Then DirX = 0: DirY = 0: Me.Cls
End Sub
