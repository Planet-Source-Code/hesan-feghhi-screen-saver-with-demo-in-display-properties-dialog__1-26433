VERSION 5.00
Begin VB.Form frmScr 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   2880
   End
End
Attribute VB_Name = "frmScr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DirX, DirY, DirR, DirB, r, b

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 ShowCursor True 'Turn the cursor back on
 End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ShowCursor True 'Turn the cursor back on
 End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Static a
 a = a + 1
 If a = 5 Then
  ShowCursor True 'Turn the cursor back on
  End
 End If
End Sub

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
 Me.Line -(Screen.Width / 500 * DirX + Screen.Width / 2 * Cos(DirX), Screen.Height / 500 * DirY + Screen.Height / 2 * Sin(DirY)), RGB(r, 100, b)
 If DirX > 500 Then DirX = 0: DirY = 0: Me.Cls
End Sub
