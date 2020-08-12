VERSION 5.00
Begin VB.UserControl Line 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   ScaleHeight     =   1335
   ScaleWidth      =   7110
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   4425
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   675
      X2              =   5025
      Y1              =   225
      Y2              =   225
   End
End
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
    On Error Resume Next
    UserControl.Height = 30
    UserControl.BackColor = UserControl.Parent.BackColor
End Sub

Private Sub UserControl_InitProperties()
        UserControl.Height = 30
End Sub

Private Sub UserControl_Paint()

Line1.X1 = 0
Line1.Y1 = 0
Line1.X2 = UserControl.Width
Line1.Y2 = 0

Line2.X1 = 0
Line2.Y1 = 20
Line2.X2 = UserControl.Width
Line2.Y2 = 20
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 30
    UserControl_Paint
End Sub

