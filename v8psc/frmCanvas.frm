VERSION 5.00
Begin VB.Form frmCanvas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "300"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   Icon            =   "frmCanvas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   4020
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerDoAnimation 
      Interval        =   1
      Left            =   120
      Top             =   360
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Events
Event OnAnimate()
Event OnMouseDown()
Event OnMouseUp()
Event OnMouseEvent(Button As Integer, Shift As Integer, X As Single, Y As Single)


Private Sub Form_Load()

'    Dim lngRetVal As Long
'    lngRetVal = ShowCursor(0)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OnMouseDown
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button > 0 Then
        RaiseEvent OnMouseEvent(Button, Shift, X, Y)
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OnMouseUp
End Sub


Private Sub Form_Unload(Cancel As Integer)

'    Dim lngRetVal As Long
'    lngRetVal = ShowCursor(1)

End Sub

Private Sub TimerDoAnimation_Timer()
    RaiseEvent OnAnimate
End Sub
