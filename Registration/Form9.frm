VERSION 5.00
Begin VB.Form Form8 
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   0
      Top             =   120
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
On Error Resume Next
Load Form1
Form1.Show
Timer1.Enabled = False
Unload Me
End Sub
