VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   2175
   ClientTop       =   -3855
   ClientWidth     =   90
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Shell_NotifyIcon Lib "SHELL32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
      Private abd As NOTIFYICONDATA

      Private Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type

      Private Const NIF_MESSAGE = &H1
      Private Const NIF_ICON = &H2
      Private Const NIF_TIP = &H4
      Private Const Mouse_Move = 512
      Private Const Mouse_Left_Down = 513
      Private Const Mouse_Left_Click = 514
      Private Const Mouse_Left_DbClick = 515
      Private Const Mouse_Right_Down = 516
      Private Const Mouse_Right_Click = 517
      Private Const Mouse_Right_DbClick = 518
      Private Const Mouse_Button_Down = 519
      Private Const Mouse_Button_Click = 520
      Private Const Mouse_Button_DbClick = 521
      

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
     msg = X
Else
     msg = X / Screen.TwipsPerPixelX
End If

Select Case msg
            Case Mouse_Left_DbClick
            If Form3.Visible = True Then
             MsgBox "Cannot Show Main Form", 16, "No Access"
            Else
            Form1.WindowState = vbNormal
            Form1.Show
            End If
            End Select
        End Sub


Private Sub Form_Load()
On Error Resume Next
Dir1.Path = "c:\"
 With abd
      .cbSize = Len(abd)
      .hwnd = Me.hwnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = Mouse_Move
      .hIcon = Me.Icon
      .szTip = "Folder Security 4.0 (New) --General Corporation Bangladesh." & vbNullChar
   End With
   Shell_NotifyIcon NIM_ADD, abd
   Me.Hide
End Sub
