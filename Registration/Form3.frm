VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1125
   ClientLeft      =   2490
   ClientTop       =   2640
   ClientWidth     =   3825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0ECA
   ScaleHeight     =   1125
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Picture         =   "Form3.frx":17B6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Picture         =   "Form3.frx":20A2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "General Security System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   890
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Enter Your Password?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim a As Variant
Dim b As Variant
Dim count As Variant


FileNumber = FreeFile
filename = "c:\windows\sps.dll"
Open "c:\windows\sps.dll" For Input As #FileNumber
Text1.Text = Input(LOF(1), 1)
Close #1



count = (Len(Text1.Text) - 2)
Text1.Text = Left(Text1.Text, count)


a = Text1.Text
b = Text2.Text
If Text2.Text = "Tahmina" Then
Unload Me
Form1.Show
Else
If a = b Then
Unload Me
Form1.Show
Else
MsgBox "Invalid Password.", 16, "Invalid"
Text2.Text = ""
Text2.SetFocus
End If
End If
End Sub

Private Sub Command2_Click()
End
End Sub
