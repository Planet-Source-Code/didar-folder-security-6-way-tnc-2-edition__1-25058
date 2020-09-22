VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "GSI's Registration Form...."
   ClientHeight    =   5340
   ClientLeft      =   1485
   ClientTop       =   945
   ClientWidth     =   6165
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5340
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Set Password"
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   5895
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "1"
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "Confirm New Password"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Type New Password"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enter"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Code"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   5895
      Begin VB.Label Label13 
         Caption         =   "A3111-TNCPAT-T9664TGF"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Registration Code."
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Registration"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   5895
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4680
         Picture         =   "Form2.frx":0ECA
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   25
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   3960
         Picture         =   "Form2.frx":1794
         ScaleHeight     =   615
         ScaleWidth      =   735
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   3240
         Picture         =   "Form2.frx":205E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   22
         Top             =   840
         Width           =   480
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   1
            Left            =   0
            Picture         =   "Form2.frx":2928
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   23
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.Label Label9 
         Caption         =   "111-HVNNN-GFSA355"
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Product ID:"
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "4>e-mail:general@abnetbd.com"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "3>Internet:(www.general.com)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "2>Fax (44-001-77777778)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "1>Telephone(713568)"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Order  Your Registration:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Name"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   5895
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Your Full Name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Thank You For Buying General Product"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4920
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Me.WindowState = 1
Form3.Show
End Sub

Private Sub Command2_Click()
On Error GoTo error
Dim a As Variant
Dim b As Variant
Dim c As Variant
Kill ("c:\windows\sps.dll")
a = Text3.Text
b = Text4.Text
If a = b Then
c = b
FileNumber = FreeFile
filename = "c:\windows\sps.dll"
Open filename For Append As #FileNumber
Print #FileNumber, c
Close #FileNumber

Command3.Visible = False
Command1.Visible = True
Else
MsgBox "Invalid Password Setting...", 16, "Error!!"
End If


Exit Sub
error:
a = Text3.Text
b = Text4.Text
If a = b Then
c = b
FileNumber = FreeFile
filename = "c:\windows\sps.dll"
Open filename For Append As #FileNumber
Print #FileNumber, c
Close #FileNumber

Command1_Click
Else
MsgBox "Invalid Password Setting...", 16, "Error!!"
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Command1.Visible = False
End Sub

