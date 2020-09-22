VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4770
   ClientLeft      =   1470
   ClientTop       =   840
   ClientWidth     =   6495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7800
      TabIndex        =   24
      Text            =   "\"
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   8280
      Pattern         =   "*.txt"
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   6120
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Picture         =   "Form1.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit To Windows System"
      Top             =   4080
      Width           =   615
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   2115
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5505
      Left            =   0
      Picture         =   "Form1.frx":264C
      ScaleHeight     =   5505
      ScaleWidth      =   6525
      TabIndex        =   3
      Top             =   0
      Width           =   6525
      Begin VB.CommandButton Command14 
         Caption         =   "Command14"
         Height          =   615
         Left            =   1560
         TabIndex        =   21
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00000000&
         Height          =   615
         Left            =   4200
         Picture         =   "Form1.frx":C7A3
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Remove Low Security."
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00000000&
         Height          =   615
         Left            =   5520
         Picture         =   "Form1.frx":CAAD
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Protect Folder As RecycleBin Icon"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00000000&
         Height          =   615
         Left            =   4920
         Picture         =   "Form1.frx":CDB7
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Protect Folder As Task Scheduler Icon"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00000000&
         Height          =   615
         Left            =   4320
         Picture         =   "Form1.frx":D0C1
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Protect Folder As Printer Icon"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00000000&
         Height          =   615
         Left            =   3720
         Picture         =   "Form1.frx":D3CB
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Protect Folder As Control Panel"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         MaskColor       =   &H00000000&
         Picture         =   "Form1.frx":D6D5
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Protect Folder As Windows Icon (Recommanded For High Security)"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   5760
         TabIndex        =   10
         Top             =   4800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   4800
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Picture         =   "Form1.frx":D9DF
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Change Password Wizard"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   6000
         Width           =   5055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Click To Unprotect The Folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folder Protection Mode..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   25
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TNC Edition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   5280
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HELP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   5400
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   6120
         TabIndex        =   18
         ToolTipText     =   "Exit To Windows"
         Top             =   0
         Width           =   240
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5760
         TabIndex        =   17
         ToolTipText     =   "Minimize"
         Top             =   120
         Width           =   285
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   6120
         X2              =   6225
         Y1              =   240
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         X1              =   6120
         X2              =   6240
         Y1              =   120
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         X1              =   5880
         X2              =   6000
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "A Program by Didar"
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
         Left            =   2400
         TabIndex        =   5
         Top             =   4410
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " General Corporation Bangladesh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1800
         TabIndex        =   4
         Top             =   4560
         Width           =   2820
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filename As String
Dim filename2 As String
Dim data As String
Dim Y As Integer
Dim X As Integer
Dim z As Integer
Dim TAHMINA As String






Private Sub Command1_Click()
Dim ans As Integer
ans = MsgBox("Do You Really Want to Quit?", 49, "Quit?")
If ans = vbCancel Then
Load Me
Else
End
End If
End Sub

Private Sub Command10_Click()
 Dim count As String
 Dim X As Long
 Dim ext As String
 Dim filename3 As String
 Dim count2 As Variant
Dim i As Variant
On Error Resume Next

Text3.Text = Dir1.Path

If UCase(Text3.Text) = "C:\WINDOWS" Then
MsgBox "Operating System Can't Be Protected", 16, "Illegal Operation"
Exit Sub
End If

If UCase(Text3.Text) = "C:\WINDOWS\DESKTOP" Then
MsgBox "Operating System Can't Be Protected", 16, "Illegal Operation"
Exit Sub
End If






For i = 0 To Len(Text3.Text)
startposition = i
startposition = InStr(startposition + 1, Text3.Text, Text4.Text, vbTextCompare)
If startposition > 0 Then
Text3.SelStart = startposition - 1
Text3.SetFocus
Text5.Text = Text3.SelStart
 count2 = Len(Text3.Text) - Text5.Text - 1
  data = Right(Text3.Text, count2)
file$ = Left(Text3.Text, Len(Text3.Text) - count2)

End If
Next i



ext = ".{645FF040-5081-101B-9F08-00AA002F954E}"
 filename = file & data & ext
Name Dir1.Path As filename

Dir1.Path = file

 


End Sub

Private Sub Command11_Click()
Dim count As String
Dim count2 As String
Dim filename7 As String
Dim filename3 As String
Dim i As Variant

On Error Resume Next


Text3.Text = Dir1.Path

For i = 0 To Len(Text3.Text)
startposition = i
startposition = InStr(startposition + 1, Text3.Text, Text4.Text, vbTextCompare)
If startposition > 0 Then
Text3.SelStart = startposition - 1
Text5.Text = Text3.SelStart
 count2 = Len(Text3.Text) - Text5.Text - 1
  data = Right(Text3.Text, count2)
 file$ = Left(Text3.Text, Len(Text3.Text) - count2)
 End If
Next i

filename = Dir1.Path
  Text1.Text = filename
  filename7 = filename

count = (Len(data) - 39)
Text1.Text = Left(data, count)
filename = file & Right(Text1.Text, count)
 
Text1.Text = "ren " & filename7 & " " & filename
Name filename7 As filename
Dir1.Path = file
End Sub




Private Sub Command3_Click()
Kill ("c:\windows\sps.dll")
Unload Me
Form4.Show
End Sub


Private Sub Command6_Click()
 Dim count As String
 Dim X As Long
 Dim ext As String
 Dim filename3 As String
 Dim count2 As Variant
Dim i As Variant
On Error Resume Next


If UCase(Text3.Text) = "C:\WINDOWS" Then
MsgBox "Operating System Can't Be Protected", 16, "Illegal Operation"
Exit Sub
End If


If UCase(Text3.Text) = "C:\WINDOWS\DESKTOP" Then
MsgBox "Operating System Can't Be Protected", 16, "Illegal Operation"
Exit Sub
End If



For i = 0 To Len(Text3.Text)
startposition = i
startposition = InStr(startposition + 1, Text3.Text, Text4.Text, vbTextCompare)
If startposition > 0 Then
Text3.SelStart = startposition - 1
Text3.SetFocus
Text5.Text = Text3.SelStart
 
 count2 = Len(Text3.Text) - Text5.Text - 1
  data = Right(Text3.Text, count2)
file$ = Left(Text3.Text, Len(Text3.Text) - count2)
End If
Next i






ext = ".{00021401-0000-0000-C000-000000000046}"
 filename = file & data & ext
Name Dir1.Path As filename

Dir1.Path = file
End Sub

Private Sub Command7_Click()
 Dim count As String
 Dim X As Long
 Dim ext As String
 Dim filename3 As String
 Dim count2 As Variant
Dim i As Variant
On Error Resume Next

Text3.Text = Dir1.Path


If UCase(Text3.Text) = "C:\WINDOWS" Then
MsgBox "Operating System Can't Be Protected", 16, "Illegal Operation"
Exit Sub
End If

If UCase(Text3.Text) = "C:\WINDOWS\DESKTOP" Then
MsgBox "Operating System Can't Be Protected", 16, "Illegal Operation"
Exit Sub
End If




For i = 0 To Len(Text3.Text)
startposition = i
startposition = InStr(startposition + 1, Text3.Text, Text4.Text, vbTextCompare)
If startposition > 0 Then
Text3.SelStart = startposition - 1
Text3.SetFocus
Text5.Text = Text3.SelStart
 
count2 = Len(Text3.Text) - Text5.Text - 1
data = Right(Text3.Text, count2)
file$ = Left(Text3.Text, Len(Text3.Text) - count2)
End If
Next i




ext = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"
 filename = file & data & ext
Name Dir1.Path As filename
Dir1.Path = file
End Sub

Private Sub Command8_Click()
 Dim count As String
 Dim X As Long
 Dim ext As String
 Dim filename3 As String
 Dim count2 As Variant
Dim i As Variant
On Error Resume Next
Text3.Text = Dir1.Path




If UCase(Text3.Text) = "C:\WINDOWS" Then
MsgBox "Operating System Can't Be Protected", 16, "Illegal Operation"
Exit Sub
End If
If UCase(Text3.Text) = "C:\WINDOWS\DESKTOP" Then
MsgBox "Operating System Can't Be Protected", 16, "Illegal Operation"
Exit Sub
End If





For i = 0 To Len(Text3.Text)
startposition = i
startposition = InStr(startposition + 1, Text3.Text, Text4.Text, vbTextCompare)
If startposition > 0 Then
Text3.SelStart = startposition - 1
Text3.SetFocus
Text5.Text = Text3.SelStart
 
 count2 = Len(Text3.Text) - Text5.Text - 1
  data = Right(Text3.Text, count2)
file$ = Left(Text3.Text, Len(Text3.Text) - count2)
End If
Next i








ext = ".{2227A280-3AEA-1069-A2DE-08002B30309D}"
  filename = file & data & ext
Name Dir1.Path As filename
Dir1.Path = file
End Sub

Private Sub Command9_Click()
 Dim count As String
 Dim X As Long
 Dim ext As String
 Dim filename3 As String
 Dim count2 As Variant
Dim i As Variant
On Error Resume Next

Text3.Text = Dir1.Path


If UCase(Text3.Text) = "C:\WINDOWS" Then
MsgBox "Operating System Can't Be Protected", 16, "Illegal Operation"
Exit Sub
End If

If UCase(Text3.Text) = "C:\WINDOWS\DESKTOP" Then
MsgBox "Operating System Can't Be Protected", 16, "Illegal Operation"
Exit Sub
End If










For i = 0 To Len(Text3.Text)
startposition = i
startposition = InStr(startposition + 1, Text3.Text, Text4.Text, vbTextCompare)
If startposition > 0 Then
Text3.SelStart = startposition - 1
Text3.SetFocus
Text5.Text = Text3.SelStart
 
 count2 = Len(Text3.Text) - Text5.Text - 1
  data = Right(Text3.Text, count2)
file$ = Left(Text3.Text, Len(Text3.Text) - count2)
End If
Next i


ext = ".{D6277990-4C6A-11CF-8D87-00AA0060F5BF}"
 filename = file & data & ext
Name Dir1.Path As filename
Dir1.Path = file
End Sub

Private Sub Dir1_Change()
On Error Resume Next

Dir1.Path = Drive1.Drive
Text3.Text = Dir1.Path
TAHMINA = Dir1.Path

End Sub

Private Sub Drive1_Change()
On Error GoTo err
Dir1.Path = Drive1.Drive
Exit Sub
err:
MsgBox "Drive Is Not Ready", 16, "Error Reading.."
End Sub


Private Sub Label5_Click()
Me.Hide
End Sub

Private Sub Label6_Click()
Dim ans As Integer
ans = MsgBox("Do You Really Want to Quit?", 49, "Quit?")
If ans = vbCancel Then
Load Me
Else
End
End If
End Sub

Private Sub Form_Load()
Dim tony As Integer
On Error Resume Next
Dir1.Path = "c:\"
X = 0
Y = 1
z = 0
End Sub
