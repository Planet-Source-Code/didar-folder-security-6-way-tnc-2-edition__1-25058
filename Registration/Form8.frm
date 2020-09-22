VERSION 5.00
Begin VB.Form Form7 
   ClientHeight    =   5220
   ClientLeft      =   1485
   ClientTop       =   660
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form7"
   Picture         =   "Form8.frx":0ECA
   ScaleHeight     =   5220
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   195
      Left            =   9120
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Del Reg"
      Height          =   495
      Left            =   9000
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   600
      Top             =   4200
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

 
 Const HKEY_CURRENT_USER = &H80000001
Const HKEY_CLASSES_ROOT = &H80000000
 
Public Sub delString(hKey As Long, StrPath As String)
   Dim KeyH&
    r = RegDeleteKey(hKey, StrPath)
     r = RegCloseKey(KeyH&)
   End Sub

Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub
Public Sub SaveString1(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 0, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub






Private Sub Command2_Click()
If Command2.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
delString HKEY_CLASSES_ROOT, "Directory\shell\Folder Security 4"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
delString HKEY_CLASSES_ROOT, "Directory\shell\Folder Security 4"
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Command3.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_CLASSES_ROOT, "Directory\shell\Folder Security 4\command", "", ""
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_CLASSES_ROOT, "Directory\shell\Folder Security 4\command", "", (App.Path & "\Folder Security 4.0.exe")
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Command3_Click
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Form10.Show
Unload Me
Load Form4
Form4.Show
Timer1.Enabled = False
End Sub


