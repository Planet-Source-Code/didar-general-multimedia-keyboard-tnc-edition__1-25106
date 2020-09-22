VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3525
   ControlBox      =   0   'False
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   80
      Width           =   3255
      Begin VB.CommandButton Command3 
         Caption         =   "Unload Me"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Disable TaskBar Control"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Always Run While Windows Starts"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CopyRight By General Corporation"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2520
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
 Const HKEY_LOCAL_MACHINE = &H80000002
 Const SPIF_UPDATEINIFILE = &H1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2
Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub


Private Sub Command1_Click()
On Error Resume Next
If Command3.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "TaskBarCont", "TaskBarCont"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "TaskBarCont", "C:\Windows\TaskShot.exe"
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Command4.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "TaskBarCont", "TaskBarCont"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "TaskBarCont", "0"
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim ans As Variant
ans = MsgBox("Do You Really Want To Quit?..", vbYesNo, "Quit")
If ans = vbNo Then
Load Me
Else
End
End If
End Sub
