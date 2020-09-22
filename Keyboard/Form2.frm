VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4500
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Height          =   5655
         Left            =   -120
         Picture         =   "Form2.frx":0ECA
         ScaleHeight     =   5595
         ScaleWidth      =   4755
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   4815
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Height          =   2895
            Left            =   0
            TabIndex        =   27
            Top             =   120
            Width           =   4335
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Back"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         ToolTipText     =   "Back To Previous Window"
         Top             =   6000
         Width           =   615
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Clean"
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         ToolTipText     =   "To Clean The Program From Selected Keypad Number"
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton Command200 
         Caption         =   "Hide"
         Default         =   -1  'True
         Height          =   255
         Left            =   2040
         TabIndex        =   0
         ToolTipText     =   "Load The Program And Hide"
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton Command204 
         Caption         =   "Disable Always Run"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         ToolTipText     =   "To Disable Run Automatically"
         Top             =   5640
         Width           =   1695
      End
      Begin VB.CommandButton Command203 
         Caption         =   "Always Run"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   23
         ToolTipText     =   "To Start Automatically While Windows Start"
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Exit"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         ToolTipText     =   "Exit To Windows System"
         Top             =   6000
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Text            =   "0"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Text            =   "0"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2"
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Text            =   "0"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "3"
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Text            =   "0"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Text            =   "0"
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton Command5 
         Caption         =   "5"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   3000
         Width           =   375
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Text            =   "0"
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CommandButton Command6 
         Caption         =   "6"
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "7"
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Text            =   "0"
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Text            =   "0"
         Top             =   4440
         Width           =   2655
      End
      Begin VB.CommandButton Command8 
         Caption         =   "8"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   4440
         Width           =   375
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Text            =   "0"
         Top             =   4920
         Width           =   2655
      End
      Begin VB.CommandButton Command9 
         Caption         =   "9"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   4920
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable Close Application"
         Height          =   195
         Left            =   720
         TabIndex        =   30
         ToolTipText     =   "To Enable/Disable Close Application."
         Top             =   5400
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   3240
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LayOut"
         Height          =   195
         Left            =   3360
         TabIndex        =   25
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General Multimedia KeyBoard"
         Height          =   195
         Left            =   720
         TabIndex        =   22
         Top             =   240
         Width           =   2505
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   21
         Top             =   720
         Width           =   1305
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   120
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.exe"
      DialogTitle     =   "General File System"
      Filter          =   "*.exe"
   End
End
Attribute VB_Name = "Form2"
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









Private Sub Check1_Click()
On Error GoTo err
tahmina:
If Check1.Value = 1 Then
Kill ("c:\windows\10.txt")
FileNumber = FreeFile
filename = "c:\windows\10.txt"
Open filename For Append As #FileNumber
Print #FileNumber, "1"
Close #FileNumber
Else
Kill ("c:\windows\10.txt")
FileNumber = FreeFile
filename = "c:\windows\10.txt"
Open filename For Append As #FileNumber
Print #FileNumber, "0"
Close #FileNumber
End If
Exit Sub
err:
FileNumber = FreeFile
filename = "c:\windows\10.txt"
Open filename For Append As #FileNumber
Close #FileNumber
GoTo tahmina
End Sub

Private Sub Command1_Click()
Dim count As Variant
On Error Resume Next
cmdlg.ShowOpen
If cmdlg.filename = "" Then
MsgBox "No File Selected.", 16, "Error"
Else
Text1.Text = cmdlg.filename
Kill ("c:\windows\1.txt")
FileNumber = FreeFile
filename = "c:\windows\1.txt"
Open filename For Append As #FileNumber
Print #FileNumber, Text1.Text
Close #FileNumber
cmdlg.filename = ""
End If
End Sub

Private Sub Command10_Click()
Dim ans As Variant
ans = MsgBox("Do You Really Want To Quit?..", vbYesNo, "Quit")
If ans = vbNo Then
Load Me
Else
End
End If
End Sub

Private Sub Command11_Click()
Dim filename As String
On Error Resume Next
filename = "c:\windows\"
x = InputBox("Enter Which Keypad Setting Program You Want To Delete..", "Delete Program")
Kill (filename & x & ".txt")
FileNumber = FreeFile
filename = "c:\windows\x.txt"
Open filename For Append As #FileNumber
Print #FileNumber, ""
Close #FileNumber
End Sub

Private Sub Command12_Click()
On Error Resume Next
Me.Hide
Form1.Timer1.Enabled = False
Form1.Frame1.Height = 6095
Form1.Height = 6320
Form1.Show
End Sub

Private Sub Command2_Click()
Dim count As Variant
On Error Resume Next
cmdlg.ShowOpen
If cmdlg.filename = "" Then
MsgBox "No File Selected.", 16, "Error"
Else
Text2.Text = cmdlg.filename
Kill ("c:\windows\2.txt")
FileNumber = FreeFile
filename = "c:\windows\2.txt"
Open filename For Append As #FileNumber
Print #FileNumber, Text2.Text
Close #FileNumber
cmdlg.filename = ""
End If
End Sub

Private Sub Command200_Click()

Load Form1
Unload Me
End Sub

Private Sub Command203_Click()
On Error Resume Next
If Command203.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "TaskBarCont", "TaskBarCont"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "TaskBarCont", "C:\Windows\TaskShot.exe"
End If
End Sub

Private Sub Command204_Click()
On Error Resume Next
If Command204.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "TaskBarCont", "TaskBarCont"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "TaskBarCont", "0"
End If

End Sub

Private Sub Command3_Click()
Dim count As Variant
On Error Resume Next
cmdlg.ShowOpen
If cmdlg.filename = "" Then
MsgBox "No File Selected.", 16, "Error"
Else
Text3.Text = cmdlg.filename
Kill ("c:\windows\3.txt")
FileNumber = FreeFile
filename = "c:\windows\3.txt"
Open filename For Append As #FileNumber
Print #FileNumber, Text3.Text
Close #FileNumber
cmdlg.filename = ""
End If
End Sub

Private Sub Command4_Click()
Dim count As Variant
On Error Resume Next
cmdlg.ShowOpen
If cmdlg.filename = "" Then
MsgBox "No File Selected.", 16, "Error"
Else

Text4.Text = cmdlg.filename
Kill ("c:\windows\4.txt")
FileNumber = FreeFile
filename = "c:\windows\4.txt"
Open filename For Append As #FileNumber
Print #FileNumber, Text4.Text
Close #FileNumber
cmdlg.filename = ""
End If
End Sub

Private Sub Command5_Click()
Dim count As Variant
On Error Resume Next
cmdlg.ShowOpen
If cmdlg.filename = "" Then
MsgBox "No File Selected.", 16, "Error"
Else

Text5.Text = cmdlg.filename
Kill ("c:\windows\5.txt")
FileNumber = FreeFile
filename = "c:\windows\5.txt"
Open filename For Append As #FileNumber
Print #FileNumber, Text5.Text
Close #FileNumber
cmdlg.filename = ""
End If
End Sub

Private Sub Command6_Click()
Dim count As Variant
On Error Resume Next
cmdlg.ShowOpen
If cmdlg.filename = "" Then
MsgBox "No File Selected.", 16, "Error"
Else
Text6.Text = cmdlg.filename
Kill ("c:\windows\6.txt")
FileNumber = FreeFile
filename = "c:\windows\6.txt"
Open filename For Append As #FileNumber
Print #FileNumber, Text6.Text
Close #FileNumber
cmdlg.filename = ""
End If
End Sub

Private Sub Command7_Click()
Dim count As Variant
On Error Resume Next
cmdlg.ShowOpen
If cmdlg.filename = "" Then
MsgBox "No File Selected.", 16, "Error"
Else
Text7.Text = cmdlg.filename
Kill ("c:\windows\7.txt")
FileNumber = FreeFile
filename = "c:\windows\7.txt"
Open filename For Append As #FileNumber
Print #FileNumber, Text7.Text
Close #FileNumber
cmdlg.filename = ""
End If
End Sub

Private Sub Command8_Click()
Dim count As Variant
On Error Resume Next
cmdlg.ShowOpen
If cmdlg.filename = "" Then
MsgBox "No File Selected.", 16, "Error"
Else
Text8.Text = cmdlg.filename
Kill ("c:\windows\8.txt")
FileNumber = FreeFile
filename = "c:\windows\8.txt"
Open filename For Append As #FileNumber
Print #FileNumber, Text8.Text
Close #FileNumber
cmdlg.filename = ""
End If
End Sub

Private Sub Command9_Click()
Dim count As Variant
On Error Resume Next
cmdlg.ShowOpen
If cmdlg.filename = "" Then
MsgBox "No File Selected.", 16, "Error"
Else
Text9.Text = cmdlg.filename
Kill ("c:\windows\9.txt")
FileNumber = FreeFile
filename = "c:\windows\9.txt"
Open filename For Append As #FileNumber
Print #FileNumber, Text9.Text
Close #FileNumber

End If
End Sub

Private Sub Form_Load()
Dim count As Variant

On Error Resume Next
Unload Form1



 FileNumber = FreeFile
filename = "c:\windows\10.txt"
Open "c:\windows\10.txt" For Input As #FileNumber
Text10.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text10.Text) - 2)
Text10.Text = Left(Text10.Text, count)

If Text10.Text = 1 Then
Check1.Value = 1
Else
Check1.Value = 0
End If



FileNumber = FreeFile
filename = "c:\windows\1.txt"
Open "c:\windows\1.txt" For Input As #FileNumber
Text1.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text1.Text) - 2)
Text1.Text = Left(Text1.Text, count)

 FileNumber = FreeFile
filename = "c:\windows\2.txt"
Open "c:\windows\2.txt" For Input As #FileNumber
Text2.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text2.Text) - 2)
Text2.Text = Left(Text2.Text, count)

 FileNumber = FreeFile
filename = "c:\windows\3.txt"
Open "c:\windows\3.txt" For Input As #FileNumber
Text3.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text3.Text) - 2)
Text3.Text = Left(Text3.Text, count)


 FileNumber = FreeFile
filename = "c:\windows\4.txt"
Open "c:\windows\4.txt" For Input As #FileNumber
Text4.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text4.Text) - 2)
Text4.Text = Left(Text4.Text, count)



 FileNumber = FreeFile
filename = "c:\windows\5.txt"
Open "c:\windows\5.txt" For Input As #FileNumber
Text5.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text5.Text) - 2)
Text5.Text = Left(Text5.Text, count)


 FileNumber = FreeFile
filename = "c:\windows\6.txt"
Open "c:\windows\6.txt" For Input As #FileNumber
Text6.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text6.Text) - 2)
Text6.Text = Left(Text6.Text, count)



 FileNumber = FreeFile
filename = "c:\windows\7.txt"
Open "c:\windows\7.txt" For Input As #FileNumber
Text7.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text7.Text) - 2)
Text7.Text = Left(Text7.Text, count)



 FileNumber = FreeFile
filename = "c:\windows\8.txt"
Open "c:\windows\8.txt" For Input As #FileNumber
Text8.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text8.Text) - 2)
Text8.Text = Left(Text8.Text, count)



 FileNumber = FreeFile
filename = "c:\windows\9.txt"
Open "c:\windows\9.txt" For Input As #FileNumber
Text9.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text9.Text) - 2)
Text9.Text = Left(Text9.Text, count)

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1.Visible = True
End Sub



Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1.Visible = False
End Sub

