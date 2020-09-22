VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6390
   ControlBox      =   0   'False
   Icon            =   "DirectInput_KeyBoard_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6135
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   4680
         TabIndex        =   44
         Text            =   "0"
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "DirectInput_KeyBoard_Form.frx":0ECA
         Left            =   4800
         List            =   "DirectInput_KeyBoard_Form.frx":0EE9
         TabIndex        =   42
         Top             =   4560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Set Vol"
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
         Left            =   4800
         TabIndex        =   41
         ToolTipText     =   "To Select The Volume Control, If (+ and --)  Not Works Properly"
         Top             =   4560
         Width           =   855
      End
      Begin VB.Timer Timer2 
         Left            =   240
         Top             =   240
      End
      Begin VB.Timer Timer1 
         Interval        =   1500
         Left            =   480
         Top             =   480
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   5880
         TabIndex        =   39
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command301 
         Caption         =   "Command3"
         Height          =   615
         Left            =   5760
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command101 
         Caption         =   "Command3"
         Height          =   495
         Left            =   5880
         TabIndex        =   37
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   5760
         TabIndex        =   36
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   5880
         TabIndex        =   35
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command304 
         Caption         =   "CClose"
         Height          =   495
         Left            =   5640
         TabIndex        =   34
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command303 
         Caption         =   "COpen"
         Height          =   375
         Left            =   5760
         TabIndex        =   33
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Hide"
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
         Left            =   4800
         TabIndex        =   1
         Top             =   5280
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Edit"
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
         Left            =   4800
         TabIndex        =   2
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Device ID"
         Height          =   255
         Left            =   4800
         TabIndex        =   43
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General Multimedia KeyBoard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   980
         TabIndex        =   40
         Top             =   200
         Width           =   4320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   32
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Numeric Keypad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   31
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label404 
         BackStyle       =   0  'Transparent
         Caption         =   $"DirectInput_KeyBoard_Form.frx":0F08
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         TabIndex        =   30
         Top             =   4680
         Width           =   3975
      End
      Begin VB.Label Label209 
         AutoSize        =   -1  'True
         Caption         =   "Default Setting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   29
         Top             =   4440
         Width           =   1290
      End
      Begin VB.OLE OLE1 
         Height          =   495
         Left            =   5520
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CopyRight By General Corporation Bangladesh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   27
         Top             =   840
         Width           =   3960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 3.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "General Multimedia KeyBoard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   25
         Top             =   240
         Width           =   4320
      End
      Begin VB.Label Label900 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9>>>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   24
         Top             =   3960
         Width           =   540
      End
      Begin VB.Label Label800 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8>>>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   23
         Top             =   3720
         Width           =   540
      End
      Begin VB.Label Label700 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7>>>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   22
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label Label600 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6>>>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   21
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label Label500 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5>>>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   20
         Top             =   3000
         Width           =   540
      End
      Begin VB.Label Label400 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4>>>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label300 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3>>>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   18
         Top             =   2520
         Width           =   540
      End
      Begin VB.Label Label200 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2>>>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1>>>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   16
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label text1 
         AutoSize        =   -1  'True
         Caption         =   "No Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   15
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label text2 
         AutoSize        =   -1  'True
         Caption         =   "No Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label text3 
         AutoSize        =   -1  'True
         Caption         =   "No Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   13
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label text4 
         AutoSize        =   -1  'True
         Caption         =   "No Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         Top             =   2760
         Width           =   1005
      End
      Begin VB.Label text5 
         AutoSize        =   -1  'True
         Caption         =   "No Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   3000
         Width           =   1005
      End
      Begin VB.Label text6 
         AutoSize        =   -1  'True
         Caption         =   "No Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   10
         Top             =   3240
         Width           =   1005
      End
      Begin VB.Label text7 
         AutoSize        =   -1  'True
         Caption         =   "No Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   9
         Top             =   3480
         Width           =   1005
      End
      Begin VB.Label text8 
         AutoSize        =   -1  'True
         Caption         =   "No Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   8
         Top             =   3720
         Width           =   1005
      End
      Begin VB.Label text9 
         AutoSize        =   -1  'True
         Caption         =   "No Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   7
         Top             =   3960
         Width           =   1005
      End
   End
   Begin VB.CommandButton Command300 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command200 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text100 
      Height          =   285
      Left            =   8880
      TabIndex        =   3
      Text            =   "10"
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Timer tmrKey 
      Left            =   7560
      Top             =   5280
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   8280
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.exe"
      DialogTitle     =   "General File System"
      Filter          =   "*.exe"
   End
   Begin ComctlLib.Slider HScroll3 
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   5760
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   327682
      LargeChange     =   1
      Max             =   30
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dx As New DirectX7
 Dim di As DirectInput
 Dim diDEV As DirectInputDevice
 Dim diState As DIKEYBOARDSTATE
 Dim iKeyCounter As Integer
Dim dd As Integer
Dim kk As Integer
Dim tnc As Integer


Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
 Const VK_F4 = &H73

  



Private Const HIGHEST_VOLUME_SETTING = 30
    
Private Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Type VolumeSetting
    LeftVol As Integer
    RightVol As Integer
End Type
 
 Private Function lSetVolume(ByRef lLeftVol As Long, ByRef lRightVol As Long, lDeviceID As Long) As Long
    Dim bReturnValue As Boolean
    Dim Volume As VolumeSetting
    Dim lAPIReturnVal As Long
    Dim lBothVolumes As Long
    Volume.LeftVol = nSigned(lLeftVol * 65535 / HIGHEST_VOLUME_SETTING)
    Volume.RightVol = nSigned(lRightVol * 65535 / HIGHEST_VOLUME_SETTING)
    lDataLen = Len(Volume)
    CopyMemory lBothVolumes, Volume.LeftVol, lDataLen
    lAPIReturnVal = waveOutSetVolume(lDeviceID, lBothVolumes)
    lSetVolume = lAPIReturnVal
End Function

Private Function nSigned(ByVal lUnsignedInt As Long) As Integer
   Dim nReturnVal As Integer
   If lUnsignedInt > 65535 Or lUnsignedInt < 0 Then
        MsgBox "Error in conversion from Unsigned to nSigned Integer"
        nSignedInt = 0
        Exit Function
    End If
        If lUnsignedInt > 32767 Then
            nReturnVal = lUnsignedInt - 65536
        Else
            nReturnVal = lUnsignedInt
       End If
            nSigned = nReturnVal
        End Function


 
 
 
                             




Private Sub Combo1_Click()
On Error Resume Next
Kill ("c:\windows\vol.txt")
FileNumber = FreeFile
filename = "c:\windows\vol.txt"
Open filename For Append As #FileNumber
Print #FileNumber, Combo1.Text
Close #FileNumber
Command2_Click
HScroll3.Value = 15
HScroll3_Change
End Sub

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command101_Click()
On Error GoTo error
Dim X, Y As Variant

kk = kk + 1

If kk > (File1.ListCount - 1) Then
kk = 0
End If

X = Drive1.ListCount
Y = X + 65
didar$ = Chr(Y)
Drive1.Drive = didar$
Dir1.Path = Drive1.Drive
Dir1.Path = "\mpegav"
File1.Path = Dir1.Path
OLE1.Delete
OLE1.CreateLink Dir1.Path & "\" & File1.List(kk)
OLE1.DoVerb
Exit Sub
error:
kk = 0
End Sub

Private Sub Command2_Click()
Form1.Visible = False
End Sub

Private Sub Command200_Click()
Text100.Text = Text100 + 1
HScroll3.Value = Text100.Text
Form4.Show
Form4.Width = Text100.Text * 320
HScroll3_Change
End Sub

Private Sub Command3_Click()
On Error Resume Next
MsgBox "Select '0 to 6' And Check Which Device Is Working For Your Volume Control.", 32, "Set Volume Control."
i = Shell("c:\windows\sndvol32.exe", vbNormalFocus)
Combo1.Visible = True
Label7.Visible = True
Command3.Visible = False
End Sub

Private Sub Command300_Click()
Text100.Text = Text100 - 1
HScroll3.Value = Text100.Text
Form4.Show
Form4.Width = Text100.Text * 320
HScroll3_Change
End Sub

Private Sub Command301_Click()
On Error GoTo error
Dim X, Y As Variant


kk = kk - 1

If kk < 0 Then
kk = 0
End If

X = Drive1.ListCount
Y = X + 65
didar$ = Chr(Y)
Drive1.Drive = didar$
Dir1.Path = Drive1.Drive
Dir1.Path = "\mpegav"
File1.Path = Dir1.Path
OLE1.Delete
OLE1.CreateLink Dir1.Path & "\" & File1.List(kk)
OLE1.DoVerb

Exit Sub
error:
kk = 0
End Sub

Private Sub Command303_Click()
SendMCIString "set cd door open", True
Timer2.Interval = 10000
End Sub

Private Sub Command304_Click()
SendMCIString "set cd door closed", True
Timer2.Interval = 7000
kk = -1
End Sub

Private Sub Form_Load()
Dim count As Variant
On Error Resume Next
kk = -1
tnc = 0
    

Set di = dx.DirectInputCreate()
     If err.Number <> 0 Then
         MsgBox "Error starting Direct Input, please make sure you have DirectX installed", vbApplicationModal
         End
     End If

     Set diDEV = di.CreateDevice("GUID_SysKeyboard")

     diDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
     diDEV.SetCooperativeLevel Me.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
     Me.Show
     diDEV.Acquire
     tmrKey.Interval = 10
     tmrKey.Enabled = True
     


     
 Form1.Hide
 
 
 
  SendMCIString "close all", False
If (App.PrevInstance = True) Then
    End
End If
fCDLoaded = False
If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
    End
End If

SendMCIString "set cd time format tmsf wait", True

   
 
     
     
     
     FileNumber = FreeFile
filename = "c:\windows\1.txt"
Open "c:\windows\1.txt" For Input As #FileNumber
text1.Caption = Input(LOF(1), 1)
Close #1
count = (Len(text1.Caption) - 2)
text1.Caption = Left(text1.Caption, count)

  FileNumber = FreeFile
filename = "c:\windows\vol.txt"
Open "c:\windows\vol.txt" For Input As #FileNumber
Combo1.Text = Input(LOF(1), 1)
Close #1
count = (Len(Combo1.Text) - 2)
Combo1.Text = Left(Combo1.Text, count)
If Combo1.Text = "" Then
Combo1.Text = "0"
End If


 FileNumber = FreeFile
filename = "c:\windows\2.txt"
Open "c:\windows\2.txt" For Input As #FileNumber
text2.Caption = Input(LOF(1), 1)
Close #1
count = (Len(text2.Caption) - 2)
text2.Caption = Left(text2.Caption, count)

 FileNumber = FreeFile
filename = "c:\windows\3.txt"
Open "c:\windows\3.txt" For Input As #FileNumber
text3.Caption = Input(LOF(1), 1)
Close #1
count = (Len(text3.Caption) - 2)
text3.Caption = Left(text3.Caption, count)


 FileNumber = FreeFile
filename = "c:\windows\4.txt"
Open "c:\windows\4.txt" For Input As #FileNumber
text4.Caption = Input(LOF(1), 1)
Close #1
count = (Len(text4.Caption) - 2)
text4.Caption = Left(text4.Caption, count)



 FileNumber = FreeFile
filename = "c:\windows\5.txt"
Open "c:\windows\5.txt" For Input As #FileNumber
text5.Caption = Input(LOF(1), 1)
Close #1
count = (Len(text5.Caption) - 2)
text5.Caption = Left(text5.Caption, count)


 FileNumber = FreeFile
filename = "c:\windows\6.txt"
Open "c:\windows\6.txt" For Input As #FileNumber
text6.Caption = Input(LOF(1), 1)
Close #1
count = (Len(text6.Caption) - 2)
text6.Caption = Left(text6.Caption, count)



 FileNumber = FreeFile
filename = "c:\windows\7.txt"
Open "c:\windows\7.txt" For Input As #FileNumber
text7.Caption = Input(LOF(1), 1)
Close #1
count = (Len(text7.Caption) - 2)
text7.Caption = Left(text7.Caption, count)



 FileNumber = FreeFile
filename = "c:\windows\8.txt"
Open "c:\windows\8.txt" For Input As #FileNumber
text8.Caption = Input(LOF(1), 1)
Close #1
count = (Len(text8.Caption) - 2)
text8.Caption = Left(text8.Caption, count)



 FileNumber = FreeFile
filename = "c:\windows\9.txt"
Open "c:\windows\9.txt" For Input As #FileNumber
text9.Caption = Input(LOF(1), 1)
Close #1
count = (Len(text9.Caption) - 2)
text9.Caption = Left(text9.Caption, count)

 FileNumber = FreeFile
filename = "c:\windows\10.txt"
Open "c:\windows\10.txt" For Input As #FileNumber
Text10.Text = Input(LOF(1), 1)
Close #1
count = (Len(Text10.Text) - 2)
Text10.Text = Left(Text10.Text, count)




Exit Sub
err:
End Sub


Private Sub Form_Unload(Cancel As Integer)
 diDEV.Unacquire
End Sub



Private Sub Timer1_Timer()
Form1.Visible = False
 z = lSetVolume(Text100.Text, Text100.Text, Combo1.Text)
Timer1.Interval = 0
End Sub

Private Sub Timer2_Timer()
Dim X, Y As Variant
On Error GoTo err
If Timer2.Interval = 10000 Then
SendMCIString "set cd door closed", True
Timer2.Interval = 0
Else


X = Drive1.ListCount
Y = X + 65
didar$ = Chr(Y)
Drive1.Drive = didar$
Dir1.Path = Drive1.Drive
Dir1.Path = "\mpegav"
Command101_Click
Timer2.Interval = 0
End If
Exit Sub
err:
Timer2.Interval = 0
End Sub


Private Sub tmrKey_Timer()
On Error Resume Next

    diDEV.GetDeviceStateKeyboard diState

     For iKeyCounter = 0 To 255
         If diState.Key(iKeyCounter) <> 0 Then
         
         
         If iKeyCounter = 83 Then
  On Error Resume Next
 i = Shell("c:\windows\rundll32 amovie.ocx,RunDll", vbNormalFocus)
End If
   
   If iKeyCounter = 56 Then
   If Text10.Text = 1 Then
           Call keybd_event(VK_F4, 0, 0, 0)
    End If
    End If
     
         
                 
            
If iKeyCounter = 79 Then
OLE1.Delete
OLE1.CreateLink text1.Caption
OLE1.DoVerb
End If


If iKeyCounter = 69 Then
tnc = Text100.Text
HScroll3.Value = 0
HScroll3_Change
Text100.Text = tnc
End If


     
If iKeyCounter = 80 Then
OLE1.Delete
OLE1.CreateLink text2.Caption
OLE1.DoVerb
End If

If iKeyCounter = 67 Then
Form2.Show
End If


If iKeyCounter = 81 Then
OLE1.Delete
OLE1.CreateLink text3.Caption
OLE1.DoVerb
End If

If iKeyCounter = 75 Then
 OLE1.Delete
OLE1.CreateLink text4.Caption
OLE1.DoVerb
End If

If iKeyCounter = 76 Then
  OLE1.Delete
OLE1.CreateLink text5.Caption
OLE1.DoVerb
End If

If iKeyCounter = 77 Then
 OLE1.Delete
OLE1.CreateLink text6.Caption
OLE1.DoVerb
End If
         
  If iKeyCounter = 71 Then
  OLE1.Delete
OLE1.CreateLink text7.Caption
OLE1.DoVerb
End If



 If iKeyCounter = 72 Then
 OLE1.Delete
OLE1.CreateLink text8.Caption
OLE1.DoVerb
End If


 If iKeyCounter = 73 Then
 OLE1.Delete
OLE1.CreateLink text9.Caption
OLE1.DoVerb
End If
     
 If iKeyCounter = 78 Then
  Command200_Click
  End If
         
 If iKeyCounter = 74 Then
  Command300_Click
  End If
  
  
    If iKeyCounter = 156 Then
  Command301_Click
  End If
      
      
    If iKeyCounter = 68 Then
        tmrKey.Enabled = False
    End If
            
      
         
 If iKeyCounter = 87 Then
 Form5.Show
   End If
  
  If iKeyCounter = 88 Then
  i = Shell("c:\windows\rundll.exe user.exe,exitwindows", vbNormalFocus)
  Form5.Show
 Form5.Label1.Caption = "Shutdown"
   End If
         
 
         
            
 If iKeyCounter = 181 Then
  Command303_Click
  End If
      
              
 If iKeyCounter = 55 Then
  Command304_Click
  End If
 
  If iKeyCounter = 82 Then
  Command101_Click
  End If
 
 
 
            
         End If
     Next
     
     
     
     
     
     
     DoEvents
End Sub


Private Sub HScroll3_Change()
On Error Resume Next
z = lSetVolume(HScroll3.Value, HScroll3.Value, Combo1.Text)
Text100.Text = (HScroll3.Value)
End Sub

Private Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200

rc = mciSendString(cmd, 0, 0, hwnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    MsgBox errStr
End If
SendMCIString = (rc = 0)
End Function

