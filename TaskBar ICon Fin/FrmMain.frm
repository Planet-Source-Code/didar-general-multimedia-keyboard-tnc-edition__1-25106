VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   1935
   ClientTop       =   -2550
   ClientWidth     =   90
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   525
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1515
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Menu MnuMain 
      Caption         =   ""
      NegotiatePosition=   1  'Left
      Visible         =   0   'False
      Begin VB.Menu MnuMainShow 
         Caption         =   "&CDOpen"
      End
      Begin VB.Menu MnuMainHide 
         Caption         =   "&CDClose"
      End
      Begin VB.Menu MnuMainS1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMainNext 
         Caption         =   "&Active Movie"
      End
      Begin VB.Menu MnuMainEna 
         Caption         =   "Enable KeyBoard"
      End
      Begin VB.Menu MnuMainDis 
         Caption         =   "Disable Keyboard"
      End
      Begin VB.Menu MnuMainOption 
         Caption         =   "Option"
      End
      Begin VB.Menu MnuMainS3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMainShut 
         Caption         =   "ShutDown"
      End
      Begin VB.Menu MnuMainRes 
         Caption         =   "Restart"
      End
      Begin VB.Menu MnuMainS2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMainClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
     
      
      
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
      





Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

   With abd
      .cbSize = Len(abd)
      .hwnd = Me.hwnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = Mouse_Move
      .hIcon = Me.Icon
      .szTip = App.Title & ".  General Corporation Bangladesh." & vbNullChar
   End With
   Shell_NotifyIcon NIM_ADD, abd





SendMCIString "close all", False
cmdClose.Visible = False
If (App.PrevInstance = True) Then
    End
End If
fCDLoaded = False
If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
    End
End If
SendMCIString "set cd time format tmsf wait", True
Form1.Show
FileCopy (App.Path & "\taskshot.exe"), ("c:\windows\taskshot.exe")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
     msg = X
Else
     msg = X / Screen.TwipsPerPixelX
End If

Select Case msg
             Case Mouse_Right_Down
            Case Mouse_Right_Click
            Me.PopupMenu MnuMain
           
            Case Mouse_Right_DbClick
            Case Mouse_Left_Down
            Case Mouse_Left_Click
            Case Mouse_Left_DbClick
            Me.WindowState = vbNormal
            Me.Show
            
            Case Mouse_Button_Down
           Case Mouse_Button_Click
            Case Mouse_Button_DbClick
            NextMpeg
            
    End Select
        







End Sub


Private Sub Form_Unload(Cancel As Integer)
SendMCIString "close all", False
Shell_NotifyIcon NIM_DELETE, abd
End: End
End Sub

Sub NextMpeg()
Dim CountM As Long
CountM = ListFiles.ListCount - 1
NFile = NFile + 1
If NFile > CountM Then NFile = 0
End Sub

Private Sub MnuMainClose_Click()
CmdExit_Click
End Sub

Private Sub MnuMainHide_Click()
On Error Resume Next
cmdclose_click
End Sub
Private Sub MnuMainOption_Click()
On Error Resume Next
Form1.Timer1.Enabled = False
Form1.Frame1.Height = 6095
Form1.Height = 6320
Form1.Show
End Sub

Private Sub MnuMainShut_Click()
On Error Resume Next
i = Shell("c:\windows\rundll.exe user.exe,exitwindows", vbNormalFocus)
End Sub
Private Sub MnuMainRes_Click()
On Error Resume Next
i = Shell("c:\windows\rundll.exe user.exe,exitwindowsexec", vbNormalFocus)
End Sub









Private Sub MnuMainNext_Click()
On Error Resume Next
 i = Shell("c:\windows\rundll32 amovie.ocx,RunDll", vbNormalFocus)
 End Sub

Private Sub MnuMainShow_Click()
On Error Resume Next
cmdopen_Click
End Sub
Private Sub MnuMainCdOpen_Click()
On Error Resume Next
cmdclose_click
End Sub

Private Sub MnuMaindis_Click()
On Error Resume Next
Form1.tmrKey.Enabled = False
End Sub

Private Sub MnuMainena_Click()
On Error Resume Next
Form1.tmrKey.Enabled = True
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
Private Sub cmdopen_Click()
SendMCIString "set cd door open", True
cmdClose.Visible = True
cmdClose.Default = True
End Sub
Private Sub cmdclose_click()
SendMCIString "set cd door closed", True
cmdClose.Visible = False
cmdClose.Default = False
cmdOpen.Default = True
End Sub

'This Program Is Dedicated To Tahmina Nur Chowdhury
'I Love Her More Than Anything Else In The
'World.
'Tahmina I'm Leaving Bangladesh Very Soon. Please Atleast
'come back once again. Please..
'Still Loving You...
'I Feel You Every Second..
'Please Come Back..
'Tahmina Why You Are So Far From Me?
'Why?
'What I Did?
'Just I Love You.. For This Why?
'O--No---...
'If You Don't Want To Love Me, It's Your Wish.But I Love You,Without You.
'I Just Want To Love You..
'Tahmina Please I'll Go...
'I Want To Hear Your Voice..
'I Want To See Your Face..
'I Want To See Your Foot Steps.
'Tahmina Please Come Back. At Least Once Again.
'Tahmina How Much Time I'll Wait For You?
'Tahmina I Love You..
'Still I Searching In Everywhere..
'I Love You, Love You, And Love You...

