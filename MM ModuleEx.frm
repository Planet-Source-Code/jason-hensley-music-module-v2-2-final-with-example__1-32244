VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   " MM Module Example"
   ClientHeight    =   5175
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Mute"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   2760
      TabIndex        =   23
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox P3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   1920
      ScaleHeight     =   135
      ScaleMode       =   0  'User
      ScaleWidth      =   2000
      TabIndex        =   20
      Top             =   2040
      Width           =   1575
      Begin VB.Label L3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   720
         TabIndex        =   21
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.PictureBox P2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   240
      ScaleHeight     =   135
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   17
      Top             =   2040
      Width           =   1575
      Begin VB.Label L2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1200
         TabIndex        =   18
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   240
      ScaleHeight     =   135
      ScaleMode       =   0  'User
      ScaleWidth      =   0.931
      TabIndex        =   11
      Top             =   2520
      Width           =   3255
      Begin VB.Label L1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   135
      End
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   1680
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   2640
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   1590
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   3255
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label cSong 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Playing: "
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Play Rate: 100%"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Volume: 76%"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Position\Played: 0%"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Remaining:"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Position:"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0:00\00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFC0&
      X1              =   120
      X2              =   3600
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFC0&
      X1              =   120
      X2              =   3600
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Music Module Example"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Load"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pause"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stop"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Play"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFC0&
      Height          =   3375
      Left            =   120
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0:00\00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFC0&
      Height          =   1095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFC0&
      BorderWidth     =   3
      FillStyle       =   0  'Solid
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MM As New MusicModule

Private Sub Check1_Click()
    If Check1.Value = Checked Then
        MM.SetAudioOff
    Else
        MM.SetAudioOn
    End If
End Sub

Private Sub cSong_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cSong.ToolTipText = Mid(cSong.Caption, 10)
End Sub

Private Sub Form_Load()
    Dim TheFile As String
    
    'Center the form
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
    
    'set the default value for the volume bar and speed rate bar
    P2.CurrentX = 76
    P3.CurrentX = 1000 'is the normal speed
    
    'I put this here in case you make your player the default player.
    'example
    'MM.MyAppDefault "MM ModuleEx", "c:\MM ModuleEx.exe" & " %1", ".mp3"
    
    
    'The following will load the associated file in list1, list2 and play
    
    'check if assoc. file was loaded
    If Command$ = "" Then Exit Sub
    
    'the long songname looks better ;)
    TheFile = MM.GetLongFilename(Command$)
    
    'add the file to the playlist
    List2.AddItem TheFile
    
    'strip off the path and ext. and put in list1
    MM.ListSingleNoChar List1, List2
    
    'click play
    Label2_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Close all audio from the mci device
    MM.CloseAudio
End Sub

Private Sub Label2_Click()
    On Error Resume Next
'If a song is not selected then exit sub
    'If List2.text = "" Then MsgBox "Please select a song to play!", , "Error": Exit Sub
'Load the filename of the song
    MM.FileName = List2

'Play the song
    MM.Play
'Set the current playing volume and playing speed
    MM.SetVolume P2.CurrentX * 10
    MM.SetSpeed P3.CurrentX
'Make sure the song has had enough time to load
'to get status information
    'MM.TimeOut 0.5
'Load the duration in seconds
    P1.ScaleWidth = MM.GetDurationInSec
    Timer1.Enabled = True
'Load the current song playing
    cSong.Caption = "Playing: " & List1

End Sub

Private Sub Label3_Click()
'Stop the song from playing
    MM.StopPlay
    Timer1.Enabled = False
End Sub

Private Sub Label4_Click()
'Pause and resume the song
    With Label4
    If .Caption = "Pause" Then
        .Caption = "Resume"
        MM.Pause
    Else
        .Caption = "Pause"
        MM.ResumePlay
    End If
    End With
End Sub

Private Sub Label5_Click()
'Open and load a .m3u playlist
    C.Filter = "M3U Playlist (*.m3u)|*.m3u|MP3 Files (*.mp3)|*.mp3|Wave Files (*.wav)|*.wav|Midi Files (*.mid)|*.mid|All Files (*.*)|*.*"
    C.ShowOpen
    If C.FileName = "" Then Exit Sub
    If C.FileName = " " Then Exit Sub
    If LCase(Right(C.FileName, 3)) = LCase("m3u") Then
        List1.Clear
        List2.Clear
        Call MM.OpenPlaylist(C.FileName, List2)
        Call MM.ListNoChar(List1, List2)
    Else
        List2.AddItem C.FileName
        Call MM.ListSingleNoChar(List1, List2)
    End If
    C.FileName = ""
End Sub

Private Sub Label6_Click()
'Save a .m3u playlist
    C.Filter = "M3U Playlist (*.m3u)|*.m3u"
    C.ShowSave
    If C.FileName = "" Then Exit Sub
    If C.FileName = " " Then Exit Sub

    Call MM.SavePlaylist(C.FileName, List2)

    C.FileName = ""
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Move the form without a title bar
    MM.FormMove Me
End Sub

Private Sub Label8_Click()
'Minimize the form
    Me.WindowState = 1
End Sub

Private Sub Label9_Click()
    Unload Me
    End
End Sub

Private Sub List1_Click()
    List2.ListIndex = List1.ListIndex
End Sub

Private Sub List1_DblClick()
    List2.ListIndex = List1.ListIndex
    Label2_Click
End Sub

Private Sub P1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Will change the current playing position
    P1.CurrentX = X
    L1.Left = P1.CurrentX
    MM.ChangePosition P1.CurrentX
End Sub
Private Sub P2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Label15.Left = P3.ScaleX
P2.CurrentX = X
    L2.Left = P2.CurrentX
    MM.SetVolume (P2.CurrentX * 10)
    Label14.Caption = "Volume: " & Int(P2.CurrentX) & "%"
End Sub

Private Sub P3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Label15.Left = P3.ScaleX
P3.CurrentX = X
    L3.Left = P3.CurrentX
    MM.SetSpeed (P3.CurrentX + 3)
    'Label16.Caption = "Speed Rate: " & Int(P3.CurrentX + 3) & "%"
    Label16.Caption = "Play Rate: " & Int((P3.CurrentX / 2000) * 200) + 3 & "%"
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    'Check to see if a song is playing
    If MM.IsPlaying = False Then Exit Sub
    'Label1 will have the position
    Label1.Caption = MM.GetFormatPosition & "\" & MM.GetFormatDuration
    'Label10 will have the time left
    Label10.Caption = MM.GetFormatTimeRemaining & "\" & MM.GetFormatDuration
    'Will keep the progress bar at the current position
    P1.CurrentX = MM.GetPositioninSec
    L1.Left = P1.CurrentX
    Label13.Caption = "Position\Played: " & Int((P1.CurrentX / MM.GetDurationInSec) * 100) & "%"
    'Check to see if the song has ended for our continuous play
    If MM.EndOfSong = True Then
    If List1.ListCount = 1 Then
        Exit Sub
    Else
        List1.ListIndex = Val(List1.ListIndex) + 1
        Label2_Click
    End If
    End If
End Sub
