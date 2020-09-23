VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Programmer: Jason Hensley
'
'Email: elitecobra@hotmail.com
'
'Website: http://www16.brinkster.com/codesource
'
'Version: 2.2 Final
'
'Description: Designed to be simple and easy to use.
'You can easily make a complete mp3, wave, asf,
'midi player, ect. with this Class Module.
'
'
Option Explicit
'Api to access the registry
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Const HKEY_CLASSES_ROOT = &H80000000
'Api to move form without titlebar
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long
'Api to send the commands to the mci device
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'Used to store the Filename
Public FileName As String
'Pause for a specified amount of time
Sub TimeOut(duration)
On Error GoTo TrapIt
    Dim StartTime As Long
    Dim x As Long
    StartTime = Timer
    Do While Timer - StartTime < duration
        x = DoEvents()
    Loop
    Exit Sub
TrapIt:      MsgBox Err.Description, , " Error"
End Sub
'Play a music file
Public Sub Play()
    On Error GoTo TrapIt
    mciSendString "close song", 0, 0, 0 'Close the previous song
    FileName = Chr$(34) + Trim(FileName) + Chr$(34) 'put " and " at the beinning and end of the filename
    mciSendString "open " & FileName & " alias song", 0, 0, 0 'open the filename with the alias named song
    mciSendString "play song", 0, 0, 0 'play the song
    Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Make your app the default app when they open a file
Public Sub MyAppDefault(ByVal sAppName As String, ByVal sEXE As String, ByVal sExt As String)
'Ex. MyAppDefault "My MP3 Player", App.Path & "\" & App.Exe & " %1", ".mp3"
    On Error GoTo TrapIt
    Dim lRegKey As Long
    Call RegCreateKey(HKEY_CLASSES_ROOT, sExt, lRegKey)
    Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sAppName, Len(sAppName))
    Call RegCloseKey(lRegKey)
    Call RegCreateKey(HKEY_CLASSES_ROOT, sAppName & "\Shell\Open\Command", lRegKey) ' adds info into the shell open command
    Call RegSetValueEx(lRegKey, "", 0&, 1, ByVal sEXE, Len(sEXE))
    Call RegCloseKey(lRegKey)
    Exit Sub
TrapIt: MsgBox Err.Description, , " Error"
End Sub
'Used to convert the short filename to the long filename
Public Function GetLongFilename(ByVal sShortName As String) As String
    On Error GoTo TrapIt
     Dim sLongName As String
     Dim sTemp As String
     Dim iSlashPos As Integer

     'Add \ to short name to prevent Instr from failing
     sShortName = sShortName & "\"

     'Start from 4 to ignore the "[Drive Letter]:\" characters
     iSlashPos = InStr(4, sShortName, "\")

     'Pull out each string between \ character for conversion
     While iSlashPos
       sTemp = Dir(Left$(sShortName, iSlashPos - 1), _
         vbNormal + vbHidden + vbSystem + vbDirectory)
       If sTemp = "" Then
         'Error 52 - Bad File Name or Number
         GetLongFilename = ""
         Exit Function
       End If
       sLongName = sLongName & "\" & sTemp
       iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
     Wend

     'Prefix with the drive letter
     GetLongFilename = Left$(sShortName, 2) & sLongName
    Exit Function
TrapIt: MsgBox Err.Description, , " Error"
   End Function
'Check if the song has ended. Good for repeat play, ect.
Public Function EndOfSong() As Boolean
    On Error GoTo TrapIt
    Dim ThePos As Long
    Dim TheDur As Long
    ThePos = GetPositioninSec
    TheDur = GetDurationInSec
    If ThePos = 0 Or TheDur = 0 Then Exit Function
    If ThePos = TheDur Then
        EndOfSong = True
    Else
        EndOfSong = False
    End If
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Returns the current device in use
Public Function DeviceName() As String
    On Error GoTo TrapIt
    Dim TheData As String * 128
    mciSendString "sysinfo song installname", TheData, 128, 0 'retrieve the devicename
    DeviceName = TheData
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Move a borderless form
Public Sub FormMove(Frm As Form)
On Error GoTo TrapIt
    ReleaseCapture
    Call SendMessage(Frm.hWnd, &HA1, 2, 0&)
    Exit Sub
TrapIt: MsgBox Err.Description, , " Error"
End Sub
'Mute just the right channel. True = Mute,False = UnMute
Public Function SetRightChannelOff(Off As Boolean) As Boolean
    On Error GoTo TrapIt
    If Off = True Then
        mciSendString "setaudio song right off", 0, 0, 0
    Else
        mciSendString "setaudio song right on", 0, 0, 0
    End If
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Mute just the left channel. True = Mute,False = UnMute
Public Function SetLeftChannelOff(Off As Boolean) As Boolean
    On Error GoTo TrapIt
    If Off = True Then
        mciSendString "setaudio song left off", 0, 0, 0
        SetLeftChannelOff = True
    Else
        mciSendString "setaudio song left on", 0, 0, 0
        SetLeftChannelOff = False
    End If
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Open and load a .m3u playlist
Public Sub OpenPlaylist(TheList As String, Listbox As Control)
    On Error GoTo TrapIt
    Dim test As String
    If TheList = "" Then Exit Sub
    Open TheList For Input As 1
    While Not EOF(1)
    Line Input #1, test
    Listbox.AddItem RTrim(test)
    Wend
    Close 1
    Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Get the time remaining in milli-seconds
Public Function GetTimeRemaininginMS() As Long
    On Error GoTo TrapIt
    GetTimeRemaininginMS = GetDurationInMS - GetPositioninMS
    Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Get the time remaining in seconds
Public Function GetTimeRemaininginSec() As Long
    On Error GoTo TrapIt
    GetTimeRemaininginSec = GetDurationInSec - GetPositioninSec
    Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Get the value of the playing speed
Public Function GetSpeed() As Long
On Error GoTo TrapIt
    Dim Rate As String * 30
    mciSendString "status song play speed", Rate, 30, 0
    Rate = CLng(Rate)
    GetSpeed = Rate
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Get the time remaining in a user friendly format
Public Function GetFormatTimeRemaining() As String
    On Error GoTo TrapIt
    GetFormatTimeRemaining = GetThisTime(GetTimeRemaininginMS)
    Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Used to find the last backslash of the file path
Private Function GetLastBackSlash(text As String) As String
    On Error GoTo TrapIt
    Dim i, Pos As Integer
    Dim lastslash As Integer
    For i = 1 To Len(text)
        Pos = InStr(i, text, "\", vbTextCompare)
        If Pos <> 0 Then lastslash = Pos
    Next i
    GetLastBackSlash = Right(text, Len(text) - lastslash)
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Used for getting the Right/Left of a string
Private Function RightLeft(source As String, token As String) As String
    On Error GoTo TrapIt
    Dim i As Long
    RightLeft = ""
    For i = Len(source) To 1 Step -1
        If Mid(source, i, 1) = token Then
        RightLeft = Left(source, i - 1)
    Exit Function
    End If
    Next i
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Strip the path and ext. from list2 and add to list1
Public Sub ListNoChar(List1 As Listbox, List2 As Control)
    On Error GoTo TrapIt
    Dim x As Long
    Dim NoChar As String
    Dim NoEnd As String
    For x = 0 To List2.ListCount - 1
        NoChar = GetLastBackSlash(List2.List(x))
        NoEnd = RightLeft(NoChar, ".")
        List1.AddItem NoEnd
    Next x
    Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Strip the path and ext. from a single file
Public Sub ListSingleNoChar(List1, List2)
    On Error GoTo TrapIt
    Dim x As String
    Dim NoChar As String
    Dim NoEnd As String
    x = List2.ListCount - 1
    List2.ListIndex = x
    x = List2.text
    NoChar = GetLastBackSlash(x)
    NoEnd = RightLeft(NoChar, ".")
    List1.AddItem NoEnd
    Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Take the .ext off the end of a file
Private Function NoEndChar(List1 As Listbox, List2 As Listbox) As String
    On Error GoTo TrapIt
    Dim n As Long
    For n = 0 To List2.ListCount - 1
        NoEndChar = Left(List2.List(n), 1)
    Next n
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Save a playlist in the .m3u format
Public Sub SavePlaylist(TheList As String, Listbox As Control)
    On Error GoTo TrapIt
    Dim i As Integer
    Dim a As String
    Open TheList For Output As #1
    For i = 0 To Listbox.ListCount - 1
        a$ = Listbox.List(i)
        Print #1, a$
    Next
    Close 1
    Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Set the left channel to a desired volume
Public Function SetLeftVolume(Value As Long) As Long
'1000 = max | 0 = min
    On Error GoTo TrapIt
    mciSendString "setaudio song left volume to " & Value, 0, 0, 0
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Set the right channel volume
Public Function SetRightVolume(Value As Long) As Long
'1000 = max | 0 = min
    On Error GoTo TrapIt
    mciSendString "setaudio song right volume to " & Value, 0, 0, 0
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Set all audio off(mute)
Public Function SetAudioOff()
    On Error GoTo TrapIt
    mciSendString "set song audio all off", 0, 0, 0
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Set all audio on(Unmute)
Public Function SetAudioOn()
    On Error GoTo TrapIt
    mciSendString "set song audio all on", 0, 0, 0
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Set the speed for the song to be played
Public Function SetSpeed(speed As Long)
    On Error GoTo TrapIt
    mciSendString "set song speed " & speed, "", 0, 0
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Set the volume for both channels
Public Function SetVolume(Value As Long) As Long
'1000 = max | 0 = min
    On Error GoTo TrapIt
    mciSendString "setaudio song volume to " & Value, 0, 0, 0
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function

'Get the position in MilliSeconds
Public Function GetPositioninMS() As Long
    On Error GoTo TrapIt
    Static Stat As String * 128
    mciSendString "set song time format milliseconds", 0, 0, 0
    mciSendString "status song position", Stat, 128, 0
    GetPositioninMS = Val(Stat)
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Get the Position in Seconds
Public Function GetPositioninSec() As Long
    On Error GoTo TrapIt
    GetPositioninSec = Val(GetPositioninMS \ 1000)
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Get the positon in a user friendly format
Public Function GetFormatPosition() As String
    On Error GoTo TrapIt
    Dim Sec As Long
    Dim Mins As Long
    Sec = GetPositioninSec
    If Sec < 60 Then GetFormatPosition = "0:" & Format(Sec, "00")
    If Sec > 59 Then
        Mins = Int(Sec / 60)
        Sec = Sec - (Mins * 60)
        GetFormatPosition = Format(Mins, "0") & ":" & Format(Sec, "00")
    End If
    Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Calculates the bitrate of a mp3 file only
Public Function GetMP3Bitrate() As Long
    On Error Resume Next
    Dim ThePath As String
    Dim TheSize As Long
    'Strip off the quotations from the filename that was added with Chr$(34)
    ThePath = FileName
    TheSize = Len(ThePath) - 1
    ThePath = Right(ThePath, TheSize)
    ThePath = Left(ThePath, Len(ThePath) - 1)
    '
    TheSize = FileLen(ThePath)
    TheSize = TheSize * 8
    GetMP3Bitrate = TheSize \ GetDurationInSec
End Function
'Change the position of the song in seconds
Public Sub ChangePosition(TheSecond As Long)
    On Error GoTo TrapIt
    TheSecond = TheSecond * 1000
    If IsPlaying = True Then mciSendString "play song from " & TheSecond, 0, 0, 0
    If IsPlaying = False Then mciSendString "seek song to " & TheSecond, 0, 0, 0
    Exit Sub
TrapIt:      MsgBox Err.Description, , " Error"
End Sub
'Close the mci device
Public Function CloseAudio()
    On Error GoTo TrapIt
    mciSendString "close all", 0, 0, 0
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Stop playing the song
Public Sub StopPlay()
    On Error GoTo TrapIt
    mciSendString "stop song", 0, 0, 0
    Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Pause the song
Public Sub Pause()
    On Error GoTo TrapIt
    mciSendString "pause song", 0, 0, 0
    Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Get duration in Milliseconds
Public Function GetDurationInMS() As Long
    On Error GoTo TrapIt
    Dim TotalTime As String * 128
    mciSendString "status song length", TotalTime, 128, 0&
    GetDurationInMS = Val(TotalTime)
    Exit Function
TrapIt:          MsgBox Err.Description, , " Error"
End Function
'Get the bitrate in kb instead of bits
Public Function GetFormatMP3Bitrate() As Long
    On Error GoTo TrapIt
    GetFormatMP3Bitrate = GetMP3Bitrate \ 1000
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Get the left channel volume
Public Function GetLeftVolume() As Long
    On Error GoTo TrapIt
    Dim LeftLevel As String * 128
    mciSendString "status song left volume", LeftLevel, 128, 0
    GetLeftVolume = Val(LeftLevel)
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Get the right channel volume value
Public Function GetRightVolume() As Long
    On Error GoTo TrapIt
    Dim Rightlevel As String * 128
    mciSendString "status song right volume", Rightlevel, 128, 0
    GetRightVolume = Val(Rightlevel)
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Get the volume value for both channels
Public Function GetVolume() As Long
    On Error GoTo TrapIt
    Dim TheLevel As String * 128
    mciSendString "status song volume", TheLevel, 128, 0
    GetVolume = Val(TheLevel)
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Get the duration in a user friendly format
Public Function GetFormatDuration() As String
    On Error GoTo TrapIt
    Dim Stat As String * 128
    Dim T As String
    Dim lTotalTime As Long
    mciSendString "set song time format ms", Stat, 128, 0&
    mciSendString "status song length", Stat, 128, 0&
    lTotalTime = Val(Stat)
    T = GetThisTime(lTotalTime)
    GetFormatDuration = T
    Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Used to format the position and duration
Private Function GetThisTime(ByVal timein As Long) As String
    On Error GoTo TrapIt
    Dim conH As Integer
    Dim conM As Integer
    Dim conS As Integer
    Dim remTime As Long
    Dim strRetTime As String
    remTime = timein / 1000
    conH = Int(remTime / 3600)
    remTime = remTime Mod 3600
    conM = Int(remTime / 60)
    remTime = remTime Mod 60
    conS = remTime
    If conH > 0 Then
        strRetTime = Trim(Str(conH)) & ":"
    Else
        strRetTime = ""
    End If
    If conM >= 10 Then
        strRetTime = strRetTime & Trim(Str(conM))
    ElseIf conM > 0 Then
        strRetTime = strRetTime & Trim(Str(conM))
    Else
        strRetTime = strRetTime & "0"
    End If
    strRetTime = strRetTime & ":"
    If conS >= 10 Then
        strRetTime = strRetTime & Trim(Str(conS))
    ElseIf conS > 0 Then
        strRetTime = strRetTime & "0" & Trim(Str(conS))
    Else
        strRetTime = strRetTime & "00"
    End If
    GetThisTime = strRetTime
    Exit Function
TrapIt:      MsgBox Err.Description, , " Error"
End Function
'Get and save setting to the registry
Private Sub GetandSaveSettings()
'You can make these function in the module yourself if you want them.
'These function are built into vb(VB 6.0 i've tested with)
'These are good for saving and loading the location of a playlist or something. For a autoload feature, ect.
'This is how you can save/get a setting to/from the registry
    'Dim a as string
    'Dim b as string
    'a = SaveSetting("Your Player Here", "What the setting is for", "Location", Form1.C.FileName)
    'b = GetSetting("Your Player Here", "What the setting was for", "Value")
End Sub
'Get the Duration in Seconds
Public Function GetDurationInSec() As Long
    On Error GoTo TrapIt
    GetDurationInSec = GetDurationInMS \ 1000
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Resume a paused song
Public Sub ResumePlay()
    On Error GoTo TrapIt
    mciSendString "resume song", 0, 0, 0
    Exit Sub
TrapIt:  MsgBox Err.Description, , " Error"
End Sub
'Check the status to see if the music is playing, stopped, ect.
Public Function GetStatus() As String
    On Error GoTo TrapIt
    Dim Status As String * 128
    mciSendString "status song mode", Status, 128, 0
    GetStatus = Status
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
'Check to see if a song is playing
Public Function IsPlaying() As Boolean
    On Error GoTo TrapIt
    Dim Stat As String * 7
    mciSendString "status song mode", Stat, 128, 0
    If Stat = "playing" Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
    Exit Function
TrapIt:  MsgBox Err.Description, , " Error"
End Function
