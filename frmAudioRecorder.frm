VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AudioRecorder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AudioRecorder"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmAudioRecorder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Settings"
      Height          =   495
      Left            =   5970
      TabIndex        =   10
      ToolTipText     =   "Change rate, stereo/mono, 8/16 bits and program an automatic recording"
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "You can choose a beginning for playing the recording"
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   500
      SmallChange     =   100
      TickStyle       =   3
   End
   Begin VB.CommandButton cmdWeb 
      Caption         =   "Web"
      Height          =   495
      Left            =   4995
      TabIndex        =   7
      ToolTipText     =   "Visit the home page of me!! (Maybe a new version is available...)"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "To start a new recording and adjusting all settings"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4020
      TabIndex        =   3
      ToolTipText     =   "Save the recording as as WAV file"
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   " "
      Orientation     =   2
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3045
      TabIndex        =   2
      ToolTipText     =   "Play the recording"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2070
      TabIndex        =   1
      ToolTipText     =   "Stop recording or playing"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record"
      Height          =   495
      Left            =   1095
      TabIndex        =   0
      ToolTipText     =   "Start recording immediate"
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Starting position for play (in milliseconds)"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   4815
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   5160
      Top             =   2400
   End
   Begin VB.Frame Frame4 
      Caption         =   "Statistics"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   4815
      Begin VB.Label StatisticsLabel 
         BackColor       =   &H00000000&
         Caption         =   " "
         ForeColor       =   &H0000FF00&
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Information about the recording"
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "AudioRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright: E. de Vries
'e-mail: eeltje@geocities.com
'This code can be used as freeware

Const AppName = "AudioRecorder"

Private Sub cmdSave_Click()
    Dim sName As String
    
    If WaveMidiFileName = "" Then
        sName = "Radio_from_" & CStr(WaveRecordingStartTime) & "_to_" & CStr(WaveRecordingStopTime)
        sName = Replace(sName, ":", "-")
        sName = Replace(sName, " ", "_")
        sName = Replace(sName, "/", "-")
    Else
        sName = WaveMidiFileName
        sName = Replace(sName, "MID", "wav")
    End If
  
    CommonDialog1.FileName = sName
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler1
    CommonDialog1.Filter = "WAV file (*.wav*)|*.wav"
    CommonDialog1.Flags = &H2 Or &H400
    CommonDialog1.ShowSave
    sName = CommonDialog1.FileName
    
    WaveSaveAs (sName)
    Exit Sub
ErrHandler1:
End Sub

Private Sub cmdRecord_Click()
    Dim settings As String
    Dim Alignment As Integer
      
    Alignment = Channels * Resolution / 8
    
    settings = "set capture alignment " & CStr(Alignment) & " bitspersample " & CStr(Resolution) & " samplespersec " & CStr(Rate) & " channels " & CStr(Channels) & " bytespersec " & CStr(Alignment * Rate)
    WaveReset
    WaveSet
    WaveRecord
    WaveRecordingStartTime = Now
    cmdStop.Enabled = True   'Enable the STOP BUTTON
    cmdPlay.Enabled = False  'Disable the "PLAY" button
    cmdSave.Enabled = False  'Disable the "SAVE AS" button
    cmdRecord.Enabled = False 'Disable the "RECORD" button
End Sub

Private Sub cmdSettings_Click()
Dim strWhat As String
    ' show the user entry form modally
    strWhat = MsgBox("If you continue your data will be lost!", vbOKCancel)
    If strWhat = vbCancel Then
        Exit Sub
    End If
    Slider1.Max = 10
    Slider1.Value = 0
    Slider1.Refresh
    cmdRecord.Enabled = True
    cmdStop.Enabled = False
    cmdPlay.Enabled = False
    cmdSave.Enabled = False
    
    WaveReset
    
    Rate = CLng(GetSetting("AudioRecorder", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("AudioRecorder", "StartUp", "Channels", "1"))
    Resolution = CInt(GetSetting("AudioRecorder", "StartUp", "Resolution", "16"))
    WaveFileName = GetSetting("AudioRecorder", "StartUp", "WaveFileName", "C:\Radio.wav")
    WaveAutomaticSave = GetSetting("AudioRecorder", "StartUp", "WaveAutomaticSave", "True")

    WaveRecordingImmediate = True
    WaveRecordingReady = False
    WaveRecording = False
    WavePlaying = False
    
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    frmSettings.optRecordImmediate.Value = True
    frmSettings.Show vbModal
End Sub

Private Sub cmdStop_Click()
    WaveStop
    cmdSave.Enabled = True  'Enable the "SAVE AS" button
    cmdPlay.Enabled = True  'Enable the "PLAY" button
    cmdStop.Enabled = False 'Disable the "STOP" button
    If WavePosition = 0 Then
        Slider1.Max = 10
    Else
        If WaveRecordingImmediate And (Not WavePlaying) Then Slider1.Max = WavePosition
        If (Not WaveRecordingImmediate) And WaveRecording Then Slider1.Max = WavePosition
    End If
    If WaveRecording Then WaveRecordingReady = True
    WaveRecordingStopTime = Now
    WaveRecording = False
    WavePlaying = False
    frmSettings.optRecordProgrammed.Value = False
    frmSettings.optRecordImmediate.Value = True
    frmSettings.lblTimes.Visible = False
End Sub

Private Sub cmdPlay_Click()
    WavePlayFrom (Slider1.Value)
    WavePlaying = True
    cmdStop.Enabled = True
    cmdPlay.Enabled = False
End Sub


Private Sub cmdWeb_Click()
  Dim ret&
  ret& = ShellExecute(Me.hwnd, "Open", "http://home.wxs.nl/~eeltjevr/", "", App.Path, 1)
End Sub




Private Sub cmdReset_Click()
    Slider1.Max = 10
    Slider1.Value = 0
    Slider1.Refresh
    cmdRecord.Enabled = True
    cmdStop.Enabled = False
    cmdPlay.Enabled = False
    cmdSave.Enabled = False
    
    WaveReset
    
    Rate = CLng(GetSetting("AudioRecorder", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("AudioRecorder", "StartUp", "Channels", "1"))
    Resolution = CInt(GetSetting("AudioRecorder", "StartUp", "Resolution", "16"))
    WaveFileName = GetSetting("AudioRecorder", "StartUp", "WaveFileName", "C:\Radio.wav")
    WaveAutomaticSave = GetSetting("AudioRecorder", "StartUp", "WaveAutomaticSave", "True")

    WaveRecordingImmediate = True
    WaveRecordingReady = False
    WaveRecording = False
    WavePlaying = False
    WaveMidiFileName = ""
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    If WaveRenameNecessary Then
        Name WaveShortFileName As WaveLongFileName
        WaveRenameNecessary = False
        WaveShortFileName = ""
    End If
End Sub

Private Sub Form_Load()
    WaveReset
    
    Rate = CLng(GetSetting("AudioRecorder", "StartUp", "Rate", "110025"))
    Channels = CInt(GetSetting("AudioRecorder", "StartUp", "Channels", "1"))
    Resolution = CInt(GetSetting("AudioRecorder", "StartUp", "Resolution", "16"))
    WaveFileName = GetSetting("AudioRecorder", "StartUp", "WaveFileName", "C:\Radio.wav")
    WaveAutomaticSave = GetSetting("AudioRecorder", "StartUp", "WaveAutomaticSave", "True")

    WaveRecordingImmediate = True
    WaveRecordingReady = False
    WaveRecording = False
    WavePlaying = False
    
    'Be sure to change the Value property of the appropriate button!!
    'if you change the default values!
    
    WaveSet
    WaveRecordingStartTime = Now + TimeSerial(0, 15, 0)
    WaveRecordingStopTime = WaveRecordingStartTime + TimeSerial(0, 15, 0)
    WaveMidiFileName = ""
    WaveRenameNecessary = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WaveClose
    Call SaveSetting("AudioRecorder", "StartUp", "Rate", CStr(Rate))
    Call SaveSetting("AudioRecorder", "StartUp", "Channels", CStr(Channels))
    Call SaveSetting("AudioRecorder", "StartUp", "Resolution", CStr(Resolution))
    Call SaveSetting("AudioRecorder", "StartUp", "WaveFileName", WaveFileName)
    Call SaveSetting("AudioRecorder", "StartUp", "WaveAutomaticSave", CStr(WaveAutomaticSave))
    If WaveRenameNecessary Then
        Name WaveShortFileName As WaveLongFileName
        WaveRenameNecessary = False
        WaveShortFileName = ""
    End If
    End
End Sub


Private Sub Timer2_Timer()
    Dim RecordingTimes As String
    Dim msg As String
    
    RecordingTimes = "Start time:  " & WaveRecordingStartTime & vbCrLf _
                    & "Stop time:  " & WaveRecordingStopTime
    
    WaveStatistics
    If Not WaveRecordingImmediate Then
        WaveStatisticsMsg = WaveStatisticsMsg & "Programmed recording"
        If WaveAutomaticSave Then
            WaveStatisticsMsg = WaveStatisticsMsg & " (automatic save)"
        Else
            WaveStatisticsMsg = WaveStatisticsMsg & " (manual save)"
        End If
        WaveStatisticsMsg = WaveStatisticsMsg & vbCrLf & vbCrLf & RecordingTimes
    End If
    StatisticsLabel.Caption = WaveStatisticsMsg
    
    WaveStatus
    If WaveStatusMsg <> AudioRecorder.Caption Then AudioRecorder.Caption = WaveStatusMsg
    If InStr(AudioRecorder.Caption, "stopped") > 0 Then
        cmdStop.Enabled = False
        cmdPlay.Enabled = True
    End If
    
    If RecordingTimes <> frmSettings.lblTimes.Caption Then frmSettings.lblTimes.Caption = RecordingTimes
    
    If (Now > WaveRecordingStartTime) _
            And (Not WaveRecordingReady) _
            And (Not WaveRecordingImmediate) _
            And (Not WaveRecording) Then
        WaveReset
        WaveSet
        WaveRecord
        WaveRecording = True
        cmdStop.Enabled = True   'Enable the STOP BUTTON
        cmdPlay.Enabled = False  'Disable the "PLAY" button
        cmdSave.Enabled = False  'Disable the "SAVE AS" button
        cmdRecord.Enabled = False 'Disable the "RECORD" button
    End If
    
    If (Now > WaveRecordingStopTime) And (Not WaveRecordingReady) And (Not WaveRecordingImmediate) Then
        WaveStop
        cmdSave.Enabled = True 'Enable the "SAVE AS" button
        cmdPlay.Enabled = True 'Enable the "PLAY" button
        cmdStop.Enabled = False 'Disable the "STOP" button
        If WavePosition > 0 Then
            Slider1.Max = WavePosition
        Else
            Slider1.Max = 10
        End If
        WaveRecording = False
        WaveRecordingReady = True
        If WaveAutomaticSave Then
            WaveFileName = "Radio_from_" & CStr(WaveRecordingStartTime) & "_to_" & CStr(WaveRecordingStopTime)
            WaveFileName = Replace(WaveFileName, ":", ".")
            WaveFileName = Replace(WaveFileName, " ", "_")
            WaveFileName = WaveFileName & ".wav"
            WaveSaveAs (WaveFileName)
            msg = "Recording has been saved" & vbCrLf
            msg = msg & "Filename: " & WaveFileName
            MsgBox (msg)
        Else
            msg = "Recording is ready" & vbCrLf
            msg = msg & "Don't forget to save recording..."
            MsgBox (msg)
        End If
        frmSettings.optRecordProgrammed.Value = False
        frmSettings.optRecordImmediate.Value = True
    End If

End Sub
