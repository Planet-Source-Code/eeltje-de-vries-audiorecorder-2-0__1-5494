VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   2760
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdMidi 
      Caption         =   "Choose midi file to record"
      Height          =   375
      Left            =   1920
      TabIndex        =   24
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton cmdOke 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame Frame6 
      Caption         =   "Recoding options"
      Height          =   3495
      Left            =   1920
      TabIndex        =   12
      Top             =   120
      Width           =   4815
      Begin VB.OptionButton optRecordImmediate 
         Caption         =   "Manual recording"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optRecordProgrammed 
         Caption         =   "Programmed recording"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   2055
      End
      Begin VB.Frame frmTimes 
         Caption         =   "Enter times"
         Height          =   1575
         Left            =   2520
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdStartTime 
            Caption         =   "Start time"
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdStopTime 
            Caption         =   "Stop time"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.Frame frmManualAuto 
         Caption         =   "Saving file"
         Height          =   1695
         Left            =   2520
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdFileName 
            Caption         =   "Filename"
            Height          =   375
            Left            =   360
            TabIndex        =   22
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Manual"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Automatic"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label lblTimes 
         Caption         =   " "
         Height          =   1215
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resolution"
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
      Begin VB.OptionButton opt8bits 
         Caption         =   "8 bits"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton opt16bits 
         Caption         =   "16 bits"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Channels"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
      Begin VB.OptionButton optMono 
         Caption         =   "mono"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optStereo 
         Caption         =   "stereo"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sample rate (Hz)"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton optRate6000 
         Caption         =   "6000"
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton optRate8000 
         Caption         =   "8000"
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   1260
         Width           =   1095
      End
      Begin VB.OptionButton optRate11025 
         Caption         =   "11025"
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optRate22050 
         Caption         =   "22050"
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   660
         Width           =   1095
      End
      Begin VB.OptionButton optRate44100 
         Caption         =   "44100"
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFileName_Click()
    WaveFileName = InputBox("Filename: ", "Filename for automatic saving", WaveFileName)
End Sub

Private Sub cmdMidi_Click()
    CommonDialog2.CancelError = True
    On Error GoTo ErrHandler1
    CommonDialog2.Filter = "Midi file (*.mid*)|*.mid"
    CommonDialog2.Flags = &H2 Or &H400
    CommonDialog2.ShowOpen
    WaveMidiFileName = CommonDialog2.FileName
    WaveMidiFileName = GetShortName(WaveMidiFileName)
ErrHandler1:
End Sub

Private Sub cmdOke_Click()
    Unload Me
End Sub

Private Sub cmdStartTime_Click()
    Dim wrst As String
    wrst = WaveRecordingStartTime
    wrst = InputBox("Enter start time recording", "Start time", wrst)
    If wrst = "" Then Exit Sub
    If Not IsDate(wrst) Then
        MsgBox ("The date/time you entered was not valid!")
    Else
    ' String returned from InputBox is a valid time,
    ' so store it as a date/time value in WaveRecordingStartTime.
        If CDate(wrst) < Now Then
            MsgBox ("Recording events in the past is not possible...")
            WaveRecordingStartTime = Now + TimeSerial(0, 15, 0)
        Else
            WaveRecordingStartTime = CDate(wrst)
        End If
        If WaveRecordingStopTime < WaveRecordingStartTime Then WaveRecordingStopTime = WaveRecordingStartTime + TimeSerial(0, 15, 0)
    End If
End Sub

Private Sub cmdStopTime_Click()
    Dim wrst As String
    
    wrst = WaveRecordingStopTime
    If wrst < WaveRecordingStartTime Then wrst = WaveRecordingStartTime + TimeSerial(0, 15, 0)
        
    wrst = InputBox("Enter stop time recording", "Stop time", wrst)
    If wrst = "" Then Exit Sub
    If Not IsDate(wrst) Then
        MsgBox ("The time you entered was not valid!")
    Else
    ' String returned from InputBox is a valid time,
    ' so store it as a date/time value in WaveRecordingStartTime.
        If CDate(wrst) < WaveRecordingStartTime Then
            MsgBox ("The stop time has to be later then the start time!")
            WaveRecordingStopTime = WaveRecordingStartTime + TimeSerial(0, 5, 0)
        Else
            WaveRecordingStopTime = CDate(wrst)
        End If
    End If
End Sub

Private Sub Form_Load()
    Select Case Rate
    Case 44100
        optRate44100.Value = True
    Case 22050
        optRate22050.Value = True
    Case 11025
        optRate11025.Value = True
    Case 8000
        optRate8000.Value = True
    Case 6000
        optRate6000.Value = True
    End Select
    
    Select Case Channels
    Case 1
        optMono.Value = True
    Case 2
        optStereo.Value = True
    End Select
    
    Select Case Resolution
    Case 8
        opt8bits.Value = True
    Case 16
        opt16bits.Value = True
    End Select
    
    If WaveRecordingImmediate Then
        optRecordImmediate.Value = True
    Else
        optRecordProgrammed.Value = True
    End If
    
    If WaveAutomaticSave Then
        Option11.Value = True
    Else
        Option10.Value = True
    End If
          
End Sub

Private Sub optRate11025_Click()
    Rate = 11025
    optRate11025.Value = True
End Sub

Private Sub optRate44100_Click()
    Rate = 44100
    optRate44100.Value = True
End Sub

Private Sub Option10_Click()
    WaveAutomaticSave = False
End Sub

Private Sub Option11_Click()
    WaveAutomaticSave = True
End Sub

Private Sub optRate22050_Click()
    Rate = 22050
    optRate22050.Value = True
End Sub


Private Sub optRate8000_Click()
    Rate = 8000
    optRate8000.Value = True
End Sub

Private Sub optRate6000_Click()
    Rate = 6000
    optRate6000.Value = True
End Sub

Private Sub optMono_Click()
    Channels = 1
    optMono.Value = True
End Sub

Private Sub optStereo_Click()
    Channels = 2
    optStereo.Value = True
End Sub

Private Sub opt8bits_Click()
    Resolution = 8
    opt8bits.Value = True
End Sub

Private Sub opt16bits_Click()
    Resolution = 16
    opt16bits.Value = True
End Sub

Private Sub optRecordImmediate_Click()
    WaveRecordingImmediate = True
    frmManualAuto.Visible = False
    frmTimes.Visible = False
    lblTimes.Visible = False
    AudioRecorder.cmdRecord.Enabled = True
End Sub

Private Sub optRecordProgrammed_Click()
    WaveRecordingImmediate = False
    frmManualAuto.Visible = True
    frmTimes.Visible = True
    lblTimes.Visible = True
    AudioRecorder.cmdRecord.Enabled = False
    If WaveRecordingStartTime < Now Then
        WaveRecordingStartTime = Now + TimeSerial(0, 15, 0)
        WaveRecordingStopTime = WaveRecordingStartTime + TimeSerial(0, 15, 0)
    End If

End Sub

