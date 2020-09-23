VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WavePlay by Nick Gisburne - nick@gisburne.com / www.gisburne.com"
   ClientHeight    =   3300
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameBalance 
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6195
      TabIndex        =   12
      Top             =   2355
      Width           =   1515
      Begin MSComctlLib.Slider SlideBalance 
         Height          =   285
         Left            =   30
         TabIndex        =   13
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   10
         Min             =   -100
         Max             =   100
         TickFrequency   =   25
         TextPosition    =   1
      End
   End
   Begin VB.Frame frameVolume 
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6195
      TabIndex        =   10
      Top             =   1680
      Width           =   1515
      Begin MSComctlLib.Slider SlideVolume 
         Height          =   285
         Left            =   30
         TabIndex        =   11
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         TickFrequency   =   10
         TextPosition    =   1
      End
   End
   Begin WavePlay.TimeEdit100 TimeDisplay 
      Height          =   315
      Left            =   4635
      TabIndex        =   9
      Top             =   2655
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
   End
   Begin VB.Frame frameSpeed 
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   6195
      TabIndex        =   6
      Top             =   525
      Width           =   1515
      Begin VB.CommandButton PlaySpeed 
         Caption         =   "Change"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1275
      End
      Begin MSComctlLib.Slider SlideSpeed 
         Height          =   285
         Left            =   30
         TabIndex        =   7
         Top             =   660
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   10
         Min             =   10
         Max             =   100
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
         TextPosition    =   1
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8355
      Top             =   2730
   End
   Begin VB.PictureBox ScrollFrame 
      AutoRedraw      =   -1  'True
      Height          =   2040
      Left            =   30
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   404
      TabIndex        =   2
      Top             =   30
      Width           =   6120
      Begin VB.PictureBox PicWave 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00008000&
         ForeColor       =   &H00000000&
         Height          =   1920
         Left            =   30
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   400
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   6000
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   30
      TabIndex        =   1
      Top             =   3060
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open WAV File"
      Height          =   420
      Left            =   6195
      TabIndex        =   0
      Top             =   45
      Width           =   1485
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   8355
      Top             =   2055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LabStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status Here"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   2820
      Width           =   840
   End
   Begin MediaPlayerCtl.MediaPlayer MP 
      Height          =   660
      Left            =   60
      TabIndex        =   4
      Top             =   2115
      Width           =   6135
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -290
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------
' WavePlay
'----------------------------------------------------------
' Author : Nick Gisburne
' Email  : nick@gisburne.com
' Web    : www.gisburne.com / www.karaokebuilder.com
'----------------------------------------------------------
' Purpose:
' Displaying a waveform in time with audio
' Ability to scroll along the whole length of the waveform
' Intelligent cursor (watch it at the start/end)
'----------------------------------------------------------
' Limitations:
' Only for 44100, 16-bit stereo waveforms, but you should
' be able to adapt it to 8-bit and other sample rates, etc
' I only needed it to play one type of waveform, sorry!
'----------------------------------------------------------
' I am not actually drawing a 'wave' as such, just a series
' of min/max lines. However, I wanted the display to look
' like the waveforms in the majority of sound editors
' (Cool Edit by preference) and it does that pretty well
'
' The wave is scanned (quickly!) and the min/max values
' for each section of the wave are found. One value is
' stored for each 1/100th of a second. These are drawn as
' the audio plays.
'
' There are some other bits (speed, volume etc) you might or
' might not find useful. All I can say is, they are useful
' to me! This code is part of a bigger project which will
' not be open-source. Having used a lot of ideas from other
' coders and thought it only fair to share this much.
'----------------------------------------------------------
' PS The TimeEdit control is adapted from my commercial
' product, Karaoke Builder. Have fun with all my source!
'----------------------------------------------------------

Option Explicit

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Type WAVEFMT
    signature As String * 4     ' must contain 'RIFF'
    RIFFsize As Long            ' size of file (in bytes) minus 8
    type As String * 4          ' must contain 'WAVE'
    fmtchunk As String * 4      ' must contain 'fmt ' (including blank)
    fmtsize As Long             ' size of format chunk, must be 16
    format As Integer           ' normally 1 (PCM)
    channels As Integer         ' number of channels, 1=mono, 2=stereo
    samplerate As Long          ' sampling frequency: 11025, 22050 or 44100
    average_bps As Long         ' average bytes per second; samplerate * channels
    align As Integer            ' 1=byte aligned, 2=word aligned
    bitspersample As Integer    ' should be 8 or 16
    datchunk As String * 4      ' must contain 'data'
    samples As Long             ' number of samples
End Type
 
Private Type POINT
    x As Long
    y As Long
End Type
 
'----------------------------------------------------------
' 882 comes from:
'   44100 = 1 second
'   441   = 1/100 second
'   882   = 441 * 2 (2 = stereo)
'----------------------------------------------------------
Private Type WAVEBLOCK
    wavinfo(1 To 882) As Integer
End Type
 
Private Type SCROLLER
    min As Long
    max As Long
    value As Long
    topval As Long  'Position represented by top line of wavform
End Type
 
Dim Fil As String
Dim WavCount As Long, WavMin() As Integer, WavMax() As Integer
Dim ttt As WAVEBLOCK, LastPt As POINT, VScroll As SCROLLER
Dim WavSpeed As Integer



Private Sub cmdOpen_Click()
    LoadWaveForm
End Sub

Private Sub Command1_Click()
    MP.CurrentPosition = MP.CurrentPosition + 0.01
End Sub

Private Sub Form_Load()
    LabStatus = ""
    WavSpeed = 50   'Default value for Speed Change button
    SlideSpeed.value = 100
    SlideVolume.value = 100
    SlideBalance.value = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MP.FileName = ""
End Sub


Private Sub PicWave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MP.CurrentPosition = (VScroll.topval + x) / 100
    VScroll.value = VScroll.topval + x
End Sub

Private Sub LoadWaveForm()
    Dim t1 As Long, tmin As Integer, tmax As Integer
    Dim wavh As WAVEFMT
    
    With CM
        .FileName = ""
        .Filter = "Wav(*.wav)|*.wav"
        .ShowOpen
        Fil = .FileName
    End With
    Me.Refresh  'Do this to remove lingering dialog box lines
    If Fil = "" Then Exit Sub
    If Dir(Fil, vbNormal) = "" Then Exit Sub
    
    MP.FileName = Fil

    '-----------------------------------------
    '  Only allow wavh.bitspersample = 16!!!
    '-----------------------------------------
    Open Fil For Binary Access Read As #1
    Get #1, , wavh
    WavCount = wavh.samples \ Len(ttt)
    
    ReDim WavMin(WavCount), WavMax(WavCount)
    
    Screen.MousePointer = vbHourglass

    ProgressBar1.max = WavCount
    ProgressBar1.value = 0
    ProgressBar1.Visible = True
    
    WavCount = -1
    Do
        Get #1, , ttt
        'Find min/max to get a better view of the wave
        tmin = 0: tmax = 0
        'Should look at all values but 'Step 32' speeds it up a bit
        'A nice C++ function to find the max/min would be handy here!
        For t1 = 1 To UBound(ttt.wavinfo) Step 32
            tmin = IIf(ttt.wavinfo(t1) < tmin, ttt.wavinfo(t1), tmin)
            tmax = IIf(ttt.wavinfo(t1) > tmax, ttt.wavinfo(t1), tmax)
        Next t1
        WavCount = WavCount + 1
        WavMin(WavCount) = tmin \ 512 + 64  'Long values become +/-64 then
        WavMax(WavCount) = tmax \ 512 + 64  'add 64 to make into co-ordinates
        
        If WavCount Mod 100 = 0 Then ProgressBar1.value = WavCount
    Loop Until EOF(1)
    
    Close #1
    ProgressBar1.Visible = False
    ProgressBar1.value = 0

    VScroll.min = 0
    VScroll.max = WavCount
    VScroll.value = 0
            
    TimeDisplay.SetVal 0, 0, WavCount   'Parameters: value, min, max
    
    DrawWaveData
    Screen.MousePointer = vbDefault
End Sub

Private Sub PlaySpeed_Click()
    SlideSpeed.value = IIf(SlideSpeed.value = 100, WavSpeed, 100)
End Sub

Private Sub SlideBalance_Change()
    MP.Balance = SlideBalance.value * 100   'Balance is +/- 10000
End Sub

Private Sub SlideBalance_Scroll()
    SlideBalance_Change
End Sub

Private Sub SlideSpeed_Change()
    If SlideSpeed.value <> 100 Then WavSpeed = SlideSpeed.value
    frameSpeed.Caption = "Speed:" & SlideSpeed.value & "%"
End Sub

Private Sub SlideSpeed_Scroll()
    SlideSpeed_Change
End Sub

Private Sub SlideVolume_Change()
    frameVolume.Caption = "Volume:" & SlideVolume.value & "%"
End Sub

Private Sub SlideVolume_Scroll()
    SlideVolume_Change
End Sub

Private Sub TimeDisplay_Change(newPacks As Long)
    MP.CurrentPosition = newPacks / 100
End Sub

Private Sub Timer1_Timer()
    Dim mpos As Long, mvol As Long
    If MP.Rate <> SlideSpeed.value / 100 Then MP.Rate = SlideSpeed.value / 100
    
    mvol = -(100 - SlideVolume.value) * 100     'Media Player volume = 0 (max) to -10000 (min)
    If MP.Volume <> mvol Then MP.Volume = mvol
    
    mpos = Int(MP.CurrentPosition * 100)
    If mpos >= VScroll.min And mpos <= VScroll.max And VScroll.value <> mpos Then
        VScroll.value = mpos
        DrawWaveData
    End If
End Sub


'-------------------------------------------------
' Main drawing routine
'-------------------------------------------------
' MoveToEx and LineTo are GDI functions, many
' times faster than their VB equivalents and
' just as easy to use
'
' Picture containing the wave form is 400 wide
' so change this code to suit your needs
'-------------------------------------------------
Sub DrawWaveData()
    Dim t1 As Long, vstart As Long, hline As Integer
    vstart = VScroll.value
    
    'Make sure the cursor starts at the left, moves to the middle,
    'then the waveform moves until at the end the cursor moves right
    If vstart < 200 Or WavCount < 400 Then      'First 200
        hline = vstart
        vstart = 0
    ElseIf vstart > WavCount - 200 Then         'Last 200
        hline = vstart - WavCount + 400
        vstart = WavCount - 400
    Else                                        'Anything else
        hline = 200
        vstart = vstart - 200
    End If
    VScroll.topval = vstart
    
    PicWave.Cls
    PicWave.ForeColor = vbGreen
    
    'Draw each line
    For t1 = 0 To IIf(WavCount < 400, WavCount, 400)
        MoveToEx PicWave.hdc, t1, WavMin(vstart + t1), LastPt
        LineTo PicWave.hdc, t1, WavMax(vstart + t1)
        
        'Marks for 0.1 (small) and 1-second (large) intervals
        If (vstart + t1) Mod 10 = 0 Then
            MoveToEx PicWave.hdc, t1, 128, LastPt
            LineTo PicWave.hdc, t1, IIf((vstart + t1) Mod 100 = 0, 100, 120)
        End If
    Next t1
    
    MoveToEx PicWave.hdc, 0, 64, LastPt 'Draw center line
    LineTo PicWave.hdc, 400, 64
    
    PicWave.ForeColor = vbRed           'Draw cursor line
    MoveToEx PicWave.hdc, hline, 0, LastPt
    LineTo PicWave.hdc, hline, 500
    
    PicWave.Refresh
    
    LabStatus.Caption = VScroll.value & " of " & WavCount & " (" & WavCount / 100 & " seconds)"
    TimeDisplay.SetVal VScroll.value
End Sub


