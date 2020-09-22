Attribute VB_Name = "ModMicroKit"
Option Explicit

' Author:        Stuart Pennington.
' Project:       Midi Percussion Sequencer (Drum Machine).
' Test Platform: Windows 98SE
' Processor:     P2 300MHz.



' Note 1: Sound Quality Determined By Soundcard's Midi Spec.
' Note 2: App Uses CALLBACKS So Shut Down From The Main Form, NOT VB IDE.
' Note 3: Unable To Sync To Other Midi Instruments.



' Midi Msg Example:
'
' Middle C Note-On On Ch9 With A Volume Of 127:-
'
' Midi Msg Format:   VOLUME :: Note Number :: Note On :: Channel's (0 To 15)
'                    &H7F   :: &H30        :: &H90    :: &H9                  = &H7F3099
'
' Middle C Note-Off On Ch9:-
'
' Midi Msg Format:   VOLUME :: Note Number :: Note Off :: Channel's (0 To 15)
'                    &H0    :: &H30        :: &H80     :: &H9                 =  &H3089



Public bDirty As Boolean             ' Indicates That Data Has Changed.
Public bFileSaved As Boolean         ' Inicates Success Or Failure Of A File Save.
Public bPatternChanged As Boolean    ' Indicates (In Pattern Mode) That The Pattern Number Has Changed.
Public bLoopSong As Boolean          ' Indicates (In Song Mode) If The Song Should Loop.

Public hMidiOut&         ' Handle Of Midi Output Device.
Public OldTempo&         ' Tempo Monitor.
Public NewTempo&         ' The Tempo.
Public TimerID&          ' Timer ID.
Public BestResolution&   ' Minimum Firing Interval (MilliSecond's).

Public StartPos%, EndPos%, CurPos%
Public NewStartPos%, NewEndPos%
Public SongPos%, SongStartPos%, SongEndPos%
Public SongPtr%

' Application Defined Structure Used For Playback, Editing etc...
Type Kit
     bNoteOn As Boolean
     PercName As Integer
     PercVol As Integer
     MidiMsg As Long
End Type

' Massive Array Of 100, 16 Step Pattern,s (16 Track... Rows Are Tracks).
Public Ptrns(15, 1599) As Kit
' Used To Playback Pattern Sequences In Song Mode.
Public Song%(99)

' Used For Determining What Operating We're Running On.
Type OSVERSIONINFO
     dwOSVersionInfoSize As Long
     dwMajorVersion As Long
     dwMinorVersion As Long
     dwBuildNumber As Long
     dwPlatformId As Long
     szCSDVersion As String * 128
End Type

' Used For Determining The Capabilities Of Multimedia Timer.
Type TIMECAPS
     wPeriodMin As Long
     wPeriodMax As Long
End Type

Type RECT
     rLeft As Long
     rTop As Long
     rRight As Long
     rBottom As Long
End Type

' General Api Function's.
Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal X1&, ByVal Y1&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Declare Function CreateRectRgn& Lib "gdi32" (ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&)
Declare Function DeleteObject& Lib "gdi32" (ByVal hObject&)
Declare Function DrawEdge& Lib "user32" (ByVal ahdc&, qrc As RECT, ByVal edge&, ByVal grfFlags&)
Declare Function DrawIconEx& Lib "user32" (ByVal ahdc&, ByVal xLeft&, ByVal yTop&, ByVal hIcon&, ByVal cxWidth&, ByVal cyWidth&, ByVal istepIfAniCur&, ByVal hbrFlickerFreeDraw&, ByVal diFlags&)
Declare Function GetFileTitle& Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile$, ByVal lpszTitle$, ByVal cbBuf&)
Declare Function GetShortPathName& Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath$, ByVal lpszShortPath$, ByVal cchBuffer&)
Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO)
Declare Function PtInRect& Lib "user32" (lpRect As RECT, ByVal ptx&, ByVal pty&)
Declare Function SetRect& Lib "user32" (lpRect As RECT, ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&)
Declare Function SetWindowRgn& Lib "user32" (ByVal ahWnd&, ByVal hRgn&, ByVal bRedraw&)
Declare Function TextOut& Lib "gdi32" Alias "TextOutA" (ByVal ahdc&, ByVal X1&, ByVal Y1&, ByVal lpString$, ByVal nCount&)

' Midi Api Function's.
Declare Function midiOutClose& Lib "winmm.dll" (ByVal hMidiOut&)
Declare Function midiOutGetNumDevs& Lib "winmm.dll" ()
Declare Function midiOutOpen& Lib "winmm.dll" (lphMidiOut&, ByVal uDeviceID&, ByVal dwCallback&, ByVal dwInstance&, ByVal dwFlags&)
Declare Function midiOutReset& Lib "winmm.dll" (ByVal hMidiOut&)
Declare Function midiOutShortMsg& Lib "winmm.dll" (ByVal hMidiOut&, ByVal dwMsg&)

' Multi Media Timer Api Function's.
Declare Function timeBeginPeriod& Lib "winmm.dll" (ByVal uPeriod&)
Declare Function timeEndPeriod& Lib "winmm.dll" (ByVal uPeriod&)
Declare Function timeGetDevCaps& Lib "winmm.dll" (lpTimeCaps As TIMECAPS, ByVal uSize&)
Declare Function timeKillEvent& Lib "winmm.dll" (ByVal uId&)
Declare Function timeSetEvent& Lib "winmm.dll" (ByVal uDelay&, ByVal uResolution&, ByVal lpFunction&, ByVal dwUser&, ByVal uFlags&)

Public Const MIDI_MAPPER = -1&
Public Const Ttl = "MicroKit"
Private Sub Main()

    #If Win32 Then

        If Not OpSysOk Then End           ' Check Out The Operating System.
        If PrevInst Then End              ' Check If We Are Already Running.
        If Not MidiOutDevs() Then End     ' Check Midi Out Device Availability.
        If Not HighResTimer Then End      ' Check If We Can Create A High Resolution Multi Media Timer.
        If Not DeviceOpen Then End        ' Check We If We Can Open The Midi Output Device.
        
        ' Set It Up So User's Can Double On Our File Icons To Launch The Exe.
        SetUpIconDblClick

        ' Show The Application.
        FrmMicroKit.Show
    #Else
        End
    #End If

End Sub
Private Function OpSysOk() As Boolean

    Dim Rv&
    Dim OSVI As OSVERSIONINFO

    OSVI.dwOSVersionInfoSize = 148
    Rv = GetVersionEx(OSVI)
    If Rv = 0 Then
       ' Unable To Determine Operating System, End For Safety.
       OpSysOk = False
       Exit Function
    End If

    Select Case OSVI.dwPlatformId
           Case 0
                ' Win32s.
                OpSysOk = False
           Case 1
                ' Win 95/98.
                OpSysOk = True
           Case 2
                ' WinNT (2000 Only).
                If OSVI.dwMajorVersion > 4 Then OpSysOk = True Else OpSysOk = False
    End Select

End Function
Private Function PrevInst() As Boolean

    Dim msg$

    ' Midi Is A Non-Shareable Resource, If We're Already Running Then Quit.

    If App.PrevInstance Then
       msg = "MicroKit is already running."
       MsgBox msg, vbInformation, "MicroKit"
       PrevInst = True
    Else
       PrevInst = False
    End If

End Function
Private Function MidiOutDevs() As Boolean

     Dim msg$

     If midiOutGetNumDevs() = 0 Then
        ' No Midi Devices On System.
        msg = "Unable to detect a midi output device."
        msg = msg & vbCrLf & vbCrLf
        msg = msg & "Terminating..."
        MsgBox msg, vbCritical, "MicroKit - Error"
        MidiOutDevs = False
    Else
        ' Found Midi Device/s.
        MidiOutDevs = True
    End If

End Function
Private Function HighResTimer() As Boolean

    ' =====================================================
    ' Possibly The Bit You've been Waiting For
    ' If You've Found The VB Timer Just Ain't Up To It.
    ' =====================================================


    Dim msg$, Rv&
    Dim TC As TIMECAPS

    ' Get The Capabilities Of A Multi Media Timer.
    Rv = timeGetDevCaps(TC, Len(TC))
    If Rv = 0 Then
       ' Find The Best Resolution Available. (Most Likely One Millisecond).
       BestResolution = TC.wPeriodMin
       ' Set That Resolution.
       ' Note: Every "timeBeginPeriod" Must Be Followed (Somewhere In The Code) By a "timeEndPeriod".
       Rv = timeBeginPeriod(BestResolution)
       If Rv = 0 Then
          ' We Have Access To Our High Res Timer.
          HighResTimer = True
       Else
          ' Can't Create It, Must End.
          GoTo TimerError
       End If
    Else
       ' Couldn't Determine Capabilities, Must End.
       GoTo TimerError
    End If

    Exit Function

TimerError:

    msg = "Unable to create high resolution timer."
    msg = msg & vbCrLf & vbCrLf
    msg = msg & "Terminating..."
    MsgBox msg, vbCritical, "MicroKit - Error"
    HighResTimer = False
    Exit Function

End Function
Private Function DeviceOpen() As Boolean
    Dim msg$, Rv&

    ' Try Opening Midi Output Device.
    ' If Successful, The "hMidiOut" Variable Will Contain It's Handle After The Call.
    Rv = midiOutOpen(hMidiOut, MIDI_MAPPER, 0, 0, 0)
    If Rv <> 0 Then
       ' Failed To Open Device.
       msg = "Unable to open a midi output device."
       msg = msg & vbCrLf & vbCrLf
       msg = msg & "Terminating..."
       MsgBox msg, vbCritical, "MicroKit - Error"
       DeviceOpen = False
    Else
       ' Success...
       ' Send Program Change To Midi Device Requesting Midi Channel Ten
       ' (Logical Channel Nine)... That's Where The Percussion Is.
       midiOutShortMsg hMidiOut, &HC9  ' &HC0 = Program-Change-Msg, &H9 = Channel.
       ' Device Is Open.
       DeviceOpen = True
       ' (See Form Unload For How To Close It).
    End If

End Function
Public Sub PatternProc(ByVal uId&, ByVal uMsg&, ByVal dwUser&, ByVal dw1&, ByVal dw2&)

    ' ======================================================
    ' Callback procedure for high res timer. (Pattern Mode).
    '
    ' The Heart Of The Pattern Sequencer.
    ' ======================================================
    

    Dim CurrentRow%  ' Counter.

    For CurrentRow = 0 To 15
        ' Play Each Instrument On All 16 Tracks At The Current Pattern Array ROW Position.
        If Ptrns(CurrentRow, CurPos).bNoteOn Then
           ' Play The Percussion.
           midiOutShortMsg hMidiOut, Ptrns(CurrentRow, CurPos).MidiMsg
           ' Drum Sounds Are One-Shot Samples So We Can
           ' Send A Note Off Midi Message Immediately.
           midiOutShortMsg hMidiOut, (CLng(Ptrns(CurrentRow, CurPos).PercName) * &H100) + &H89
        End If
    Next

    ' Move To The Next Row In The Pattern Array.
    CurPos = CurPos + 1
    If CurPos > EndPos Then
       ' See If The User Has Changed The Pattern No.
       If bPatternChanged Then
          bPatternChanged = False
          ' Set New Start And End Pointers.
          StartPos = NewStartPos
          EndPos = NewEndPos
          CurPos = NewStartPos
       Else
          ' Go To Start Of Current Pattern Again (Loop).
          CurPos = StartPos
       End If
    End If

    ' Check For Tempo Changes.
    If NewTempo <> OldTempo Then
       ' Tempo Has Changed.
       ' Kill This Timer And Create A New One With New Firing Interval.
       timeKillEvent uId
       TimerID = timeSetEvent(NewTempo, BestResolution, AddressOf PatternProc, 0, 1)
       OldTempo = NewTempo
    End If

End Sub
Public Sub SongProc(ByVal uId&, ByVal uMsg&, ByVal dwUser&, ByVal dw1&, ByVal dw2&)

    ' ===================================================
    ' Callback procedure for high res timer. (Song Mode).
    '
    ' The Heart Of The Song Sequencer.
    ' ===================================================


    Dim CurrentRow%   ' Counter.

    For CurrentRow = 0 To 15
        ' Play Each Instrument On All 16 Tracks At The Current Pattern Array ROW Position.
        If Ptrns(CurrentRow, SongPos).bNoteOn Then
           ' Play The Percussion.
           midiOutShortMsg hMidiOut, Ptrns(CurrentRow, SongPos).MidiMsg
           ' Drum Sounds Are One-Shot Samples So We Can
           ' Send A Note Off Midi Message Immediately.
           midiOutShortMsg hMidiOut, (CLng(Ptrns(CurrentRow, SongPos).PercName) * &H100) + &H89
        End If
    Next

    ' Update The Song Position.
    SongPos = SongPos + 1
    If SongPos > SongEndPos Then
       ' Update The Song Pointer.
       SongPtr = SongPtr + 1
       If SongPtr > 99 Or Song(SongPtr) = 0 Then
          If bLoopSong Then
             SongStartPos = (Song(0) - 1) * 16
             SongEndPos = SongStartPos + 15
             SongPos = SongStartPos
             SongPtr = 0
          Else
             FrmSong.CmdSongStop_Click
             Exit Sub
          End If
       End If
       SongStartPos = (Song(SongPtr) - 1) * 16
       SongEndPos = SongStartPos + 15
       SongPos = SongStartPos
    End If

End Sub
Public Function GetTitle$(FileNameIn$)

     ' Purpose: Returns The File Title From A Full Path.

     Dim Buffer$, Pos%

     Buffer = Space(260)
     GetFileTitle FileNameIn, Buffer, 260
     Pos = InStr(Buffer, vbNullChar)
     GetTitle = StrConv(Left(Buffer, Pos - 1), vbProperCase)

End Function
Public Function GetShortPath$(FileNameIn$)

    ' Purpose: Converts A Path Name To Dos Format.

    Dim Buffer$, Pos%

    Buffer = Space(260)
    GetShortPathName FileNameIn, Buffer, 260
    Pos = InStr(Buffer, vbNullChar)
    GetShortPath = Left(Buffer, Pos - 1)

End Function
Public Sub PositionForm(Frm As Form)

    With Frm
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 3
    End With

End Sub
