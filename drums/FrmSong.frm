VERSION 5.00
Begin VB.Form FrmSong 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3795
   ClientLeft      =   1935
   ClientTop       =   1635
   ClientWidth     =   5415
   Icon            =   "FrmSong.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   253
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra1 
      Height          =   2130
      Left            =   3645
      TabIndex        =   14
      Top             =   30
      Width           =   1635
      Begin VB.CommandButton CmdUp 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   795
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   1365
         Width           =   705
      End
      Begin VB.CommandButton CmdDown 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         Top             =   1365
         Width           =   645
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   165
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Top             =   1695
         Width           =   1335
      End
      Begin VB.CommandButton CmdReplace 
         Caption         =   "R&eplace"
         Height          =   375
         Left            =   165
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   1005
         Width           =   1335
      End
      Begin VB.VScrollBar VsbSongPtrn 
         Height          =   195
         Left            =   1200
         Max             =   1
         Min             =   100
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   285
         Value           =   1
         Width           =   240
      End
      Begin VB.CommandButton CmdInsert 
         Caption         =   "&Insert"
         Height          =   375
         Left            =   795
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   645
         Width           =   705
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   165
         OLEDropMode     =   1  'Manual
         TabIndex        =   0
         Top             =   645
         Width           =   645
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Pattern:"
         Height          =   195
         Index           =   2
         Left            =   165
         LinkTimeout     =   0
         TabIndex        =   17
         Top             =   285
         Width           =   555
      End
      Begin VB.Label LblSongPtrn 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "001"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   840
         LinkTimeout     =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Top             =   255
         Width           =   630
      End
   End
   Begin VB.CommandButton btnSave 
      Enabled         =   0   'False
      Height          =   360
      Left            =   4920
      Picture         =   "FrmSong.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Save Wave"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   14.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Play Wave"
      Top             =   3120
      Width           =   435
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "g"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Stop Wave"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton btnRec 
      BackColor       =   &H000000FF&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Record Wave"
      Top             =   3120
      Width           =   435
   End
   Begin VB.Timer tmrStatus 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   3540
   End
   Begin VB.Frame Fra2 
      Height          =   1065
      Left            =   3645
      TabIndex        =   18
      Top             =   1995
      Width           =   1635
      Begin VB.CheckBox ChkLoop 
         Caption         =   "&Loop Song"
         Height          =   195
         Left            =   165
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Top             =   255
         Width           =   1260
      End
      Begin VB.CommandButton CmdSongStop 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   825
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         Top             =   585
         Width           =   705
      End
      Begin VB.CommandButton CmdSongRun 
         Caption         =   "&Play"
         Height          =   375
         Left            =   105
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   585
         Width           =   705
      End
   End
   Begin VB.ListBox LstSong 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   3375
      ItemData        =   "FrmSong.frx":0090
      Left            =   120
      List            =   "FrmSong.frx":0092
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Top             =   120
      Width           =   3390
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   3540
      Width           =   5415
   End
End
Attribute VB_Name = "FrmSong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Author:        Stuart Pennington.
' Project:       Midi Percussion Sequencer (Drum Machine).
' Test Platform: Windows 98SE
' Processor:     P2 300MHz.

Const Msg1 As String = "Part Number: "
Const Msg2 As String = " ---> Pattern Number: "

Private bSongPlaying As Boolean

Private Sub Form_Load()
  Dim C%

  Icon = FrmMicroKit.Icon
  Caption = "MicroKit - Song Mode"

  ' Populate The List Box With Song Data (If There Is Any).
  For C = 0 To 99
    If Song(C) > 0 Then
       LstSong.AddItem "Part Number: " & Format(C + 1, "000") & " ---> Pattern Number: " & Format(Song(C), "000")
    End If
  Next

  ' Did The User Want The Song To Loop?
  ChkLoop.Value = Abs(bLoopSong)

  ' Position Form.
  PositionForm Me
End Sub

Private Sub Form_Unload(Cancel%)
  If bSongPlaying Then CmdSongStop_Click
End Sub

Private Sub btnPlay_Click()
  btnStop_Click
  WavePlay
End Sub

Private Sub btnRec_Click()
  btnStop.Enabled = True
  btnPlay.Enabled = False
  tmrStatus.Enabled = True
  WaveRecord
  tmrStatus_Timer
  CmdSongRun_Click
End Sub

Private Sub btnStop_Click()
  If tmrStatus.Enabled Then
    tmrStatus.Enabled = False
    btnStop.Enabled = False
    btnPlay.Enabled = True
    btnSave.Enabled = True
    WaveStop
    tmrStatus_Timer
  End If
End Sub

Private Sub btnSave_Click()
  Dim filename As String, path As String
  
  path = App.path
  If Right(path, 1) <> "\" Then path = path & "\"
  
  filename = SaveDialog(Me, "Wave Audio (*.wav)|*.wav", "Save As", path, "audio.wav")
  If Len(filename) Then
    WaveSaveAs filename
  End If
End Sub

Private Sub CmdSongRun_Click()
  ' If The List Box Is Empty Then There's Nothing To Play.
  If LstSong.ListCount = 0 Then
    Beep
    Exit Sub
  End If

  ' Update Interface.
  CmdSongRun.Enabled = False
  CmdSongStop.Enabled = True
  CmdSongStop.SetFocus
  VsbSongPtrn.Enabled = False
  CmdAdd.Enabled = False
  CmdInsert.Enabled = False
  CmdReplace.Enabled = False
  CmdDelete.Enabled = False
  CmdUp.Enabled = False
  CmdDown.Enabled = False
  
  btnRec.Enabled = False
  DoEvents

  ' Indicate That SONG Is Playing.
  bSongPlaying = True
  ' Initialize Song Monitor Variables.
  SongStartPos = (Song(0) - 1) * 16
  SongEndPos = SongStartPos + 15
  SongPos = SongStartPos
  SongPtr = 0

  ' Create High Res (Multi Media) Timer To Do The Sequencing.
  TimerID = timeSetEvent(NewTempo, BestResolution, AddressOf SongProc, 0, 1)
  ' We Are Now Sequencing The Song Patterns.
End Sub

Public Sub CmdSongStop_Click()
  ' Kill High Res Timer.
  timeKillEvent TimerID
  ' Return Timer ID To Zero.
  TimerID = 0
  ' Silence Any Midi Note That May Have "Stuck".
  midiOutReset hMidiOut
  ' Indicate That The Song Is Nolonger Playing.
  bSongPlaying = False

  ' Update Interface.
  VsbSongPtrn.Enabled = True
  CmdAdd.Enabled = True
  CmdInsert.Enabled = True
  CmdReplace.Enabled = True
  CmdDelete.Enabled = True
  CmdUp.Enabled = True
  CmdDown.Enabled = True
  CmdSongStop.Enabled = False
  CmdSongRun.Enabled = True
  CmdSongRun.SetFocus
  btnRec.Enabled = True
End Sub

Private Sub CmdAdd_Click()
  AddItem Insert:=False
End Sub

Private Sub CmdInsert_Click()
  AddItem Insert:=True
End Sub

Private Sub CmdUp_Click()
  Dim msg As String
  With LstSong
    If .ListIndex > 0 Then
      msg = .List(.ListIndex - 1)
      .List(.ListIndex - 1) = Msg1 & Format(.ListIndex, "000") & Msg2 & Right(.List(.ListIndex), 3)
      .List(.ListIndex) = Msg1 & Format(.ListIndex + 1, "000") & Msg2 & Right(msg, 3)
      .Selected(.ListIndex - 1) = True
    End If
  End With
End Sub

Private Sub CmdDown_Click()
  Dim msg As String
  With LstSong
    If .ListIndex < .ListCount - 1 Then
      msg = .List(.ListIndex + 1)
      .List(.ListIndex + 1) = Msg1 & Format(.ListIndex + 2, "000") & Msg2 & Right(.List(.ListIndex), 3)
      .List(.ListIndex) = Msg1 & Format(.ListIndex + 1, "000") & Msg2 & Right(msg, 3)
      .Selected(.ListIndex + 1) = True
    End If
  End With
End Sub

Private Sub AddItem(Insert As Boolean)
  ' Purpose: Adds A Pattern Number To The List Box And
  '          Updates The Song Playback Array.

  Dim msg$, i As Integer

  With LstSong
    If .ListCount = 100 Then
      msg = "Sorry, A maximum of 100 entries has been reached."
      MsgBox msg, vbExclamation, "MicroKit - Error"
      Exit Sub
    End If
    If Insert And .ListCount > 0 Then
      .AddItem Msg1 & Format(.ListIndex + 1, "000") & Msg2 & LblSongPtrn.Caption, .ListIndex
      For i = .ListIndex To .ListCount - 1
        .List(i) = Msg1 & Format(i + 1, "000") & Msg2 & Right(.List(i), 3)
      Next
      .Selected(.ListIndex) = True
    Else
      .AddItem Msg1 & Format(.ListCount + 1, "000") & Msg2 & LblSongPtrn.Caption
      .Selected(.ListCount - 1) = True
    End If
  End With

  UpdateSongArray
End Sub

Private Sub UpdateSongArray()
  ' Purpose: Updates The Song PlayBack Array.
  Dim K%

  Erase Song
  
  For K = 0 To LstSong.ListCount - 1
      Song(K) = Val(Right(LstSong.List(K), 3)) ' Fill It With The Pattern Number's.
  Next
  
  ' Data Has Changed.
  bDirty = True
End Sub

Private Sub CmdReplace_Click()
  ' Purpose: Allows A User To Change Any Pattern Number For Another
  '          Anywhere In The List Box And Updates The Song Playback Array.
  
  With LstSong
       ' If The List Box Is Empty There's Nothing To Do.
       If .ListCount = 0 Or .Text = "" Then
           Beep
           Exit Sub
       End If

       ' Replace The Pattern Number For The New Number Yhe User Has Chosen.
      .List(.ListIndex) = "Part Number: " & Format(.ListIndex + 1, "000") & " ---> Pattern Number: " & LblSongPtrn.Caption
  End With

  UpdateSongArray
End Sub

Private Sub CmdDelete_Click()
  ' Purpose: Removes A Pattern From The List Box And Resorts It's Content's.

  Dim K%, S%, E%

  With LstSong
     If .ListCount = 0 Or .Text = "" Then
         Beep
         Exit Sub
     End If

     S = .ListIndex + 1
     E = .ListCount
     If S < E Then
        S = S - 1
        E = E - 1
        For K = S To E
           .List(K) = Left(.List(K), 22) & Right(.List(K + 1), 19)
        Next
     End If

    .RemoveItem .ListCount - 1
  End With

  ' Update Song Play Back Array.
  UpdateSongArray
End Sub

Private Sub tmrStatus_Timer()
  lblStatus = WaveStatus & ": " & WaveStatistics
End Sub

Private Sub VsbSongPtrn_Change()
  LblSongPtrn.Caption = Format(VsbSongPtrn.Value, "000")
End Sub

Private Sub ChkLoop_Click()
  ' Toggle Song Looping.
  bLoopSong = CBool(ChkLoop.Value)
  bDirty = True
End Sub

