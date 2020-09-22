Attribute VB_Name = "ModFileIO"
Option Explicit

' Author:        Stuart Pennington.
' Project:       Midi Percussion Sequencer (Drum Machine).
' Test Platform: Windows 98SE
' Processor:     P2 300MHz.



Public NewFileName$     ' File Name From Open Or SaveAs Dialog Boxes.
Public NewFileTitle$    ' File Title From Open Or SaveAs Dialog Boxes.

' API Common Dialog Initialization Structure (Open/SaveAs).
Type OPENFILENAME
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     lpstrFilter As String
     lpstrCustomFilter As String
     nMaxCustFilter As Long
     nFilterIndex As Long
     lpstrFile As String
     nMaxFile As Long
     lpstrFileTitle As String
     nMaxFileTitle As Long
     lpstrInitialDir As String
     lpstrTitle As String
     flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     lpstrDefExt As String
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type

' Private To This Module.
Private OFN As OPENFILENAME

' API Functions Used To Invoke The File Open And File SaveAs Dialogs.
Declare Function GetOpenFileName& Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME)
Declare Function GetSaveFileName& Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME)

' Maximum Path And File Name String On Windows.
Public Const MAX_PATH = 260

' Constants (Flag's) Used With Open And SaveAs Dialogs.
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000

Public Function OpenFile() As Boolean
  Dim Rv&

  ' Prepare The OFN Structure For An Open Dialog.
  PrepStruct "Open"
  ' Get A File Name.
  Rv = GetOpenFileName(OFN)
  ' Allow The Dialog To Disappear.
  DoEvents

  If Rv = 1 Then
     GetFileNameAndTitle
     ' Return True.
     OpenFile = True
  End If
End Function

Public Function SaveFile() As Boolean
  Dim Rv&

  ' Prepare The OFN Structure For A Save As Dialog.
  PrepStruct "Save As"
  ' Get A File Name.
  Rv = GetSaveFileName(OFN)
  ' Allow The Dialog To Disappear.
  DoEvents

  If Rv = 1 Then
     GetFileNameAndTitle
     ' Return True.
     SaveFile = True
  End If
End Function

Private Sub PrepStruct(StructType$)
  ' Purpose: Prepares The OFN Structure Ready
  '          For Displaying An Open Or SaveAs Dialog.

  With OFN
     If StructType = "Open" Then
       .flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
       .lpstrTitle = "MicroKit - Open"
     Else
       .flags = OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
       .lpstrTitle = "MicroKit - Save As"
     End If
    .hInstance = App.hInstance
    .hwndOwner = FrmMicroKit.hWnd
    .lpstrDefExt = ".mkf"             ' (Micro Kit File).
    .lpstrFile = String(MAX_PATH, 0)
    .lpstrFileTitle = String(MAX_PATH, 0)
    .lpstrFilter = "MicroKit Files (*.mkf)" & vbNullChar & "*.mkf" & vbNullChar & vbNullChar
    .nFilterIndex = 1
    .lStructSize = Len(OFN)
    .nMaxFile = MAX_PATH
    .nMaxFileTitle = MAX_PATH
  End With
End Sub

Private Sub GetFileNameAndTitle()
  Dim Pos%

  ' Find The Position Of The Null Character In The File Name.
  Pos = InStr(OFN.lpstrFile, vbNullChar)
  ' Tidy Up The File Name.
  NewFileName = Left(OFN.lpstrFile, Pos - 1)

  ' Find The Position Of The Null Character In The File Title.
  Pos = InStr(OFN.lpstrFileTitle, vbNullChar)
  ' Tidy Up The File Title.
  NewFileTitle = Left(OFN.lpstrFileTitle, Pos - 1)
  ' Pretty It Up In Case Of Presentation.
  NewFileTitle = StrConv(NewFileTitle, vbProperCase)
End Sub

Public Function WriteDataToFile(PaFn$, Tempo%) As Boolean
  Dim R%, C%                ' Row, Column Counters.
  Dim SP%                   ' Start Position Of Pattern Being Written.
  Dim EP%                   ' End Position Of Pattern Being Written
  Dim PatternNumber%        ' The Pattern We Are Currently Writing.
  Dim FileNo%               ' Used To Hold The Next Available File Number.
  Dim Pad$                  ' Used For The Creation Of Binary Strings.
  Dim bFileOpen As Boolean  ' Indicates That File Is Open (See Error Trap).

  ' Prep Error Trap.
  On Error GoTo WriteDataError

  FileNo = FreeFile
  Open PaFn For Output As #FileNo
    bFileOpen = True
    ' Save Tempo In File.
    Print #FileNo, Tempo
    For PatternNumber = 0 To 99  ' Write All 100 Patern's.
      ' Add The Note On/Off Data For Each Track Of Each Pattern.
      SP = PatternNumber * 16
      EP = SP + 15
      For R = 0 To 15
        Pad = ""
        For C = SP To EP
            ' Note On/Off (1 Or 0) Info Can Be Looked At As A 16 bit Binary
            ' Which We Then Convert To A Decimal.
            ' i.e. A Form Of Data Compression.
            If Ptrns(R, C).bNoteOn Then Pad = Pad & "1" Else Pad = Pad & "0"
        Next
        If Pad = "0000000000000000" Then ' Track Is Unused In This Pattern.
           ' No Need To Convert (It's Zero).
           Print #FileNo, 0
        Else
           ' Get The Decimal Value Of The Binary String.
           Print #FileNo, GetVal(Pad)
        End If
        ' Save Percusion Name.
        Print #FileNo, Ptrns(R, SP).PercName
        ' Save Percussion Volume.
        Print #FileNo, Ptrns(R, SP).PercVol
      Next
    Next
    ' Save Whether The Song Array Is In Loop Mode.
    Print #FileNo, Abs(CInt(bLoopSong))
    ' Save The Song Patterns.
    For C = 0 To 99
        ' Print The Sequence Of The Patterns.
        Print #FileNo, Song(C)
    Next
  Close #FileNo

  ' Turn Off Error Trapping.
  On Error GoTo 0

  ' Return True.
  WriteDataToFile = True
  ' Indicate That The File Has Been Saved (For If We're Saving Data On A Form Unload).
  bFileSaved = True
  ' Avoid The Error Trap.
  Exit Function

WriteDataError:

    ' Clear The Error (Stop's Error Propagation).
    Err.Clear
    ' If The File Is Still Open, Close It.
    If bFileOpen Then Close #FileNo
    ' Indicate That The File Was Not Saved.
    ' (Used If A Save Fails When The App Is Terminating).
    bFileSaved = False
    ' Return False.
    WriteDataToFile = False
    ' Get Outa Here.
    Exit Function
End Function

Public Function ReadDataFromFile(PaFn$) As Boolean
  Dim R%, C%                        ' Row, Column Counters.
  Dim FileNo%                       ' Used To Hold The Next Available File Number.
  Dim V%, N%, P%                    ' Volume, Instrument And Song Pattern Number Variables.
  Dim SP%                           ' Start Position Of Pattern Being Read.
  Dim PatternNumber%                ' The Pattern We Are Currently Reading.
  Dim DecVal&                       ' Used To Input Track Data In Decimal Format.
  Dim BS$                           ' Used To Turn Decimal Format Into Note On/Off Info.
  Dim LoopStatus%                   ' Holds The Song Loop Status From The File.
  Dim Tmpo%                         ' Holds The Tempo From The File.
  Dim DataTest(15, 1599) As Kit     ' Temp Pattern Array.
  Dim TmpSong(99)                   ' Temp Song Array.
  Dim bFileOpen As Boolean          ' Indicates That File Is Open.

  ' Prep Error Trap.
  On Error GoTo ReadDataError:

  FileNo = FreeFile
  Open PaFn For Input As #FileNo
    bFileOpen = True
    ' Read The Tempo From The File.
    Input #FileNo, Tmpo
    ' Is It A Valid Tempo Setting?
    If Tmpo < 40 Or Tmpo > 255 Then
      Close #FileNo
      ReadDataFromFile = False
      Exit Function
    End If
    ' Input And Build The Pattern Array.
    For PatternNumber = 0 To 99  ' One Hundred Patterns, 16 Beats Wide (1,600 Steps).
      ' Get The Start Pos For Each Pattern SECTION.
      SP = PatternNumber * 16
      For R = 0 To 15
        ' Get The Beat For The Track As A Decimal Value (That's So Wierd).
        Input #FileNo, DecVal
        If DecVal < 0 Or DecVal > 65535 Then
          ' It's Not Track Data.
          Close #FileNo
          ReadDataFromFile = False
          Exit Function
        End If
        ' Convert The Decimal Value To A Binary String (Note On/Off Data).
        If DecVal = 0 Then BS = "0000000000000000" Else BS = GetString(DecVal)
        ' Get The Instrument Name.
        Input #FileNo, N
        If N < 35 Or N > 81 Then
          ' It's Not A Value That One Of The Instruments Can Have.
          Close #FileNo
          ReadDataFromFile = False
          Exit Function
        End If
        ' Get The Track Volume.
        Input #FileNo, V
        If V < 0 Or V > 127 Then
          ' It's A Value That The Volume Can Have.
          Close #FileNo
          ReadDataFromFile = False
          Exit Function
        End If
        For C = 1 To 16
          ' Add The Data To A Temporary Pattern Array.
          ' Must Use A Temp Because If The Data Prove's
          ' To Be Flawed, We Don't Want To Mess Up Any Data We Currently Have.
          ' I.E. Our Last Creation.
          If Mid(BS, C, 1) = "1" Then DataTest(R, C + SP - 1).bNoteOn = True
          DataTest(R, C + SP - 1).PercName = N
          DataTest(R, C + SP - 1).PercVol = V
          DataTest(R, C + SP - 1).MidiMsg = V * &H10000 + N * &H100 + &H99&
        Next
      Next
    Next
    ' Input The "Loop Song" Indicator.
    Input #FileNo, LoopStatus
    If LoopStatus < 0 Or LoopStatus > 1 Then
       ' Can Only Be Zero Or One.
       Close #FileNo
       ReadDataFromFile = False
       Exit Function
    End If
    ' Input The Song Data.
    For C = 0 To 99 ' 100 Patterns.
      Input #FileNo, P
      If P < 0 Or P > 100 Then
        ' It's Not A Valid Pattern Number.
        Close #FileNo
        ReadDataFromFile = False
        Exit Function
      End If
      ' Add The Pattern Number To The Temporary Song Sequencing Array.
      TmpSong(C) = P
    Next
  Close #FileNo

  ' If We Got Here Then All The Data In The File Was Valid.
  ' Let's Now Set The Application Up With The File Data.

  ' Set Up The Tempo.
  FrmMicroKit.VsbTempo.Value = Tmpo
  ' Set The LoopSong Variable.
  bLoopSong = CBool(LoopStatus)

  ' Build The Array That Contains All One Hundred Patterns.
  For R = 0 To 15
    For C = 0 To 1599
      Ptrns(R, C).bNoteOn = DataTest(R, C).bNoteOn
      Ptrns(R, C).PercName = DataTest(R, C).PercName
      Ptrns(R, C).PercVol = DataTest(R, C).PercVol
      Ptrns(R, C).MidiMsg = DataTest(R, C).MidiMsg
    Next
  Next

  ' Build The Song Array.
  For C = 0 To 99
    Song(C) = TmpSong(C)
  Next

  ' Turn Off Error Trapping.
  On Error GoTo 0
  ' Return True.
  ReadDataFromFile = True
  ' Avoid Error Trap.
  Exit Function

ReadDataError:

  ' Clear The Error.
  Err.Clear
  ' If The File Is Still Open, Close It.
  If bFileOpen Then Close #FileNo
  ' Return False.
  ReadDataFromFile = False
  Exit Function
End Function

Private Function GetString$(InVal&)
  ' Purpose: Accepts A Decimal Value And Returns A 16 Bit Binary String.
  '          The Note On/Off Data.

  Dim K%, Pad$

  Pad = "0000000000000000"

  For K = 16 To 1 Step -1
    If InVal Mod 2 Then Mid(Pad, K, 1) = "1"
    InVal = InVal \ 2
    If InVal = 0 Then Exit For
  Next

  GetString = Pad
End Function

Public Function GetVal&(BS$)
  ' Purpose: Accepts A 16 Bit Binary String And Converts It Into A Decimal.
  '          Data Compression.

  Dim K%, Rv&

  For K = 1 To 16
      If Mid(BS, K, 1) = "1" Then Rv = Rv + 2 ^ (16 - K)
  Next

  GetVal = Rv
End Function
