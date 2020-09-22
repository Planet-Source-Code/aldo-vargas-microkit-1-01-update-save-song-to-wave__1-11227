Attribute VB_Name = "WaveRecording"
Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrrtning As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Type OPENFILENAME
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
    
Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000 ' new look commdlg
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String, DefaultFilename As String) As String
  Dim OFN As OPENFILENAME
  Dim A As Long
  On Local Error Resume Next
  OFN.lStructSize = Len(OFN)
  OFN.hwndOwner = Form1.hWnd
  OFN.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

  For A = 1 To Len(Filter)
      If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
  Next
  OFN.lpstrFilter = Filter
  OFN.lpstrFile = Space$(254)
  Mid(OFN.lpstrFile, 1, 254) = DefaultFilename
  OFN.nMaxFile = 255
  OFN.lpstrFileTitle = Space$(254)
  OFN.nMaxFileTitle = 255
  OFN.lpstrInitialDir = InitDir
  OFN.lpstrTitle = Title
  OFN.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
  A = GetSaveFileName(OFN)


  If (A) Then
      SaveDialog = Trim$(OFN.lpstrFile)
  Else
      SaveDialog = ""
  End If
End Function

Public Sub WaveRecord()
  On Error Resume Next
  If WaveReset Then
    WaveSet Channels:=2, Resolution:=16, Rate:=44100
    WaveCapture
  End If
End Sub

Private Function WaveReset() As Boolean
    Dim rtn As String
    Dim i As Long
    
    rtn = Space$(260)
    'Close any MCI operations from previous VB programs
    i = mciSendString("close all", rtn, Len(rtn), 0)
    If i <> 0 Then
      MsgBox "Closing all MCI operations failed!"
      Exit Function
    End If
    'Open a new WAV with MCI Command...
    i = mciSendString("open new type waveaudio alias capture", rtn, Len(rtn), 0)
    If i <> 0 Then
      MsgBox "Opening new wave failed!"
      Exit Function
    End If
    WaveReset = True
End Function

Private Sub WaveSet(Channels As Integer, Resolution As Integer, Rate As Long)
   Dim rtn As String
   Dim i As Long
   Dim settings As String
   Dim Alignment As Integer
      
   rtn = Space$(260)
 
   Alignment = Channels * Resolution / 8
   
   settings = "set capture alignment " & CStr(Alignment) & " bitspersample " & CStr(Resolution) & " samplespersec " & CStr(Rate) & " channels " & CStr(Channels) & " bytespersec " & CStr(Alignment * Rate)

   'Rate = Samples Per Second that are supported:
   '11025     low quality
   '22050     medium quality
   '44100     high quality (CD music quality)
   'Resolution = Bits per sample is 16 or 8
   'Channels = 1 (mono) or 2 (stereo)

   'i = mciSendString("seek capture to start", rtn, Len(rtn), 0) 'Always start at the beginning
   'If i <> 0 Then MsgBox ("Starting recording failed!")
   'You can use at least the following combinations
    
   ' i = mciSendString("set capture alignment 4 bitspersample 16 samplespersec 44100 channels 2 bytespersec 176400", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 2 bitspersample 16 samplespersec 44100 channels 1 bytespersec 88200", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 4 bitspersample 16 samplespersec 22050 channels 2 bytespersec 88200", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 2 bitspersample 16 samplespersec 22050 channels 1 bytespersec 44100", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 4 bitspersample 16 samplespersec 11025 channels 2 bytespersec 44100", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 2 bitspersample 16 samplespersec 11025 channels 1 bytespersec 22050", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 2 bitspersample 8 samplespersec 11025 channels 2 bytespersec 22050", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 1 bitspersample 8 samplespersec 11025 channels 1 bytespersec 11025", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 2 bitspersample 8 samplespersec 8000 channels 2 bytespersec 16000", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 1 bitspersample 8 samplespersec 8000 channels 1 bytespersec 8000", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 2 bitspersample 8 samplespersec 6000 channels 2 bytespersec 12000", rtn, Len(rtn), 0)
   ' i = mciSendString("set capture alignment 1 bitspersample 8 samplespersec 6000 channels 1 bytespersec 6000", rtn, Len(rtn), 0)
   
   i = mciSendString(settings, rtn, Len(rtn), 0)

   If i <> 0 Then MsgBox ("Settings for recording not consistent")
   ' If the combination is not supported you get an error!
End Sub

Private Sub WaveCapture()
    Dim rtn As String
    Dim i As Long
    
    rtn = Space$(260)
   
    i = mciSendString("record capture", rtn, Len(rtn), 0)  'Start the recording
    If i <> 0 Then MsgBox "Recording not possible."
 End Sub

Public Sub WaveStop()
    Dim rtn As String
    Dim i As Long
    i = mciSendString("stop capture", rtn, Len(rtn), 0)
    If i <> 0 Then MsgBox ("Stopping recording failed!")
End Sub

Public Sub WavePlay()
    Dim rtn As String
    Dim i As Long
    i = mciSendString("play capture from 0", rtn, Len(rtn), 0)
    If i <> 0 Then MsgBox ("Start playing failed!")
End Sub

Public Function WaveStatus() As String
    Dim i As Long
    WaveStatus = Space(255)
    i = mciSendString("status capture mode", WaveStatus, 255, 0)
    If i <> 0 Then
      WaveStatus = "N/A"
    Else
      i = InStr(WaveStatus, vbNullChar)
      If i Then
        WaveStatus = Left$(WaveStatus, i - 1)
      Else
        WaveStatus = WaveStatus
      End If
    End If
End Function

Public Function WaveStatistics() As String
    Dim mssg As String * 255
    Dim i As Long
    On Error Resume Next
    i = mciSendString("set capture time format ms", 0&, 0, 0)
    'If i <> 0 Then MsgBox ("Setting time format in milliseconds failed!")
    i = mciSendString("status capture length", mssg, 255, 0)
    mssg = CStr(CLng(mssg) / 1000)
    'If i <> 0 Then MsgBox ("Finding length recording in milliseconds failed!")
    WaveStatistics = "Length recording " & Str(mssg) & " s"

    i = mciSendString("set capture time format bytes", 0&, 0, 0)
    'If i <> 0 Then MsgBox ("Setting time format in bytes failed!")
    i = mciSendString("status capture length", mssg, 255, 0)
    'If i <> 0 Then MsgBox ("Finding length recording in bytes failed!")
    WaveStatistics = WaveStatistics & " (" & Str(mssg) & " bytes)" & vbCrLf

    i = mciSendString("status capture channels", mssg, 255, 0)
    'If i <> 0 Then MsgBox ("Finding number of channels failed!")
    If Str(mssg) = 1 Then
        WaveStatistics = WaveStatistics & "Mono - "
        ElseIf Str(mssg) = 2 Then
            WaveStatistics = WaveStatistics & "Stereo - "
    End If

    i = mciSendString("status capture bitspersample", mssg, 255, 0)
    'If i <> 0 Then MsgBox ("Finding resolution failed!")
    WaveStatistics = WaveStatistics & Str(mssg) & " bits - "

    i = mciSendString("status capture samplespersec", mssg, 255, 0)
    'If i <> 0 Then MsgBox ("Finding sample rate failed!")
    WaveStatistics = WaveStatistics & Str(mssg) & " samples per second " & vbCrLf & vbCrLf
End Function

Public Sub WaveClose()
    Dim rtn As String
    Dim i As Long
    i = mciSendString("close capture", rtn, Len(rtn), 0)
    'If i <> 0 Then MsgBox ("Closing MCI failed!")
End Sub

Public Sub WaveSaveAs(sName As String)
   Dim rtn As String
   Dim i As Long
   Dim WaveShortFileName As String
   
   'If file already exists then remove it
   
    If FileExists(sName) Then
        Kill (sName)
    End If
 
    'The mciSendString API call doesn't seem to like'
    'long filenames that have spaces in them, so we
    'will make another API call to get the short
    'filename version.
    'This is accomplished by the function GetShortName
            
    'MCI command to save the WAV file
     If InStr(sName, " ") Then
        WaveShortFileName = GetShortName(sName)
        ' These are necessary in order to be able to rename file
        i = mciSendString("save capture " & WaveShortFileName, rtn, Len(rtn), 0)
     Else
        i = mciSendString("save capture " & sName, rtn, Len(rtn), 0)
     End If
     If i <> 0 Then MsgBox ("Saving file failed, file name was: " & sName)
End Sub

Private Function FileExists(strFileName As String) As Boolean
  On Error Resume Next
  FileExists = FileLen(strFileName) > 0
End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    If lRetVal = 0 Then 'The file does not exist, first create it!
        Open sLongFileName For Random As #1
        Close #1
        'Now another try!
        lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
        'Delete file now!
        Kill (sLongFileName)
    End If
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
End Function


