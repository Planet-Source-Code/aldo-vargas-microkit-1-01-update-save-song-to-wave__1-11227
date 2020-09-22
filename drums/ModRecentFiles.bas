Attribute VB_Name = "ModRecentFiles"
Option Explicit

' Author:        Stuart Pennington.
' Project:       Midi Percussion Sequencer (Drum Machine).
' Test Platform: Windows 98SE
' Processor:     P2 300MHz.



Type SECURITY_ATTRIBUTES
     nLength As Long
     lpSecurityDescriptor As Long
     bInheritHandle As Long
End Type

Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult&, lpdwDisposition&)
Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey&, ByVal lpValueName$)
Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)
Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, lpType&, ByVal lpData$, lpcbData&)
Declare Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal Reserved&, ByVal dwType&, ByVal lpData$, ByVal cbData&)

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const KEY_ALL_ACCESS = &HF003F

Public Sub GetRecentFileList(MRec As Object)
   ' Purpose: Get's The Four Most Recent File's The User Has Accessed
   '          And Add's Them to The "Recent" Menu.

   Dim K%, Pos%, N%
   Dim Buffer$, hKey&, Rv&
  
   N = 1

   Rv = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\CyberVision\MicroKit\Recent Files", 0, KEY_ALL_ACCESS, hKey)
   If Rv = 0 Then
      For K = 1 To 4
        ' Prep A Buffer To Accept String Data From The Reg.
        Buffer = Space(260)
        ' Query The Value.
        Rv = RegQueryValueEx(hKey, CStr(K) & ".", 0, 0, Buffer, 260)
        ' Find The Null Char In The Returned String.
        Pos = InStr(Buffer, vbNullChar)
        If Pos <> 0 Then
           ' Make The Seperator Visible.
           MRec(0).Visible = 1
           ' Add The Full Path To The Menu Tag.
           MRec(N).Tag = Left(Buffer, Pos - 1)
           ' Add The File Title To The Menu (Neat).
           MRec(N).Caption = "&" & CStr(N) & " " & GetTitle(MRec(N).Tag)
           ' Show The File Title On The Menu.
           MRec(N).Visible = 1
           ' Increase The Count.
           N = N + 1
        End If
        ' Remove The Extracted Value From The Registry.
        RegDeleteValue hKey, CStr(K) & "."
      Next
      ' Don't Forget To Close The Reg Key Now We're Done.
      RegCloseKey hKey
   End If
End Sub

Public Sub SaveRecentFileList(MRec As Object)
  Dim StrVal$, K%, hKey&, Rv&
  Dim SA As SECURITY_ATTRIBUTES  ' Ignored By Windows 95/98 But Not By Win 2000.

  SA.bInheritHandle = True
  SA.lpSecurityDescriptor = 0
  SA.nLength = Len(SA)

  ' Open/Create The Key.
  Rv = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\CyberVision\MicroKit\Recent Files", 0, vbNullString, 0, KEY_ALL_ACCESS, SA, hKey, 0)
  If Rv = 0 Then
    For K = 1 To 4
      ' Get The Full Path From The Recent Menu Item's Tag.
      StrVal = MRec(K).Tag
      If StrVal = "" Then
         ' Nothing Else To Save To Reg.
         Exit For
      Else
         ' Save The Recent File To The Registry.
         RegSetValueEx hKey, CStr(K) & ".", 0, 1, StrVal, Len(StrVal)
      End If
    Next
    ' Don't Forget To Close The Reg Key Now We're Done.
    RegCloseKey hKey
  End If
End Sub

Public Sub GetLastDir()
   ' Purpose: Retreives The Last Directory The User Was Working In.
   '          Handy Because It Let's The Continue Browsing From Where
   '          They Left Off.

   Dim Buffer$, hKey&, Rv&
  
   ChDrive Left(App.Path, 3)
   ChDir App.Path

   ' Open The Registry Key For Reading.
   Rv = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\CyberVision\MicroKit\LastDir", 0, KEY_ALL_ACCESS, hKey)
   If Rv = 0 Then
      ' Prep A Buffer To Accept String Data From The Reg.
      Buffer = Space(260)
      Rv = RegQueryValueEx(hKey, "LWD", 0, 0, Buffer, 260)
      If Rv = 0 Then
        On Error Resume Next
          ' Is The Last Place They Went Still Available?
          Buffer = Left(Buffer, InStr(Buffer, vbNullChar) - 1)
          ChDrive Left(Buffer, 3)
          ChDir Buffer
          ' Ifd It Ain't, Just Clear The Error.
          If Err.Number Then Err.Clear
        On Error GoTo 0
      End If
      ' Delete The Value, We've Got It.
      RegDeleteValue hKey, "LWD"
      ' Don't Forget To Close The Reg Key Now We're Done.
      RegCloseKey hKey
   End If
End Sub

Public Sub SaveLastDir()
  ' Purpose: Save's The Last Folder The User Was In.

  Dim hKey&, Rv&, Ldir$
  Dim SA As SECURITY_ATTRIBUTES  ' Ignored By Windows 95/98 But Not By Win 2000.

  SA.bInheritHandle = True
  SA.lpSecurityDescriptor = 0
  SA.nLength = Len(SA)

  Rv = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\CyberVision\MicroKit\LastDir", 0, vbNullString, 0, KEY_ALL_ACCESS, SA, hKey, 0)
  If Rv = 0 Then
     ' Get The Current Directory.
     Ldir = CurDir
     ' Save It.
     RegSetValueEx hKey, "LWD", 0, 1, Ldir, Len(Ldir)
     ' Don't Forget To Close The Reg Key Now We're Done.
     RegCloseKey hKey
  End If
End Sub

Public Sub UpdateRecent(MRec As Object)
  Dim K%, N
  Dim NewArray$(1 To 4)

  N = 1

  For K = 1 To 4
    If MRec(K).Tag <> "" Then
      NewArray(N) = MRec(K).Tag
      N = N + 1
      MRec(K).Tag = ""
    End If
  Next

  For K = 1 To 4
    If NewArray(K) <> "" Then
       MRec(K).Tag = NewArray(K)
    Else
       MRec(K).Visible = 0
    End If
  Next

  If MRec(1).Visible = 0 Then
     MRec(0).Visible = 0
  Else
     BuildMenu MRec
  End If
End Sub

Public Sub BuildMenu(MRec As Object)
  Dim K%

  For K = 1 To 4
      If MRec(K).Tag <> "" Then
         MRec(K).Caption = "&" & CStr(K) & " " & GetTitle(MRec(K).Tag)
      End If
  Next
End Sub

Public Sub AddToRecent(Fn$, MRec As Object)
  Dim Temp$
  Dim K%, N%
  Dim bFound As Boolean

  MRec(0).Visible = 1

  For K = 1 To 4
    If LCase(MRec(K).Tag) = LCase(Fn) Then
       bFound = True
       Exit For
    End If
  Next

  If bFound Then
    If K > 1 Then
       Temp = MRec(K).Tag
       For N = K To 1 Step -1
           MRec(N).Tag = MRec(N - 1).Tag
       Next
       MRec(1).Tag = Temp
    End If
  Else
    For K = 4 To 1 Step -1
        MRec(K).Tag = MRec(K - 1).Tag
    Next
    MRec(1).Tag = Fn
  End If

  For K = 1 To 4
    If MRec(K).Tag <> "" Then MRec(K).Visible = 1
  Next

  BuildMenu MRec
End Sub

Public Sub SetUpIconDblClick()
  Dim RegData$, hKey&, Rv&
  Dim SA As SECURITY_ATTRIBUTES

  SA.bInheritHandle = True
  SA.lpSecurityDescriptor = 0
  SA.nLength = Len(SA)

  ' Give Our File Extension To Reg.
  Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, ".mkf", 0, vbNullString, 0, KEY_ALL_ACCESS, SA, hKey, 0)
  If Rv = 0 Then
     RegSetValueEx hKey, vbNullString, 0, 1, "mkffile", 7
     RegCloseKey hKey
  End If
  ' Tie It Up.
  Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, "mkffile", 0, vbNullString, 0, KEY_ALL_ACCESS, SA, hKey, 0)
  If Rv = 0 Then
     RegSetValueEx hKey, vbNullString, 0, 1, "MicroKit File", 14
     RegCloseKey hKey
  End If

  ' Prep Default Icon String For Our Files.
  If Right(App.Path, 1) = "\" Then
     RegData = App.Path & "MicroKit.exe,-10"  ' Icon Resource ID = 10, Therefor Specify -10.
  Else
     RegData = App.Path & "\MicroKit.exe,-10"
  End If
  ' Write Default Icon Data.
  Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, "mkffile\DefaultIcon", 0, vbNullString, 0, KEY_ALL_ACCESS, SA, hKey, 0)
  If Rv = 0 Then
     RegSetValueEx hKey, vbNullString, 0, 1, RegData, Len(RegData)
     RegCloseKey hKey
  End If

  ' Set Up Icon Double Click. (Will Work When App Is An Exe).
  Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, "mkffile\Shell\Open\Command", 0, vbNullString, 0, KEY_ALL_ACCESS, SA, hKey, 0)
  If Rv = 0 Then
     RegData = Left(RegData, Len(RegData) - 4)
     RegData = RegData & " /open %1"
     RegSetValueEx hKey, vbNullString, 0, 1, RegData, Len(RegData)
     RegCloseKey hKey
  End If
End Sub
