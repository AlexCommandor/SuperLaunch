Attribute VB_Name = "UC"
Option Explicit
Option Compare Text

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const Flags = SWP_NOSIZE Or SWP_NOMOVE


Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2

Const WM_CLOSE = &H10
Const WM_QUIT = &H0
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const WM_NCACTIVATE = &H86

Const SC_CLOSE = &HF060
Const MF_BYCOMMAND = &H0&

Const DRIVE_REMOVABLE = 2
Const DRIVE_FIXED = 3
Const DRIVE_REMOTE = 4
Const DRIVE_CDROM = 5
Const DRIVE_RAMDISK = 6

Const SM_CXSCREEN = 0
Const SM_CYSCREEN = 1

'Public Const DRIVES_INIT = "C"

Public Enum kbLayout
  kbdENG = 0
  kbdRUS = 1
  kbdUKR = 2
End Enum

Public Enum myBoxState
    myBSStop
    myBSGood
End Enum

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

'Private FS As New Scripting.FileSystemObject
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, lParam As Any) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
      
Private Declare Function SetWindowPos _
Lib "user32" ( _
ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal X As Long, _
ByVal Y As Long, _
ByVal cx As Long, _
ByVal cy As Long, _
ByVal wFlags As Long _
) As Long

Private Const LOCALE_SDECIMAL = &HE         '  decimal separator
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" _
  (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
    (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long


Private Const SPI_SCREENSAVERRUNNING = 97&
Private Declare Function SystemParametersInfo Lib "user32" _
          Alias "SystemParametersInfoA" (ByVal uAction As Long, _
          ByVal uParam As Long, lpvParam As Any, _
          ByVal fuWinIni As Long) As Long

    
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
    "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
    "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public bStopExecution As Boolean
Private ResList() As String, FilesCount As Long

Public Function FindHWND(ByVal InitHWND As Long, ByVal PartialName As String) As Long
Dim CurrWnd As Long, Len1 As Long, ListItem As String
  On Error GoTo ErrHand
  'Получаем hWnd, который будет первым в списке
  'через него, мы сможем отыскать другие задачи
  CurrWnd = GetWindow(InitHWND, GW_HWNDFIRST)
  'Пока возвращаемый hWnd имеет смысл, выполняем цикл
  Do While CurrWnd <> 0
    'Получаем длину имени задания по CurrWnd
    Len1 = GetWindowTextLength(CurrWnd)
    'Получить имя задачи из списка
    ListItem = Space$(Len1 + 1)
    Len1 = GetWindowText(CurrWnd, ListItem, Len1 + 1)
    'Если получили имя задачи, проверяем на SS
    If Len1 > 0 And InStr(UCase$(ListItem), UCase$(PartialName)) > 0 Then
      FindHWND = CurrWnd
      Exit Function
    End If
    'Переходим к следующей задаче из списка
    CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
    DoEvents
  Loop
  FindHWND = 0
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Function FindWndName(ByVal InitHWND As Long, ByVal PartialName As String) As String
Dim CurrWnd As Long, Len1 As Long, ListItem As String
  On Error GoTo ErrHand
  'Получаем hWnd, который будет первым в списке
  'через него, мы сможем отыскать другие задачи
  CurrWnd = GetWindow(InitHWND, GW_HWNDFIRST)
  'Пока возвращаемый hWnd имеет смысл, выполняем цикл
  Do While CurrWnd <> 0
    'Получаем длину имени задания по CurrWnd
    Len1 = GetWindowTextLength(CurrWnd)
    'Получить имя задачи из списка
    ListItem = Space(Len1 + 1)
    Len1 = GetWindowText(CurrWnd, ListItem, Len1 + 1)
    'Если получили имя задачи, проверяем на SS
    If Len1 > 0 And InStr(UCase$(ListItem), UCase$(PartialName)) > 0 Then
      FindWndName = ListItem
      Exit Function
    End If
    'Переходим к следующей задаче из списка
    CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
    DoEvents
  Loop
  FindWndName = ""
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Sub KillTaskByEXEName(ByVal EXETaskName As String, Optional ByVal hwndIgnore As Long = 0)
    Dim hSnapShot As Long, nProcess As Long
    Dim uProcess As PROCESSENTRY32
    Dim hProcess As Long
  On Error GoTo ErrHand
    hSnapShot = CreateToolhelpSnapshot(2, 0)
    uProcess.dwSize = LenB(uProcess)
    nProcess = Process32First(hSnapShot, uProcess)
    Do While nProcess
      If InStr(UCase$(uProcess.szExeFile), UCase$(EXETaskName)) Then
        hProcess = OpenProcess(&H1F0FFF, 1, uProcess.th32ProcessID)
        If hwndIgnore <> hProcess Then TerminateProcess hProcess, 0
        Exit Do
      End If
      nProcess = Process32Next(hSnapShot, uProcess)
    Loop
    CloseHandle hSnapShot
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Public Function TaskIsRunning(ByVal EXETaskName As String) As Boolean
    Dim hSnapShot As Long, nProcess As Long
    Dim uProcess As PROCESSENTRY32
    Dim hProcess As Long, IsRun As Boolean
  On Error GoTo ErrHand
    IsRun = False
    hSnapShot = CreateToolhelpSnapshot(2, 0)
    uProcess.dwSize = LenB(uProcess)
    nProcess = Process32First(hSnapShot, uProcess)
    Do While nProcess
      If InStr(UCase$(uProcess.szExeFile), UCase$(EXETaskName)) Then
        IsRun = True
        Exit Do
      End If
      nProcess = Process32Next(hSnapShot, uProcess)
    Loop
    CloseHandle hSnapShot
    TaskIsRunning = IsRun
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Function EnumerateRunningTasks() As Variant
    Dim hSnapShot As Long, nProcess As Long
    Dim uProcess As PROCESSENTRY32
    Dim hProcess As Long, vArr() As String, i As Long
  On Error GoTo ErrHand
    hSnapShot = CreateToolhelpSnapshot(2, 0)
    uProcess.dwSize = LenB(uProcess)
    nProcess = Process32First(hSnapShot, uProcess)
    ReDim vArr(1 To 1)
    i = 1
    Do While nProcess
        ReDim Preserve vArr(1 To i)
        vArr(i) = Trim$(uProcess.szExeFile)
        uProcess.szExeFile = Space$(260)
        nProcess = Process32Next(hSnapShot, uProcess)
        i = i + 1
    Loop
    ReDim Preserve vArr(1 To i)
    CloseHandle hSnapShot
    EnumerateRunningTasks = vArr
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Function UC_NZ(chkString, Optional DefaultValue As String = vbNullString) As String
  Dim SS As String
  On Error GoTo ErrHand
  If Not IsNull(chkString) Then
    SS = CStr(chkString)
  Else
    SS = DefaultValue
  End If
  UC_NZ = SS
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Function EnumerateDrives(ByVal WithDriveType As Boolean)
  Dim DriveCount As Long, strDrive As String, DrType As Long
  Dim Drives() As String, DriveString As String, FoundedCount As Long
  On Error GoTo ErrHand
  DriveString = ""
  FoundedCount = 0
  ReDim Preserve Drives(1 To 2, 1 To 1)
  For DriveCount = Asc("A") To Asc("Z")
    strDrive = Chr(DriveCount) & ":\"
    DrType = GetDriveType(strDrive)
    If DrType > 1 Then
      FoundedCount = FoundedCount + 1
      If WithDriveType Then
        ReDim Preserve Drives(1 To 2, 1 To FoundedCount)
        Drives(1, FoundedCount) = Chr(DriveCount)
        Select Case DrType
          Case DRIVE_REMOVABLE
            Drives(2, FoundedCount) = "REMOVABLE"
          Case DRIVE_FIXED
            Drives(2, FoundedCount) = "FIXED"
          Case DRIVE_REMOTE
            Drives(2, FoundedCount) = "REMOTE"
          Case DRIVE_CDROM
            Drives(2, FoundedCount) = "CDROM"
          Case DRIVE_RAMDISK
            Drives(2, FoundedCount) = "RAMDISK"
        End Select
      Else
        DriveString = DriveString & Chr(DriveCount)
      End If
    End If
  Next DriveCount
  If WithDriveType Then
    EnumerateDrives = Drives
  Else
    EnumerateDrives = DriveString
  End If
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Function GetDriveTypeString(cLetter As String) As String
  Dim cLett As String * 3, DrType As Long, sRes As String
  On Error GoTo ErrHand
  If Trim$(cLetter) = vbNullString Then
    GetDriveTypeString = vbNullString
    Exit Function
  End If
  cLett = UCase$(Left$(cLetter, 1)) & ":\"
  DrType = GetDriveType(cLett)
  If DrType > 1 Then
    Select Case DrType
      Case DRIVE_REMOVABLE
        sRes = "REMOVABLE"
      Case DRIVE_FIXED
        sRes = "FIXED"
      Case DRIVE_REMOTE
        sRes = "REMOTE"
      Case DRIVE_CDROM
        sRes = "CDROM"
      Case DRIVE_RAMDISK
        sRes = "RAMDISK"
    End Select
  Else
    sRes = vbNullString
  End If
  GetDriveTypeString = sRes
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Sub SetOnTopWindow(hwnd As Long, OnTop As Boolean)

  On Error GoTo ErrHand
  If OnTop = True Then 'Make the window topmost
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags
  Else
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags
  End If
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Sub MyWinPlace(InitHWND As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal W As Long, ByVal H As Long, ByVal wndName As String)
    Dim ScreenX&, ScreenY&, NewX&, NewY&, NewW&, NewH&, hwnd As Long
  On Error GoTo ErrHand
    ScreenX& = GetSystemMetrics(SM_CXSCREEN)
    ScreenY& = GetSystemMetrics(SM_CYSCREEN)
    NewX& = X * ScreenX& / 1000
    NewY& = Y * ScreenY& / 1000
    NewW& = W * ScreenX& / 1000
    NewH& = H * ScreenY& / 1000
    hwnd = FindHWND(InitHWND, wndName)
    MoveWindow hwnd, NewX&, NewY&, NewW&, NewH&, 1
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Public Sub ActivateWindow(InitHWND As Long, PartialWindowName As String)
  Dim hW As Long
  On Error GoTo ErrHand
  hW = FindHWND(InitHWND, PartialWindowName)
  If hW <> 0 Then SetForegroundWindow hW
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Public Sub SetKbdrdLayout(KeybLayout As kbLayout)
  Dim i As Long
  On Error GoTo ErrHand
  Select Case KeybLayout
    Case kbdENG
      i = LoadKeyboardLayout("00000409", 1)
    Case kbdRUS
      i = LoadKeyboardLayout("00000419", 1)
    Case kbdUKR
      i = LoadKeyboardLayout("00000422", 1)
  End Select
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Public Function GetFilteredStringArray(ByRef vInputArray As Variant, _
      sLikeFilter As String, Optional bCaseSensitive As Boolean = False) As Variant
  Dim vRes() As String, LB As Long, UB As Long, i As Long
  Dim MaxElem As Long, bMainCondition As Boolean
  On Error Resume Next
  LB = LBound(vInputArray): UB = UBound(vInputArray)
'  If LB = UB Or (UB - LB) < 5 Then
  If LB = UB Then
    GetFilteredStringArray = vbNullString
    Exit Function
  End If
  On Error GoTo ErrHand
  ReDim vRes(1 To 1)
  For i = LB To UB
    bMainCondition = False
    If bCaseSensitive Then
      bMainCondition = vInputArray(i) Like sLikeFilter
    Else
      bMainCondition = UCase$(vInputArray(i)) Like UCase$(sLikeFilter)
    End If
    If bMainCondition Then
      If IsArray(vRes) Then MaxElem = UBound(vRes) + 1 Else MaxElem = 1
      ReDim Preserve vRes(LB To MaxElem)
      vRes(MaxElem) = vInputArray(i)
    End If
  Next i
  GetFilteredStringArray = vRes
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Sub DisableCloseButton(ByVal FormHWND As Long)
  Dim hMenu As Long, Success As Long
  On Error GoTo ErrHand
  hMenu = GetSystemMenu(FormHWND, 0)
  Success = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
  SendMessage FormHWND, WM_NCACTIVATE, 0&, 0&
  SendMessage FormHWND, WM_NCACTIVATE, 1&, 0
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Public Function GetLoggedInUser() As String
  Dim lpN As String, lpU As String, lpL As Long, lRes As Long
  On Error Resume Next
  lpL = 255: lpU = Space$(256)
  lRes = WNetGetUser(lpN, lpU, lpL)
  If lRes = 0 Then 'All OK
    lpU = Left$(lpU, InStr(lpU, Chr$(0)) - 1)
  Else
    lpU = vbNullString
  End If
  GetLoggedInUser = lpU
  Err.Clear
  On Error GoTo 0
End Function

Public Function GetNetBIOSComputerName() As String
  Dim lpU As String, lpL As Long, lRes As Boolean
  On Error GoTo ErrHand
  lpL = 255: lpU = Space$(256)
  lRes = GetComputerName(lpU, lpL)
  If lRes Then 'All OK
    lpU = Left$(lpU, InStr(lpU, Chr$(0)) - 1)
  Else
    lpU = vbNullString
  End If
  GetNetBIOSComputerName = lpU
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Sub SetCurrProcessVisibleInTaskList(Optional ByVal bVisible As Boolean = True)
  On Error GoTo ErrHand
  If bVisible Then
    RegisterServiceProcess GetCurrentProcessId, 0 'Show app
  Else
    RegisterServiceProcess GetCurrentProcessId, 1 'Hide app
  End If
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Public Sub SetDecimalSeparator(ByVal sDecSeparator As String)
  Dim iLocale As Integer, sTmpStr As String, lRes As Long
  On Error Resume Next
  If Len(sDecSeparator) = 0 Then Exit Sub
  sTmpStr = sDecSeparator
  If Len(sTmpStr) > 4 Then sTmpStr = Left$(sTmpStr, 4)
  sTmpStr = sTmpStr & Chr$(0)
  iLocale = GetUserDefaultLCID()
  lRes = SetLocaleInfo(iLocale, LOCALE_SDECIMAL, sTmpStr)
  Err.Clear
  On Error GoTo 0
End Sub

Public Function GetDecimalSeparator() As String
  Dim iLocale As Integer, sTmpStr As String, lRes As Long, aLen As Long
  On Error Resume Next
  sTmpStr = String$(255, " ") & Chr$(0)
  aLen = 1
  iLocale = GetUserDefaultLCID()
  lRes = GetLocaleInfo(iLocale, LOCALE_SDECIMAL, sTmpStr, aLen)
  GetDecimalSeparator = Left$(sTmpStr, aLen)
  Err.Clear
  On Error GoTo 0
End Function

Public Sub SetCtrlAltDelAndAltTab(Optional ByVal bOff As Boolean = True)
  Dim lngRet As Long
  Dim blnOld As Boolean
  On Error GoTo ErrHand
  lngRet = SystemParametersInfo(SPI_SCREENSAVERRUNNING, bOff, blnOld, 0&)
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Public Function GetWinDir(Optional ByVal bWithSlash As Boolean = False) As String
  Dim lpBuffer As String * 255
  Dim lLength As Long
  On Error GoTo ErrHand
  lLength = GetWindowsDirectory(lpBuffer, Len(lpBuffer))
  If bWithSlash Then
    GetWinDir = Left$(lpBuffer, lLength) & "\"
  Else
    GetWinDir = Left$(lpBuffer, lLength)
  End If
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Function GetSysDir(Optional ByVal bWithSlash As Boolean = False) As String
  Dim lpBuffer As String * 255
  Dim lLength As Long
  On Error GoTo ErrHand
  lLength = GetSystemDirectory(lpBuffer, Len(lpBuffer))
  If bWithSlash Then
    GetSysDir = Left$(lpBuffer, lLength) & "\"
  Else
    GetSysDir = Left$(lpBuffer, lLength)
  End If
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Function FileExists(sFileName As String) As Boolean
  Dim iFNum As Integer
  On Error Resume Next
  Err.Clear: iFNum = FreeFile
  Open sFileName For Input As iFNum
  If Err.Number <> 0 Then 'File may be not exists or access denied!
    Err.Clear: FileExists = False
  Else
    Close #iFNum: FileExists = True
  End If
End Function

Public Sub ErrMsgO(ByRef oErr As ErrObject, _
      Optional ByVal sMsgCaption As String = "Ошибка!", _
      Optional bStop As Boolean = False)
  MsgBox "Ошибка выполнения <" & oErr.Description & ">, код " & _
      oErr.Number & "." & vbCrLf & _
      "По старой доброй привычке считаем это недопустимой операцией и закрываемся." & vbCrLf & _
      "Если слабо разобраться в причинах ошибки, обратитесь к разработчику :) .", _
      vbCritical, sMsgCaption
  If bStop Then End
End Sub


Public Sub MySleep(ByVal lSeconds As Long)
   Dim t1 As Date, t2 As Date
   t1 = Now()
   t2 = DateAdd("s", lSeconds, t1)
   Do While t2 >= Now()
      DoEvents
   Loop
End Sub

