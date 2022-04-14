Attribute VB_Name = "modData"
Option Explicit

Public vEPS_DATA As Variant

' CONSTANTS
' Return codes from Registration functions.
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&
Public Const HKEY_CLASSES_ROOT = &H80000000
Const MAX_PATH = 256&
Const REG_SZ = 1
Const KEY_READ = &H20019 'To allow us to READ the registry keys

Public Const sINI = "SSTARTER.INI"
'VARIABLES
Public sIniFile As String
Public FS As Object, FO As Object, FI As Object

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
    (ByVal hwnd As Long) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias _
                "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
                phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
                "RegQueryValueExA" (ByVal hKey As Long, _
                ByVal lpValueName As String, ByVal lpReserved As Long, _
                lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
                 "RegCreateKeyA" (ByVal hKey As Long, _
                 ByVal lpszSubKey As String, _
                 lphKey As Long) As Long

Private Declare Function RegSetValue Lib "advapi32.dll" Alias _
                 "RegSetValueA" (ByVal hKey As Long, _
                 ByVal lpszSubKey As String, _
                 ByVal fdwType As Long, _
                 ByVal lpszValue As String, _
                 ByVal dwLength As Long) As Long
                 
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'*----------------------------------------------------------*
'* Name       : CreateAssociation                           *
'*----------------------------------------------------------*
'* Purpose    : Associate a file type with a program in     *
'*            : Win95 and WinNT                             *
'*----------------------------------------------------------*
'* Parameters : strAppKey   Required. File type alias.      *
'*            : strAppName  Required. File type name.       *
'*            : strExt      Required. File type extension.  *
'*            : strCommand  Required. Command associated    *
'*            :                       with file type.       *
'*----------------------------------------------------------*
Private Sub CreateAssociation(strAppKey As String, _
                              strAppName As String, _
                              strExt As String, _
                              strCommand As String, _
                              Optional strIcon As String = vbNullString)
  Dim sKeyName As String  ' Holds Key Name in registry.
  Dim sKeyValue As String ' Holds Key Value in registry.
  Dim ret As Long         ' Holds error status if any from
  ' API calls.
  Dim lphKey As Long      ' Holds created key handle from
  ' RegCreateKey.
  
On Error GoTo CreateErr

  'Creates a Root entry called strKeyName.
  sKeyName = strAppKey
  sKeyValue = strAppName
  ret = RegCreateKey(HKEY_CLASSES_ROOT, sKeyName, lphKey)
  ret = RegSetValue(lphKey&, "", REG_SZ, sKeyValue, 0&)

  'Creates a Root entry called strExt associated with strKeyName.
  sKeyName = strExt
  sKeyValue = strAppKey
  ret = RegCreateKey(HKEY_CLASSES_ROOT, sKeyName, lphKey)
  ret = RegSetValue(lphKey&, "", REG_SZ, sKeyValue, 0&)

  'Sets the command line for strKeyName.
  sKeyName = strAppKey
  sKeyValue = strCommand
  ret = RegCreateKey(HKEY_CLASSES_ROOT, sKeyName, lphKey)
  ret = RegSetValue(lphKey&, "shell\open\command", REG_SZ, _
                    sKeyValue, MAX_PATH)
  If Len(strIcon) > 0 Then
    sKeyValue = strIcon
    ret = RegSetValue(lphKey&, "DefaultIcon", REG_SZ, _
                    sKeyValue, MAX_PATH)
  End If
  Exit Sub
CreateErr:
  MsgBox Err.Description
  End
End Sub

Public Function GetCommandLine(ByRef fForm As Form)
   Dim CmdLine As String, i As Long, NumArgs As Long, ArgArray As Variant

   CmdLine = Command()
   If Len(CmdLine) < 2 Then
'      Call ChDir(Left$(CurDir, 3))
'      Do
'          UC.Sleep 10000
'          DoEvents
'          UC.Sleep 1000
'          DoEvents
'          UC.Sleep 1000
'          DoEvents
'          UC.Sleep 1000
'          DoEvents
'          UC.Sleep 1000
'          DoEvents
'          UC.Sleep 1000
'          DoEvents
'          UC.Sleep 1000
'          DoEvents
'          UC.Sleep 1000
'          DoEvents
'          UC.Sleep 1000
'          DoEvents
'          UC.Sleep 1000
'          DoEvents
          
'          Call RegisterExtensions
          
'      Loop Until App.PrevInstance
'      End
      fForm.Timer1.Enabled = True
      Exit Function
   End If
   
   If Len(CmdLine) < 8 Then
      MsgBox "Command line argument length is incorrect! May be it's not a file!", vbCritical, "ERROR"
      fForm.Timer1.Enabled = True
      Exit Function
   End If
   ArgArray = Split(CmdLine, Chr$(34) & " " & Chr$(34))
   For i = LBound(ArgArray) To UBound(ArgArray)
      ArgArray(i) = Replace(ArgArray(i), Chr$(34), "")
   Next i
   'Return Array in Function name.
   GetCommandLine = ArgArray
End Function

Public Function FindHWND(ByVal hwnd As Long, ByVal SS$) As Long
Dim CurrWnd As Long, Len1 As Long, ListItem As String
  'ѕолучаем hWnd, который будет первым в списке
  'через него, мы сможем отыскать другие задачи
  CurrWnd = GetWindow(hwnd, 0) 'first window, GW_HWNDFIRST = 0
  'ѕока возвращаемый hWnd имеет смысл, выполн€ем цикл
  Do While CurrWnd <> 0
    'ѕолучаем длину имени задани€ по CurrWnd
    Len1 = GetWindowTextLength(CurrWnd)
    'ѕолучить им€ задачи из списка
    ListItem = Space(Len1 + 1)
    Len1 = GetWindowText(CurrWnd, ListItem, Len1 + 1)
    '≈сли получили им€ задачи, провер€ем на SS
    If Len1 > 0 And InStr(UCase$(ListItem), UCase$(SS$)) > 0 Then
      FindHWND = CurrWnd
      Exit Function
    End If
    'ѕереходим к следующей задаче из списка
    CurrWnd = GetWindow(CurrWnd, 2) ' 2 - next window, GW_HWNDNEXT = 2
    DoEvents
  Loop
  FindHWND = 0
End Function

Public Sub RegisterExtensions_OLD()
   CreateAssociation "Encapsulated PostScript graphics", "Encapsulated PostScript graphics", ".eps", _
                        Chr$(34) & App.Path & "\" & App.EXEName & _
                        ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), _
                        Chr$(34) & App.Path & "\_eps_file.ico" & Chr$(34)
   DoEvents
   CreateAssociation "Adobe Illustrator document", "Adobe Illustrator compatible document", ".ai", _
                        Chr$(34) & App.Path & "\" & App.EXEName & _
                        ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), _
                        Chr$(34) & App.Path & "\_ai_file.ico" & Chr$(34)
   DoEvents
   CreateAssociation "QuarkXPress document", "QuarkXpress (versions 5 and below) compatible document", ".qxd", _
                        Chr$(34) & App.Path & "\" & App.EXEName & _
                        ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), _
                        Chr$(34) & App.Path & "\_qxd_file.ico" & Chr$(34)
   DoEvents
   CreateAssociation "QuarkXPress project", "QuarkXpress (versions 6 and above) compatible document", ".qxp", _
                        Chr$(34) & App.Path & "\" & App.EXEName & _
                        ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), _
                        Chr$(34) & App.Path & "\_qxd_file.ico" & Chr$(34)
   DoEvents
   CreateAssociation "TIFF raster graphics", "Tagged Image File Format", ".tif", _
                        Chr$(34) & App.Path & "\" & App.EXEName & _
                        ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), _
                        Chr$(34) & App.Path & "\_tif_file.ico" & Chr$(34)
   DoEvents
   CreateAssociation "Adobe Indesign document", "Adobe Indesign document", ".indd", _
                        Chr$(34) & App.Path & "\" & App.EXEName & _
                        ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), _
                        Chr$(34) & App.Path & "\_indd_file.ico" & Chr$(34)
   DoEvents
   CreateAssociation "Adobe Poratble Document Format (PDF)", "Adobe Poratble Document Format (PDF)", ".pdf", _
                        Chr$(34) & App.Path & "\" & App.EXEName & _
                        ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), _
                        Chr$(34) & App.Path & "\_pdf_file.ico" & Chr$(34)
   DoEvents
   CreateAssociation "Adobe Indesign interchange document", "Adobe Indesign exchange document", ".inx", _
                        Chr$(34) & App.Path & "\" & App.EXEName & _
                        ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), _
                        Chr$(34) & App.Path & "\_indd_file.ico" & Chr$(34)
   DoEvents
End Sub

Public Sub RegisterExtensions()
    Dim WSO As Object, sKey As String
    Dim tliSSiconGUID As String
    
    On Error Resume Next
    Set WSO = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Error accessing Windows Script Shell Object! Looks like Windows installation is broken :(", vbCritical, "SuperStarter"
        End
    End If
    
' Here we check for appropriated registry key for approved icon handler

    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SSIcon.SuperStarterIconHandler\Clsid\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        frmMain.Timer1.Enabled = False
        Unload frmMain
        Load frmMain
        Exit Sub
    End If
    tliSSiconGUID = sKey

'[HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved]
'"{5D841B9C-1DA2-4ECF-8718-8EAC03988978}" = "SuperStarter icon handler"
    sKey = vbNullString
    sKey = WSO.RegRead("HKLM\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" & tliSSiconGUID)
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" & tliSSiconGUID, _
                "SuperStarter icon handler", "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
        Err.Clear
    Else
        If sKey <> "SuperStarter icon handler" Then
            WSO.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" & tliSSiconGUID, _
                "SuperStarter icon handler", "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        End If
    End If
    Err.Clear
    
    
'GoTo ERERER
    
'---------------------------------------------------------------
' Next we check for appropriated types with correct icon handler
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SuperStarterCDR\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\SuperStarterCDR\", "SuperStarter CorelDRAW illustration", "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterCDR\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterCDR\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
    Else
        If sKey <> "SuperStarter CorelDRAW illustration" Then
            WSO.RegWrite "HKCR\SuperStarterCDR\", "SuperStarter CorelDRAW illustration", "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterCDR\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterCDR\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        Else
            If WSO.RegRead("HKCR\SuperStarterCDR\shellex\IconHandler\") <> tliSSiconGUID Then
                WSO.RegWrite "HKCR\SuperStarterCDR\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            End If
            If WSO.RegRead("HKCR\SuperStarterCDR\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\SuperStarterCDR\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SuperStarterEPS\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\SuperStarterEPS\", "SuperStarter Encapsulated PostScript illustration", "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterEPS\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterEPS\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
    Else
        If sKey <> "SuperStarter Encapsulated PostScript illustration" Then
            WSO.RegWrite "HKCR\SuperStarterEPS\", "SuperStarter Encapsulated PostScript illustration", "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterEPS\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterEPS\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        Else
            If WSO.RegRead("HKCR\SuperStarterEPS\shellex\IconHandler\") <> tliSSiconGUID Then
                WSO.RegWrite "HKCR\SuperStarterEPS\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            End If
            If WSO.RegRead("HKCR\SuperStarterEPS\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\SuperStarterEPS\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SuperStarterAI\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\SuperStarterAI\", "SuperStarter Adobe Illustrator document", "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterAI\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterAI\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
    Else
        If sKey <> "SuperStarter Adobe Illustrator document" Then
            WSO.RegWrite "HKCR\SuperStarterAI\", "SuperStarter Adobe Illustrator document", "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterAI\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterAI\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        Else
            If WSO.RegRead("HKCR\SuperStarterAI\shellex\IconHandler\") <> tliSSiconGUID Then
                WSO.RegWrite "HKCR\SuperStarterAI\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            End If
            If WSO.RegRead("HKCR\SuperStarterAI\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\SuperStarterAI\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SuperStarterTIF\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\SuperStarterTIF\", "SuperStarter Tagged Image File Format", "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterTIF\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterTIF\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
    Else
        If sKey <> "SuperStarter Tagged Image File Format" Then
            WSO.RegWrite "HKCR\SuperStarterTIF\", "SuperStarter Tagged Image File Format", "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterTIF\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterTIF\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        Else
            If WSO.RegRead("HKCR\SuperStarterTIF\shellex\IconHandler\") <> tliSSiconGUID Then
                WSO.RegWrite "HKCR\SuperStarterTIF\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            End If
            If WSO.RegRead("HKCR\SuperStarterTIF\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\SuperStarterTIF\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SuperStarterQXD\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\SuperStarterQXD\", "SuperStarter QuarkXPress document", "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterQXD\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterQXD\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
    Else
        If sKey <> "SuperStarter QuarkXPress document" Then
            WSO.RegWrite "HKCR\SuperStarterQXD\", "SuperStarter QuarkXPress document", "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterQXD\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterQXD\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        Else
            If WSO.RegRead("HKCR\SuperStarterQXD\shellex\IconHandler\") <> tliSSiconGUID Then
                WSO.RegWrite "HKCR\SuperStarterQXD\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            End If
            If WSO.RegRead("HKCR\SuperStarterQXD\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\SuperStarterQXD\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SuperStarterQXP\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\SuperStarterQXP\", "SuperStarter QuarkXPress project", "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterQXP\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterQXP\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
    Else
        If sKey <> "SuperStarter QuarkXPress project" Then
            WSO.RegWrite "HKCR\SuperStarterQXP\", "SuperStarter QuarkXPress project", "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterQXP\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterQXP\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        Else
            If WSO.RegRead("HKCR\SuperStarterQXP\shellex\IconHandler\") <> tliSSiconGUID Then
                WSO.RegWrite "HKCR\SuperStarterQXP\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            End If
            If WSO.RegRead("HKCR\SuperStarterQXP\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\SuperStarterQXP\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SuperStarterPDF\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\SuperStarterPDF\", "SuperStarter Adobe Portable Document Format", "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterPDF\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterCDR\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
    Else
        If sKey <> "SuperStarter Adobe Portable Document Format" Then
            WSO.RegWrite "HKCR\SuperStarterPDF\", "SuperStarter Adobe Portable Document Format", "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterPDF\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterPDF\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        Else
            If WSO.RegRead("HKCR\SuperStarterPDF\shellex\IconHandler\") <> tliSSiconGUID Then
                WSO.RegWrite "HKCR\SuperStarterPDF\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            End If
            If WSO.RegRead("HKCR\SuperStarterPDF\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\SuperStarterPDF\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SuperStarterINDD\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\SuperStarterINDD\", "SuperStarter Adobe Indesign document", "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterINDD\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterINDD\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
    Else
        If sKey <> "SuperStarter Adobe Indesign document" Then
            WSO.RegWrite "HKCR\SuperStarterINDD\", "SuperStarter Adobe Indesign document", "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterINDD\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterINDD\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        Else
            If WSO.RegRead("HKCR\SuperStarterINDD\shellex\IconHandler\") <> tliSSiconGUID Then
                WSO.RegWrite "HKCR\SuperStarterINDD\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            End If
            If WSO.RegRead("HKCR\SuperStarterINDD\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\SuperStarterINDD\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SuperStarterINX\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\SuperStarterINX\", "SuperStarter Adobe Indesign interchange format", "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterINX\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterINX\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
    Else
        If sKey <> "SuperStarter Adobe Indesign interchange format" Then
            WSO.RegWrite "HKCR\SuperStarterINX\", "SuperStarter Adobe Indesign interchange format", "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterINX\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterINX\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        Else
            If WSO.RegRead("HKCR\SuperStarterINX\shellex\IconHandler\") <> tliSSiconGUID Then
                WSO.RegWrite "HKCR\SuperStarterINX\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            End If
            If WSO.RegRead("HKCR\SuperStarterINX\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\SuperStarterINX\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\SuperStarterIDML\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\SuperStarterIDML\", "SuperStarter Adobe Indesign MarkUp Language document", "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterIDML\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
        WSO.RegWrite "HKCR\SuperStarterIDML\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        If Err.Number <> 0 Then ' access denied???
            Err.Clear
            MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                    "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
            End
        End If
    Else
        If sKey <> "SuperStarter Adobe Indesign MarkUp Language document" Then
            WSO.RegWrite "HKCR\SuperStarterIDML\", "SuperStarter Adobe Indesign MarkUp Language document", "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterIDML\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            WSO.RegWrite "HKCR\SuperStarterIDML\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            If Err.Number <> 0 Then ' access denied???
                Err.Clear
                MsgBox "Error accessing Windows Registry for writing! Looks like you don't have appropriate permissions :(" & vbCrLf & _
                        "You must log in as Administrator for correct working or enable program in UAC (under VISTA/7).", vbCritical, "SuperStarter"
                End
            End If
        Else
            If WSO.RegRead("HKCR\SuperStarterIDML\shellex\IconHandler\") <> tliSSiconGUID Then
                WSO.RegWrite "HKCR\SuperStarterIDML\shellex\IconHandler\", tliSSiconGUID, "REG_SZ"
            End If
            If WSO.RegRead("HKCR\SuperStarterIDML\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\SuperStarterIDML\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    
'ERERER:
'------------------------------------------------------------------------------
'remove any persistens handlers from extensions
    WSO.RegRead ("HKCR\.cdr\PersistentHandler\")
    If Err.Number = 0 Then WSO.RegDelete "HKCR\.cdr\PersistentHandler\"
    Err.Clear
    
    WSO.RegRead ("HKCR\.ai\PersistentHandler\")
    If Err.Number = 0 Then WSO.RegDelete "HKCR\.ai\PersistentHandler\"
    Err.Clear
    
    WSO.RegRead ("HKCR\.eps\PersistentHandler\")
    If Err.Number = 0 Then WSO.RegDelete "HKCR\.eps\PersistentHandler\"
    Err.Clear
    
    WSO.RegRead ("HKCR\.tif\PersistentHandler\")
    If Err.Number = 0 Then WSO.RegDelete "HKCR\.tif\PersistentHandler\"
    Err.Clear
    
    WSO.RegRead ("HKCR\.pdf\PersistentHandler\")
    If Err.Number = 0 Then WSO.RegDelete "HKCR\.pdf\PersistentHandler\"
    Err.Clear
    
    WSO.RegRead ("HKCR\.qxd\PersistentHandler\")
    If Err.Number = 0 Then WSO.RegDelete "HKCR\.qxd\PersistentHandler\"
    Err.Clear
    
    WSO.RegRead ("HKCR\.qxp\PersistentHandler\")
    If Err.Number = 0 Then WSO.RegDelete "HKCR\.qxp\PersistentHandler\"
    Err.Clear
    
    WSO.RegRead ("HKCR\.indd\PersistentHandler\")
    If Err.Number = 0 Then WSO.RegDelete "HKCR\.indd\PersistentHandler\"
    Err.Clear
    
    WSO.RegRead ("HKCR\.inx\PersistentHandler\")
    If Err.Number = 0 Then WSO.RegDelete "HKCR\.inx\PersistentHandler\"
    Err.Clear
    
    WSO.RegRead ("HKCR\.idml\PersistentHandler\")
    If Err.Number = 0 Then WSO.RegDelete "HKCR\.idml\PersistentHandler\"
    Err.Clear
'------------------------------------------------------------------------------
' and at end we check for relations between files extensions and file types
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\.cdr\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\.cdr\", "SuperStarterCDR", "REG_SZ"
        WSO.RegWrite "HKCR\.cdr\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        WSO.RegWrite "HKCR\.cdr\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",1", "REG_SZ"
        Err.Clear
    Else
        If sKey <> "SuperStarterCDR" Then
            WSO.RegWrite "HKCR\.cdr\", "SuperStarterCDR", "REG_SZ"
            WSO.RegWrite "HKCR\.cdr\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            WSO.RegWrite "HKCR\.cdr\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",1", "REG_SZ"
            Err.Clear
        Else
            If WSO.RegRead("HKCR\.cdr\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\.cdr\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
            If WSO.RegRead("HKCR\.cdr\DefaultIcon\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",1") Then
                WSO.RegWrite "HKCR\.cdr\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",1", "REG_SZ"
            End If
        End If
    End If
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\.eps\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\.eps\", "SuperStarterEPS", "REG_SZ"
        WSO.RegWrite "HKCR\.eps\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        WSO.RegWrite "HKCR\.eps\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",2", "REG_SZ"
        Err.Clear
    Else
        If sKey <> "SuperStarterEPS" Then
            WSO.RegWrite "HKCR\.eps\", "SuperStarterEPS", "REG_SZ"
            WSO.RegWrite "HKCR\.eps\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            WSO.RegWrite "HKCR\.eps\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",2", "REG_SZ"
            Err.Clear
        Else
            If WSO.RegRead("HKCR\.eps\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\.eps\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
            If WSO.RegRead("HKCR\.eps\DefaultIcon\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",2") Then
                WSO.RegWrite "HKCR\.eps\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",2", "REG_SZ"
            End If
        End If

    End If
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\.ai\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\.ai\", "SuperStarterAI", "REG_SZ"
        WSO.RegWrite "HKCR\.ai\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        WSO.RegWrite "HKCR\.ai\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",3", "REG_SZ"
        Err.Clear
    Else
        If sKey <> "SuperStarterAI" Then
            WSO.RegWrite "HKCR\.ai\", "SuperStarterAI", "REG_SZ"
            WSO.RegWrite "HKCR\.ai\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            WSO.RegWrite "HKCR\.ai\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",3", "REG_SZ"
            Err.Clear
        Else
            If WSO.RegRead("HKCR\.ai\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\.ai\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
            If WSO.RegRead("HKCR\.ai\DefaultIcon\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",3") Then
                WSO.RegWrite "HKCR\.ai\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",3", "REG_SZ"
            End If
        End If

    End If
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\.tif\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\.tif\", "SuperStarterTIF", "REG_SZ"
        WSO.RegWrite "HKCR\.tif\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        WSO.RegWrite "HKCR\.tif\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",4", "REG_SZ"
        Err.Clear
    Else
        If sKey <> "SuperStarterTIF" Then
            WSO.RegWrite "HKCR\.tif\", "SuperStarterTIF", "REG_SZ"
            WSO.RegWrite "HKCR\.tif\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            WSO.RegWrite "HKCR\.tif\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",4", "REG_SZ"
            Err.Clear
        Else
            If WSO.RegRead("HKCR\.tif\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\.tif\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
            If WSO.RegRead("HKCR\.tif\DefaultIcon\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",4") Then
                WSO.RegWrite "HKCR\.tif\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",4", "REG_SZ"
            End If
        End If
    End If
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\.qxd\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\.qxd\", "SuperStarterQXD", "REG_SZ"
        WSO.RegWrite "HKCR\.qxd\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        WSO.RegWrite "HKCR\.qxd\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",5", "REG_SZ"
        Err.Clear
    Else
        If sKey <> "SuperStarterQXD" Then
            WSO.RegWrite "HKCR\.qxd\", "SuperStarterQXD", "REG_SZ"
            WSO.RegWrite "HKCR\.qxd\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            WSO.RegWrite "HKCR\.qxd\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",5", "REG_SZ"
            Err.Clear
        Else
            If WSO.RegRead("HKCR\.qxd\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\.qxd\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
            If WSO.RegRead("HKCR\.qxd\DefaultIcon\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",5") Then
                WSO.RegWrite "HKCR\.qxd\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",5", "REG_SZ"
            End If
        End If
    End If
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\.qxp\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\.qxp\", "SuperStarterQXP", "REG_SZ"
        WSO.RegWrite "HKCR\.qxp\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        WSO.RegWrite "HKCR\.qxp\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",5", "REG_SZ"
        Err.Clear
    Else
        If sKey <> "SuperStarterQXP" Then
            WSO.RegWrite "HKCR\.qxp\", "SuperStarterQXP", "REG_SZ"
            WSO.RegWrite "HKCR\.qxp\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            WSO.RegWrite "HKCR\.qxp\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",5", "REG_SZ"
            Err.Clear
        Else
            If WSO.RegRead("HKCR\.qxp\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\.qxp\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
            If WSO.RegRead("HKCR\.qxp\DefaultIcon\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",5") Then
                WSO.RegWrite "HKCR\.qxp\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",5", "REG_SZ"
            End If
        End If
    End If
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\.pdf\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\.pdf\", "SuperStarterPDF", "REG_SZ"
        WSO.RegWrite "HKCR\.pdf\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        WSO.RegWrite "HKCR\.pdf\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",6", "REG_SZ"
        Err.Clear
    Else
        If sKey <> "SuperStarterPDF" Then
            WSO.RegWrite "HKCR\.pdf\", "SuperStarterPDF", "REG_SZ"
            WSO.RegWrite "HKCR\.pdf\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            WSO.RegWrite "HKCR\.pdf\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",6", "REG_SZ"
            Err.Clear
        Else
            If WSO.RegRead("HKCR\.pdf\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\.pdf\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
            If WSO.RegRead("HKCR\.pdf\DefaultIcon\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",6") Then
                WSO.RegWrite "HKCR\.pdf\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",6", "REG_SZ"
            End If
        End If
    End If
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\.indd\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\.indd\", "SuperStarterINDD", "REG_SZ"
        WSO.RegWrite "HKCR\.indd\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        WSO.RegWrite "HKCR\.indd\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",7", "REG_SZ"
        Err.Clear
    Else
        If sKey <> "SuperStarterINDD" Then
            WSO.RegWrite "HKCR\.indd\", "SuperStarterINDD", "REG_SZ"
            WSO.RegWrite "HKCR\.indd\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            WSO.RegWrite "HKCR\.indd\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",7", "REG_SZ"
            Err.Clear
        Else
            If WSO.RegRead("HKCR\.indd\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\.indd\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
            If WSO.RegRead("HKCR\.indd\DefaultIcon\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",7") Then
                WSO.RegWrite "HKCR\.indd\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",7", "REG_SZ"
            End If
        End If
    End If
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\.inx\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\.inx\", "SuperStarterINX", "REG_SZ"
        WSO.RegWrite "HKCR\.inx\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        WSO.RegWrite "HKCR\.inx\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",8", "REG_SZ"
        Err.Clear
    Else
        If sKey <> "SuperStarterINX" Then
            WSO.RegWrite "HKCR\.inx\", "SuperStarterINX", "REG_SZ"
            WSO.RegWrite "HKCR\.inx\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            WSO.RegWrite "HKCR\.inx\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",8", "REG_SZ"
            Err.Clear
        Else
            If WSO.RegRead("HKCR\.inx\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\.inx\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
            If WSO.RegRead("HKCR\.inx\DefaultIcon\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",8") Then
                WSO.RegWrite "HKCR\.inx\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",8", "REG_SZ"
            End If
        End If
    End If
    
    sKey = vbNullString
    sKey = WSO.RegRead("HKCR\.idml\")
    If Err.Number <> 0 Then ' key is missing?
        Err.Clear
        WSO.RegWrite "HKCR\.idml\", "SuperStarterIDML", "REG_SZ"
        WSO.RegWrite "HKCR\.idml\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
        WSO.RegWrite "HKCR\.idml\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",9", "REG_SZ"
        Err.Clear
    Else
        If sKey <> "SuperStarterAI" Then
            WSO.RegWrite "HKCR\.idml\", "SuperStarterIDML", "REG_SZ"
            WSO.RegWrite "HKCR\.idml\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            WSO.RegWrite "HKCR\.idml\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",9", "REG_SZ"
            Err.Clear
        Else
            If WSO.RegRead("HKCR\.idml\shell\open\command\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)) Then
                WSO.RegWrite "HKCR\.idml\shell\open\command\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), "REG_SZ"
            End If
            If WSO.RegRead("HKCR\.idml\DefaultIcon\") <> (Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",9") Then
                WSO.RegWrite "HKCR\.idml\DefaultIcon\", Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & ",9", "REG_SZ"
            End If
        End If
    End If
    Err.Clear
    
    On Error GoTo 0
    
End Sub

