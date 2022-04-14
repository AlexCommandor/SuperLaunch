VERSION 5.00
Object = "{043CDC0A-B54E-4AD9-A637-BBD69EE04568}#31.0#0"; "INIPRO.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr_TempCleaner 
      Interval        =   60000
      Left            =   3120
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4440
      Top             =   480
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin IniPro.Ini Ini1 
      Left            =   0
      Top             =   0
      _ExtentX        =   3387
      _ExtentY        =   635
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bEPSAnalizing As Boolean

Public WithEvents FormSys As FrmSysTray
Attribute FormSys.VB_VarHelpID = -1

'Private Declare Function CreateMutex Lib "kernel32" _
        Alias "CreateMutexA" _
       (ByVal lpMutexAttributes As Long, _
        ByVal bInitialOwner As Long, _
        ByVal lpName As String) As Long

'Private Type TypeIcon
'    cbSize As Long
'    picType As PictureTypeConstants
'    hIcon As Long
'End Type

'Private Type CLSID
'    id(16) As Byte
'End Type

'Private Const MAX_PATH = 260
'Private Type SHFILEINFO
'    hIcon As Long                      '  out: icon
'    iIcon As Long                      '  out: icon index
'    dwAttributes As Long               '  out: SFGAO_ flags
'    szDisplayName As String * MAX_PATH '  out: display name (or path)
'    szTypeName As String * 80          '  out: type name
'End Type

'Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
'Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'Private Const SHGFI_ICON = &H100
'Private Const SHGFI_LARGEICON = &H0
'Private Const SHGFI_SMALLICON = &H1

'' Convert an icon handle into an IPictureDisp.
'Private Function IconToPicture(hIcon As Long) As IPictureDisp
'Dim cls_id As CLSID
'Dim hRes As Long
'Dim new_icon As TypeIcon
'Dim lpUnk As IUnknown'
'
'    With new_icon
'        .cbSize = Len(new_icon)
'        .picType = vbPicTypeIcon
'        .hIcon = hIcon
'    End With
'    With cls_id
'        .id(8) = &HC0
'        .id(15) = &H46
'    End With
'    hRes = OleCreatePictureIndirect(new_icon, _
'        cls_id, 1, lpUnk)
'    If hRes = 0 Then Set IconToPicture = lpUnk
'End Function

' Return a file's icon.
'Private Function GetIcon(ByVal filename As String, ByVal icon_size As Long) As IPictureDisp
'Dim index As Integer
'Dim hIcon As Long
'Dim item_num As Long
'Dim icon_pic As IPictureDisp
'Dim sh_info As SHFILEINFO'
'
'    SHGetFileInfo filename, 0, sh_info, _
'        Len(sh_info), SHGFI_ICON + icon_size
'    hIcon = sh_info.hIcon
'    Set icon_pic = IconToPicture(hIcon)
'    Set GetIcon = icon_pic
'End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then bEPSAnalizing = False: Me.Hide
End Sub

Private Sub Form_Load()
   Dim vCommands, i As Long, sCreator As String, sProg As String, sTmp As String
   Dim lRes As Long, lHandleI As Long, sExtension As String * 3
   Dim tEPS As ftEPSInfo, tTIF As ftTIFFInfo, sINFO As String, iPDFVersion As Integer, j As Integer
   Dim FLSO As Object, iAcrobatVersion As Integer, sAcrobatVersion As String
   Dim oSSI As Object, arrData() As Byte, WSO As Object, lResTemp As Long
      
   Me.Hide
   
    On Error Resume Next
    Set FLSO = CreateObject("Scripting.FileSystemObject")
    Set oSSI = CreateObject("SSIcon.SuperStarterIconHandler")
    If Err.Number <> 0 Then
        Err.Clear
'        If Not FLSO.FileExists(App.Path & "\SSICON.DLL") Then
'            arrData = LoadResData(101, "SSICON.DLL")
'            Open App.Path & "\SSICON.DLL" For Binary Access Write As #1
'                Put #1, , arrData
'            Close #1
'            Erase arrData
'        End If
'        Err.Clear
'        Call Shell(Environ$("SYSTEMROOT") & "\system32\regsvr32.exe /s " & Chr$(34) & App.Path & "\SSICON.DLL" & Chr$(34), vbHide)
'        If Err.Number <> 0 Then
'            MsgBox "Error accessing windows registry for writing!" & vbCrLf & _
'            "Please ensure you have administrator access!", _
'                vbCritical, "SuperStarter"
'            End
'        End If
'
'        Err.Clear
'        Set oSSI = CreateObject("SSIcon.SuperStarterIconHandler")
'        If Err.Number <> 0 Then
            MsgBox "Error creating and/or accessing SuperStarterIconHandler object!" & _
            "You have to reinstall programm and use it under administrator access!", _
                vbCritical, "SuperStarter"
            End
'        End If
    End If
    
    Set oSSI = Nothing
    
    Set oSSI = CreateObject("GflAx.GflAx")
    If Err.Number <> 0 Then
        Err.Clear
'        If Not FLSO.FileExists(App.Path & "\GFLAX.DLL") Then
'            arrData = LoadResData(101, "GFLAX.DLL")
'            Open App.Path & "\GFLAX.DLL" For Binary Access Write As #1
'                Put #1, , arrData
'            Close #1
'        End If
'        Err.Clear
'        Call Shell(Environ$("SYSTEMROOT") & "\system32\regsvr32.exe /s " & Chr$(34) & App.Path & "\GFLAX.DLL" & Chr$(34), vbHide)
'        If Err.Number <> 0 Then
'            MsgBox "Error accessing windows system folder!" & vbCrLf & _
'            "Please ensure you have administrator access!", _
'                vbCritical, "SuperStarter"
'            End
'        End If
'
'        Err.Clear
'        Set oSSI = CreateObject("GflAx.GflAx")
'        If Err.Number <> 0 Then
            MsgBox "Error creating and/or accessing GflAx object!" & _
            "You have to reinstall programm and use it under administrator access!", _
                vbCritical, "SuperStarter"
            End
'        End If
    End If
    
    Set oSSI = CreateObject("MSXML.DOMDocument")
    If Err.Number <> 0 Then
        MsgBox "Error accessing Microsoft XML runtime! Looks like you have Windows 98 :) , not XP or above." & vbCrLf & _
                "You have to google 'Microsoft XML' and install it." & vbCrLf & _
                "Without MSXML you won't able to see correct versions of Adobe Indesign INX and IDML formats.", _
                vbCritical, "SuperStarter"
        End
    End If
    Set oSSI = Nothing
    
'    If Not FLSO.FileExists(App.Path & "\_CDR_FILE.ICO") Then
'        Erase arrData
'        arrData = LoadResData(101, "ICONS")
'        Open App.Path & "\_CDR_FILE.ICO" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
'    If Not FLSO.FileExists(App.Path & "\_EPS_FILE.ICO") Then
'        Erase arrData
'        arrData = LoadResData(102, "ICONS")
'        Open App.Path & "\_EPS_FILE.ICO" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
'    If Not FLSO.FileExists(App.Path & "\_AI_FILE.ICO") Then
'        Erase arrData
'        arrData = LoadResData(103, "ICONS")
'        Open App.Path & "\_AI_FILE.ICO" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
'    If Not FLSO.FileExists(App.Path & "\_TIF_FILE.ICO") Then
'        Erase arrData
'        arrData = LoadResData(104, "ICONS")
'        Open App.Path & "\_TIF_FILE.ICO" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
'    If Not FLSO.FileExists(App.Path & "\_QXD_FILE.ICO") Then
'        Erase arrData
'        arrData = LoadResData(105, "ICONS")
'        Open App.Path & "\_QXD_FILE.ICO" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
'    If Not FLSO.FileExists(App.Path & "\_PDF_FILE.ICO") Then
'        Erase arrData
'        arrData = LoadResData(106, "ICONS")
'        Open App.Path & "\_PDF_FILE.ICO" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
'    If Not FLSO.FileExists(App.Path & "\_INDD_FILE.ICO") Then
'        Erase arrData
'        arrData = LoadResData(107, "ICONS")
'        Open App.Path & "\_INDD_FILE.ICO" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
'    If Not FLSO.FileExists(App.Path & "\_INX_FILE.ICO") Then
'        Erase arrData
'        arrData = LoadResData(108, "ICONS")
'        Open App.Path & "\_INX_FILE.ICO" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
'    If Not FLSO.FileExists(App.Path & "\_IDML_FILE.ICO") Then
'        Erase arrData
'        arrData = LoadResData(109, "ICONS")
'        Open App.Path & "\_IDML_FILE.ICO" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
    If Not FLSO.FileExists(App.Path & "\NO_PREVIEW.ICO") Then
        Erase arrData
        arrData = LoadResData(101, "ICONS")
        Open App.Path & "\NO_PREVIEW.ICO" For Binary Access Write As #1
            Put #1, , arrData
        Close #1
    End If
    Err.Clear
    
'    If Not FLSO.FileExists(App.Path & "\VBSHELL.TLB") Then
'        Erase arrData
'        arrData = LoadResData(101, "VBSHELL.TLB")
'        Open App.Path & "\VBSHELL.TLB" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
'    If Not FLSO.FileExists(App.Path & "\VBSHELL.ODL") Then
'        Erase arrData
'        arrData = LoadResData(102, "VBSHELL.TLB")
'        Open App.Path & "\VBSHELL.ODL" For Binary Access Write As #1
'            Put #1, , arrData
'        Close #1
'    End If
'    Err.Clear
    
    Set FLSO = Nothing
    
    Err.Clear
   
   sIniFile = App.Path & "\" & sINI
   
   If GetSetting(App.EXEName, "Parameters", "FirstRunHappened") <> "YES" Then
        Call Shell("regedit.exe /s /e " & Chr$(34) & App.Path & "\backup_cdr.reg" & Chr$(34) & " " & _
                Chr$(34) & "HKEY_CLASSES_ROOT\.cdr" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /e " & Chr$(34) & App.Path & "\backup_eps.reg" & Chr$(34) & " " & _
                Chr$(34) & "HKEY_CLASSES_ROOT\.eps" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /e " & Chr$(34) & App.Path & "\backup_ai.reg" & Chr$(34) & " " & _
                Chr$(34) & "HKEY_CLASSES_ROOT\.ai" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /e " & Chr$(34) & App.Path & "\backup_tif.reg" & Chr$(34) & " " & _
                Chr$(34) & "HKEY_CLASSES_ROOT\.tif" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /e " & Chr$(34) & App.Path & "\backup_qxd.reg" & Chr$(34) & " " & _
                Chr$(34) & "HKEY_CLASSES_ROOT\.qxd" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /e " & Chr$(34) & App.Path & "\backup_qxp.reg" & Chr$(34) & " " & _
                Chr$(34) & "HKEY_CLASSES_ROOT\.qxp" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /e " & Chr$(34) & App.Path & "\backup_pdf.reg" & Chr$(34) & " " & _
                Chr$(34) & "HKEY_CLASSES_ROOT\.pdf" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /e " & Chr$(34) & App.Path & "\backup_indd.reg" & Chr$(34) & " " & _
                Chr$(34) & "HKEY_CLASSES_ROOT\.indd" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /e " & Chr$(34) & App.Path & "\backup_inx.reg" & Chr$(34) & " " & _
                Chr$(34) & "HKEY_CLASSES_ROOT\.inx" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /e " & Chr$(34) & App.Path & "\backup_idml.reg" & Chr$(34) & " " & _
                Chr$(34) & "HKEY_CLASSES_ROOT\.idml" & Chr$(34), vbHide)
        
        SaveSetting App.EXEName, "Parameters", "FirstRunHappened", "YES"
        Ini1.Writing "MAIN", "SUPERSTARTER", App.Path & "\" & App.EXEName, sIniFile

   End If
   
   Call RegisterExtensions
   
If Not WINE_DETECTED Then
   
    sProg = UCase$(Ini1.Reading("MAIN", "OPEN_MAC_ON_PC", sIniFile))
    If (sProg <> "YES") And (sProg <> "NO") Then
      Ini1.Writing "MAIN", "OPEN_MAC_ON_PC", "NO", sIniFile
      sProg = "NO"
    End If
    
    Me.FormSys.mOpenMacFilesOnPC.Checked = (sProg = "YES")
    Me.FormSys.mUseVMWare.Enabled = (sProg = "NO")
      
    sProg = UCase$(Ini1.Reading("MAIN", "USE_VMWARE", sIniFile))
    If (sProg <> "YES") And (sProg <> "NO") Then
         Ini1.Writing "MAIN", "USE_VMWARE", "NO", sIniFile
        sProg = "NO"
    End If
   
    Me.FormSys.mUseVMWare.Checked = (sProg = "YES")
   End If
   
   'Get command line arguments
   'All arguments must be quoted!!!!!    "arg1" "arg2" ...
   vCommands = GetCommandLine(Me)
   
   
   If Not IsArray(vCommands) Then
      Call RegisterExtensions
      Me.Timer1.Enabled = True
      Exit Sub
   End If
   
   If Trim$(UCase$(vCommands(i))) = "UNINSTALL" Then
        On Error Resume Next
        UC.KillTaskByEXEName App.Title, App.hInstance
        UC.KillTaskByEXEName App.Title, App.hInstance
        UC.KillTaskByEXEName App.Title, App.hInstance
        UC.KillTaskByEXEName App.Title, App.hInstance
        UC.KillTaskByEXEName App.Title, App.hInstance
        DeleteSetting App.EXEName
        
        Set WSO = CreateObject("WScript.Shell")
        sTmp = vbNullString
        sTmp = WSO.RegRead("HKCR\SSIcon.SuperStarterIconHandler\Clsid\")
        WSO.RegDelete "HKLM\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" & sTmp
        'HKCR\SuperStarterCDR\shellex\IconHandler\
        'HKCR\SuperStarterCDR\shell\open\command\
        WSO.RegDelete "HKCR\SuperStarterCDR\shellex\IconHandler\"
        WSO.RegDelete "HKCR\SuperStarterCDR\shell\open\command\"
        WSO.RegDelete "HKCR\SuperStarterCDR\"
        
        WSO.RegDelete "HKCR\SuperStarterEPS\shellex\IconHandler\"
        WSO.RegDelete "HKCR\SuperStarterEPS\shell\open\command\"
        WSO.RegDelete "HKCR\SuperStarterEPS\"
        
        WSO.RegDelete "HKCR\SuperStarterAI\shellex\IconHandler\"
        WSO.RegDelete "HKCR\SuperStarterAI\shell\open\command\"
        WSO.RegDelete "HKCR\SuperStarterAI\"
        
        WSO.RegDelete "HKCR\SuperStarterTIF\shellex\IconHandler\"
        WSO.RegDelete "HKCR\SuperStarterTIF\shell\open\command\"
        WSO.RegDelete "HKCR\SuperStarterTIF\"
        
        WSO.RegDelete "HKCR\SuperStarterQXD\shellex\IconHandler\"
        WSO.RegDelete "HKCR\SuperStarterQXD\shell\open\command\"
        WSO.RegDelete "HKCR\SuperStarterQXD\"
        
        WSO.RegDelete "HKCR\SuperStarterQXP\shellex\IconHandler\"
        WSO.RegDelete "HKCR\SuperStarterQXP\shell\open\command\"
        WSO.RegDelete "HKCR\SuperStarterQXP\"
        
        WSO.RegDelete "HKCR\SuperStarterPDF\shellex\IconHandler\"
        WSO.RegDelete "HKCR\SuperStarterPDF\shell\open\command\"
        WSO.RegDelete "HKCR\SuperStarterPDF\"
        
        WSO.RegDelete "HKCR\SuperStarterINDD\shellex\IconHandler\"
        WSO.RegDelete "HKCR\SuperStarterINDD\shell\open\command\"
        WSO.RegDelete "HKCR\SuperStarterINDD\"
        
        WSO.RegDelete "HKCR\SuperStarterINX\shellex\IconHandler\"
        WSO.RegDelete "HKCR\SuperStarterINX\shell\open\command\"
        WSO.RegDelete "HKCR\SuperStarterINX\"
        
        WSO.RegDelete "HKCR\SuperStarterIDML\shellex\IconHandler\"
        WSO.RegDelete "HKCR\SuperStarterIDML\shell\open\command\"
        WSO.RegDelete "HKCR\SuperStarterIDML\"
        
        WSO.RegDelete "HKCR\SSIcon.SuperStarterIconHandler\"
        Set WSO = Nothing
        
        Call Shell("regedit.exe /s /c " & Chr$(34) & App.Path & "\backup_cdr.reg" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /c " & Chr$(34) & App.Path & "\backup_eps.reg" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /c " & Chr$(34) & App.Path & "\backup_ai.reg" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /c " & Chr$(34) & App.Path & "\backup_tif.reg" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /c " & Chr$(34) & App.Path & "\backup_qxd.reg" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /c " & Chr$(34) & App.Path & "\backup_qxp.reg" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /c " & Chr$(34) & App.Path & "\backup_pdf.reg" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /c " & Chr$(34) & App.Path & "\backup_indd.reg" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /c " & Chr$(34) & App.Path & "\backup_inx.reg" & Chr$(34), vbHide)
        Call Shell("regedit.exe /s /c " & Chr$(34) & App.Path & "\backup_idml.reg" & Chr$(34), vbHide)
        Kill Chr$(34) & App.Path & "\no_preview.ico" & Chr$(34)
        End
   End If
      
    sEPSFormat(0) = "Vector EPS"
    sEPSFormat(1) = "EPS DCS1"
    sEPSFormat(2) = "EPS DCS2"
    sEPSFormat(3) = "Photoshop EPS"
    
    sTIFFormat(0) = "TIFF PC"
    sTIFFormat(1) = "TIFF Macintosh"
    
    sTIFFColorMode(0) = "Black&White"
    sTIFFColorMode(1) = "Grayscale"
    sTIFFColorMode(2) = "RGB!!!"
    sTIFFColorMode(3) = "RGB Paletted!!!"
    sTIFFColorMode(4) = "Transparency Mask!!!"
    sTIFFColorMode(5) = "CMYK"
    sTIFFColorMode(6) = "YCbCr!!!"
    sTIFFColorMode(8) = "LAB!!!"
    
    sTIFFCompression(0) = "Unspecified!!!"
    sTIFFCompression(1) = "Uncompressed"
    sTIFFCompression(2) = "CCIT 1D"
    sTIFFCompression(3) = "CCIT Group 3 Fax"
    sTIFFCompression(4) = "CCIT Group 4 Fax"
    sTIFFCompression(5) = "LZW"
    sTIFFCompression(6) = "JPEG!!!"
    sTIFFCompression(7) = "Photoshop JPEG!!!"
    sTIFFCompression(8) = "Photoshop ZIP!!!"
    sTIFFCompression(9) = "PackedBits or unknown!!!"
    
    sPhotoshopColorMode(0) = ""
    sPhotoshopColorMode(1) = "Grayscale"
    sPhotoshopColorMode(2) = "LAB"
    sPhotoshopColorMode(3) = "RGB"
    sPhotoshopColorMode(4) = "CMYK"
    sPhotoshopColorMode(5) = "Multichannel"
    
    sTIFFUnits(0) = ""
    sTIFFUnits(1) = "unspecified"
    sTIFFUnits(2) = "inch"
    sTIFFUnits(3) = "cm"
   
   bEPSAnalizing = False
   
   
   'process each file name
   For i = LBound(vCommands) To UBound(vCommands)
      sProg = "": sCreator = ""
      'MsgBox vCommands(i)
      sExtension = UCase$(Right$(vCommands(i), 3))

'-----  QXD-QXP  ------
      If sExtension Like "QX?" Then
         sCreator = GetQXDVersion(vCommands(i))
         sTmp = sCreator
         sCreator = Replace$(sCreator, "Quark", vbNullString, , , vbTextCompare)
         sCreator = Replace$(sCreator, " Passport", vbNullString, , , vbTextCompare)
         sCreator = Replace$(sCreator, " MAC", vbNullString, , , vbTextCompare)
         Err.Clear
         lRes = 0
         If Val(sCreator) <= 9 Then
            lRes = XMsgBox("Open file(s) normally (with QUARK) - <YES>, FLIGHTCHEK it - <NO>", _
                   vbYesNoCancel + vbQuestion + vbDefaultButton1, "SuperStarter:  " & sTmp & " detected!", _
                   dlgEPS_TIF.imEXT.ListImages("qxd" & sCreator).Picture.Handle)
                    'GetIcon(vCommands(i), SHGFI_LARGEICON))
         Else
            lRes = XMsgBox("Open file(s) normally (with QUARK) - <YES>, FLIGHTCHEK it - <NO>", _
                   vbYesNoCancel + vbQuestion + vbDefaultButton1, "SuperStarter:  " & sTmp & " detected!", _
                    dlgEPS_TIF.imEXT.ListImages("qxd").Picture.Handle)
                    'GetIcon(vCommands(i), SHGFI_LARGEICON))
         End If
         If lRes = vbCancel Then
            GoTo NNEXT
         ElseIf lRes = vbYes Then
            sCreator = sTmp
         Else
            sCreator = "FLIGHTCHEK"
         End If

'-----  EPS  ------
      ElseIf sExtension = "EPS" Then
         lRes = MsgBox("Generate preview for this file?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "SuperStarter")
         If lRes = vbCancel Then
            GoTo NNEXT
         ElseIf lRes = vbYes Then
            If GetGflaxImagePreview(vCommands(i)) Then frmPreview.Show: SetOnTopWindow frmPreview.hwnd, True
         End If
         
        bEPSAnalizing = True
        sINFO = vbNullString
        tEPS = GetEPSInfo(vCommands(i), Me)
        bEPSAnalizing = False
        If tEPS.EPSType <> epsNonEPSImage Then
            sINFO = sINFO & "EPS format: " & sEPSFormat(tEPS.EPSType)
            sINFO = sINFO & ", EPS main creator: " & tEPS.EPSCreator
            If Len(tEPS.EPSDocProcessColors) > 0 Then
                sINFO = sINFO & ", process colors: " & tEPS.EPSDocProcessColors
            End If
            If Len(tEPS.EPSDocCustomColors) > 0 Then
                    sINFO = sINFO & ", custom colors: " & tEPS.EPSDocCustomColors
            End If
            
            If tEPS.EPS_BPS > 0 Then sINFO = sINFO & ", color depth (bit): " & tEPS.EPS_BPS
            If tEPS.EPS_BPS = bpc1 Then
                sINFO = sINFO & " (Black&White)"
            Else
                sINFO = sINFO & " (" & sPhotoshopColorMode(tEPS.EPS_PhotoshopMode) & ")"
            End If
            If (tEPS.EPSType = epsDCS1) Or (tEPS.EPSType = epsDCS2) Then
                sINFO = sINFO & ", DCS info: " & Join(tEPS.EPS_DCSPlates, ", ")
            End If
            
        End If
        If InStr(tEPS.EPSCreator, "Photoshop") > 0 Then
            lHandleI = dlgEPS_TIF.imEXT.ListImages("pseps").Picture.Handle
            If tEPS.EPSType = epsDCS1 Then
                lHandleI = dlgEPS_TIF.imEXT.ListImages("dcs1").Picture.Handle
            ElseIf tEPS.EPSType = epsDCS2 Then
                lHandleI = dlgEPS_TIF.imEXT.ListImages("dcs2").Picture.Handle
            End If
        ElseIf InStr(tEPS.EPSCreator, "Illustrator") > 0 Then
            lHandleI = dlgEPS_TIF.imEXT.ListImages("eps").Picture.Handle
        ElseIf InStr(tEPS.EPSCreator, "Corel") > 0 Then
            lHandleI = dlgEPS_TIF.imEXT.ListImages("cdreps").Picture.Handle
        Else
            lHandleI = dlgEPS_TIF.imEXT.ListImages("unknowneps").Picture.Handle
        End If
        'lHandleI = GetIcon(vCommands(i), SHGFI_LARGEICON)
         lRes = XMsgBox(sINFO & vbCrLf & vbCrLf & _
                "Open file(s) normally (with Illustrator or Photoshop) - <YES>, send to DISTILLER - <NO>", _
                vbYesNoCancel + vbQuestion + vbDefaultButton1, "SuperStarter:  " & tEPS.EPSCreator & " detected!", _
                lHandleI)
         frmPreview.Hide
         If lRes = vbCancel Then
            GoTo NNEXT
         ElseIf lRes = vbYes Then
            If InStr(1, tEPS.EPSCreator, "Adobe Illustrator(R) 8.0") > 0 Then
                sCreator = tEPS.EPSCreator
            ElseIf tEPS.EPS_AI8_Creator_found Then
                sCreator = tEPS.EPSCreator
            Else
                bEPSAnalizing = True
                sCreator = GetEPSCreator(vCommands(i), Me)
                bEPSAnalizing = False
            End If
         Else
            sCreator = "DISTILLER"
         End If

'-----  AI  ------
      ElseIf sExtension = ".AI" Then
         If Ensure_file_is_PDFcompatible(vCommands(i)) <> 0 Then
            lRes = MsgBox("AI file seems to be PDF-compatible." & vbCrLf & _
               "Try to open it in Acrobat?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "SuperStarter")
            If lRes = vbYes Then sCreator = "ACROBAT": GoTo NEXT_AI
            If lRes = vbCancel Then GoTo NNEXT
         End If
            bEPSAnalizing = True
            sINFO = vbNullString
           ' tEPS = GetEPSInfo(vCommands(i), Me)
           ' sCreator = GetEPSCreator(vCommands(i), Me)
           
           sCreator = GetAIVersion(vCommands(i), Me)
            bEPSAnalizing = False
           
NEXT_AI:

'-----  PDF  ------
      ElseIf sExtension = "PDF" Then
            lHandleI = dlgEPS_TIF.imEXT.ListImages("pdf").Picture.Handle
            'lHandleI = GetIcon(vCommands(i), SHGFI_LARGEICON)
            lRes = XMsgBox("Try to detect Illustrator private data inside PDF - <YES>, open PDF normally in Acrobat - <NO>", _
                vbYesNoCancel + vbQuestion + vbDefaultButton2, "SuperStarter:  Adobe PDF document!", _
                lHandleI)
'         lRes = ZMsgBox("SuperStarter", "Try to detect Illustrator private data inside PDF - <YES>, open PDF normally in Acrobat - <NO>", _
                "Adobe PDF document", "PDF1.4", dlgEPS_TIF.imEXT.ListImages("pdf").Picture, True)
         If lRes = vbYes Then
            'bEPSAnalizing = True
            'sINFO = vbNullString
            'tEPS = GetEPSInfo(vCommands(i), Me, True)
            'bEPSAnalizing = False
            sCreator = GetAIVersion(vCommands(i))
            If Val(sCreator) > 0 Then
                sCreator = "Adobe Illustrator(R) " & Format$(sCreator, "#0.0")
            Else
                sCreator = vbNullString
            End If
         End If
         If lRes = vbNo Then sCreator = "ACROBAT"
         'If lRes = vbNo Then sCreator = "PDF1." & Trim$(CInt(Ensure_file_is_PDFcompatible(vCommands(i))))
         If lRes = vbCancel Then GoTo NNEXT

'-----  TIF  ------
      ElseIf sExtension = "TIF" Then
         lRes = MsgBox("Generate preview for this file?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "SuperStarter")
         If lRes = vbCancel Then
            GoTo NNEXT
         ElseIf lRes = vbYes Then
            If GetGflaxImagePreview(vCommands(i)) Then frmPreview.Show: SetOnTopWindow frmPreview.hwnd, True
         End If
        sINFO = vbNullString
        tTIF = GetTIFInfo(vCommands(i))
        If tTIF.TIFFBytesOrder <> tifNonTIFImage Then
            sINFO = "TIF color mode: " & sTIFFColorMode(tTIF.TIFFColorMode)
            If tTIF.TIFFAlfaChannels > 0 Then sINFO = sINFO & _
                    " + " & CStr(tTIF.TIFFAlfaChannels) & " alpha channels"
            sINFO = sINFO & ", format: " & sTIFFormat(tTIF.TIFFBytesOrder)
            sINFO = sINFO & "," & vbCrLf & _
                    "bits per channel: " & tTIF.TIFFBitsPerSample
            
            If tTIF.TIFFXRes <> tTIF.TIFFYRes Then
                sINFO = sINFO & "," & vbCrLf & _
                    "X res: " & Format$(tTIF.TIFFXRes, "0.0") & " pixels/" & sTIFFUnits(tTIF.TIFFUnits) & _
                    ", Y res: " & Format$(tTIF.TIFFXRes, "0.0") & " pixels/" & sTIFFUnits(tTIF.TIFFUnits)
            Else
                sINFO = sINFO & "," & vbCrLf & _
                    "res: " & tTIF.TIFFXRes & " pixels/" & sTIFFUnits(tTIF.TIFFUnits)
            End If
            If tTIF.TIFFUnits < 2 Then
                sINFO = sINFO & "," & vbCrLf & _
                    "width: " & tTIF.TIFFWidth & " pixels, height: " & tTIF.TIFFHeight & " pixels"
            ElseIf tTIF.TIFFUnits = tifINCH Then
                sINFO = sINFO & "," & vbCrLf & _
                    "width: " & Format$(25.4 * tTIF.TIFFWidth / tTIF.TIFFXRes, "0.0") & " mm, height: " & _
                    Format$(25.4 * tTIF.TIFFHeight / tTIF.TIFFYRes, "0.0") & " mm"
            Else
                sINFO = sINFO & "," & vbCrLf & _
                    "width: " & Format$(10 * tTIF.TIFFWidth / tTIF.TIFFXRes, "0.0") & " mm, height: " & _
                    Format$(10 * tTIF.TIFFHeight / tTIF.TIFFYRes, "0.0") & " mm"
            End If
            
            sINFO = sINFO & "," & vbCrLf & "compression: " & sTIFFCompression(tTIF.TIFFCompression)
            
            If Len(tTIF.TIFFProfile) > 0 Then sINFO = sINFO & vbCrLf & vbCrLf & "Embedded color profile: '" & tTIF.TIFFProfile & "'"
        End If
         lHandleI = dlgEPS_TIF.imEXT.ListImages("tif").Picture.Handle
         If tTIF.TIFFColorMode = tifRGB Then lHandleI = dlgEPS_TIF.imEXT.ListImages("tifrgb").Picture.Handle
         If tTIF.TIFFColorMode = tifCMYK Then lHandleI = dlgEPS_TIF.imEXT.ListImages("tifcmyk").Picture.Handle
         'lHandleI = GetIcon(vCommands(i), SHGFI_LARGEICON)
            lRes = XMsgBox(sINFO & vbCrLf & vbCrLf & _
                   "Open file(s) with Photoshop - <YES>, convert to PDF - <NO>", _
                   vbYesNoCancel + vbQuestion + vbDefaultButton1, "SuperStarter:  TIFF image detected!", _
                   lHandleI)
        frmPreview.Hide
' extract TIFF profile
         If Len(tTIF.TIFFProfile) > 0 Then
            lResTemp = XMsgBox( _
                    "Dou you want to extract embedded profile in current directory?", _
                   vbYesNoCancel + vbQuestion + vbDefaultButton2, "SuperStarter:  TIFF image profile", _
                   lHandleI)
         End If
         If lResTemp = vbYes Then
            If UBound(tTIF.TIFFProfileData) > 1 Then
                On Error Resume Next
                Set FLSO = CreateObject("Scripting.FileSystemObject")
                tTIF.TIFFProfile = Replace(tTIF.TIFFProfile, ":", "_")
                tTIF.TIFFProfile = Replace(tTIF.TIFFProfile, "/", "_")
                tTIF.TIFFProfile = Replace(tTIF.TIFFProfile, "\", "_")
                tTIF.TIFFProfile = Replace(tTIF.TIFFProfile, "|", "_")
                If UC.FileExists(FLSO.GetParentFolderName(vCommands(i)) & "\" & tTIF.TIFFProfile & ".icc") = False Then
                    Open (FLSO.GetParentFolderName(vCommands(i)) & "\" & tTIF.TIFFProfile & ".icc") For Binary Access Write As #223
                        Put #223, , tTIF.TIFFProfileData
                    Close #223
                Else
                    Open (FLSO.GetParentFolderName(vCommands(i)) & "\" & tTIF.TIFFProfile & _
                                Format$(Now(), "_yyyymmdd_hhnnss_") & ".icc") For Binary Access Write As #223
                        Put #223, , tTIF.TIFFProfileData
                    Close #223
                End If
                Set FLSO = Nothing
                Err.Clear
                On Error GoTo 0
            End If
        End If
        
         If lRes = vbCancel Then
            GoTo NNEXT
         ElseIf lRes = vbYes Then
            sCreator = "TIFF image"
         ElseIf lRes = vbNo Then
            sCreator = "ACROBAT"
         End If
'-----  INDD  ------
      ElseIf UCase$(Right$(vCommands(i), 4)) = "INDD" Then
         sCreator = GetINDDVersion(vCommands(i))

'-----  IDML  ------
      ElseIf UCase$(Right$(vCommands(i), 4)) = "IDML" Then
         sCreator = GetINXVersion(vCommands(i))

'-----  INX  ------
      ElseIf sExtension = "INX" Then
         sCreator = GetINXVersion(vCommands(i))
         
'-----  INX  ------
      ElseIf sExtension = "CDR" Then
         sCreator = "CorelDRAW version " & CStr(GetCDRVersion(vCommands(i)))
         
      End If
      
DO_CHECK_PROGRAMM:
      
      If Len(sCreator) > 0 Then
         On Error Resume Next
         sProg = UCase$(Ini1.Reading("MAIN", "OPEN_MAC_ON_PC", sIniFile))
         If (sCreator Like "*MAC") And (sProg <> "YES") Then ' we have unchecked parameter "Open MAC files on PC"
            MsgBox "You're trying to open MAC version of document (created by " & Chr$(34) & sCreator & Chr$(34) & ")," & _
            vbCrLf & "but item 'Open MAC files on PC' in SuperStarter tray menu is UNCHECKED." & vbCrLf & _
            "So, you have better to open this file under Mac OS X to avoid serious fonts troubles." & vbCrLf & _
            "If you still wont open file on this PC, you must CHEK item 'Open MAC files on PC' before.", vbExclamation, "SuperStarter"
            GoTo NNEXT
         End If
         sProg = Ini1.Reading("MAIN", sCreator, sIniFile)
         If (Err.Number <> 0) Or (Len(sProg) = 0) Or (UC.FileExists(sProg) = False) Then
            Err.Clear
            On Error GoTo 0
            MsgBox "Select program to file created by " & Chr$(34) & sCreator & Chr$(34) & " open with!", vbInformation, "SuperStarter"
            sProg = ShowOpenFileDialog("Executables (*.exe)|*.exe", "exe", , OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY, Me.hwnd)
            If Len(sProg) > 0 Then
               Ini1.Writing "MAIN", sCreator, sProg, sIniFile
            Else
               GoTo NNEXT
            End If
         ElseIf sCreator = "ACROBAT" Then ' here we shall to ensure that selected Acrobat understands PDF version of selected file
            iPDFVersion = Ensure_file_is_PDFcompatible(vCommands(i))
            'PDF version + 1 = max Acrobat verison
            Set FLSO = CreateObject("Scripting.FileSystemObject")
            sAcrobatVersion = FLSO.GetFileVersion(sProg)
            Set FLSO = Nothing
            iAcrobatVersion = CInt(Val(sAcrobatVersion))
            If iAcrobatVersion - iPDFVersion < 1 Then
                'First we'll check for existing definitions of programm for PDF these version
                sCreator = "PDF1." & Trim$(CStr(iPDFVersion))
                sProg = Ini1.Reading("MAIN", sCreator, sIniFile)
                lHandleI = dlgEPS_TIF.imEXT.ListImages("pdf").Picture.Handle
                'lHandleI = GetIcon(vCommands(i), SHGFI_LARGEICON)
                If (Err.Number <> 0) Or (Len(sProg) = 0) Or (UC.FileExists(sProg) = False) Then
                    Err.Clear
                    On Error GoTo 0
                    lRes = XMsgBox("PDF version of selected file (" & sCreator & ") is newer than" & vbCrLf & _
                        "selected version of Acrobat (" & sAcrobatVersion & ") can operate!" & vbCrLf & _
                        "If you wanna select another Acrobat programm (i.e. portable version)" & vbCrLf & _
                        "for opening this version of PDF - press <YES>," & vbCrLf & _
                        "otherwise (open PDF in older Acrobat) - press <NO>", _
                        vbYesNoCancel + vbQuestion + vbDefaultButton2, "SuperStarter:  Adobe PDF newest version detected!", lHandleI)
                    If lRes = vbCancel Then GoTo NNEXT
                    If lRes = vbYes Then GoTo DO_CHECK_PROGRAMM
                    If lRes = vbNo Then
                        sCreator = "ACROBAT"
                        sProg = Ini1.Reading("MAIN", sCreator, sIniFile)
                    End If
                Else 'Ask user to opne with older or newer Acrobat
                    lRes = XMsgBox("PDF version of selected file (" & sCreator & ") is newer than" & vbCrLf & _
                        "our usual version of Acrobat (" & sAcrobatVersion & ")." & vbCrLf & _
                        "But you've selected another Acrobat programm for this newer PDF version" & vbCrLf & _
                        "(" & sProg & ")." & vbCrLf & _
                        "Open this version of PDF in newer Acrobat - press <YES>," & vbCrLf & _
                        "otherwise (open PDF in usual Acrobat) - press <NO>", _
                        vbYesNoCancel + vbQuestion + vbDefaultButton1, "SuperStarter:  Adobe PDF document!", lHandleI)
                    If lRes = vbCancel Then GoTo NNEXT
                    If lRes = vbNo Then
                        sCreator = "ACROBAT"
                        sProg = Ini1.Reading("MAIN", sCreator, sIniFile)
                    End If
                End If
            End If
         End If
         
         Set WSO = CreateObject("WScript.Shell")
         WSO.Run Chr$(34) & sProg & Chr$(34) & " " & Chr$(34) & vCommands(i) & Chr$(34), 1
         Set WSO = Nothing
         
         Err.Clear
         On Error GoTo 0
         'Call Shell(Chr$(34) & sProg & Chr$(34) & " " & Chr$(34) & vCommands(i) & Chr$(34), vbNormalFocus)
      Else
         MsgBox "Creator not identified in " & vCommands(i) & "! Try to open file manually.", vbExclamation, "ERROR"
      End If
NNEXT:
   Next i
   Close
   
If App.PrevInstance Then
    On Error Resume Next
    Unload Me.FormSys
    End
End If

'CreateMutex 0&, Me.hwnd, "SuperStarterMutexWorks"
Me.Timer1.Enabled = True


End Sub


Private Sub Timer1_Timer()
   Call RegisterExtensions
   If App.PrevInstance Then
        On Error Resume Next
        Unload Me.FormSys
        End
   End If
End Sub

Private Sub tmr_TempCleaner_Timer()
    Dim tFS As Object, tFSO As Object, tFSF As Object, tFSOS As Object, iiii As Long
    If Minute(Now()) <> 0 Then Exit Sub
    On Error Resume Next

    Set tFS = CreateObject("Scripting.FileSystemObject")
    Set tFSO = tFS.GetFolder(Environ$("TEMP"))
    For iiii = 1 To tFSO.Files.Count
        For Each tFSF In tFSO.Files
            If Not (UCase$(tFSF.ShortName) Like "VB???.TMP") Then
                tFSF.Delete True
            Else
                If DateDiff("Y", tFSF.DateCreated, Now()) > 2 Then tFSF.Delete True
            End If
        Next tFSF
    Next iiii
    Err.Clear
    For iiii = 1 To tFSO.SubFolders.Count
        For Each tFSOS In tFSO.SubFolders
            tFSOS.Delete True
        Next tFSOS
    Next iiii
    Err.Clear
    
    
    Set tFSO = tFS.GetFolder(Environ$("TMP"))
    For iiii = 1 To tFSO.Files.Count
        For Each tFSF In tFSO.Files
            tFSF.Delete True
        Next tFSF
    Next iiii
    Err.Clear
    For iiii = 1 To tFSO.SubFolders.Count
        For Each tFSOS In tFSO.SubFolders
            tFSOS.Delete True
        Next tFSOS
    Next iiii
    Err.Clear
    
    On Error GoTo 0
End Sub
