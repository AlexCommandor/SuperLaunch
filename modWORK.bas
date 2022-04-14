Attribute VB_Name = "modWORK"
Option Explicit

Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CreateIconFromResourceEx Lib "user32" (pbIconBits As Byte, ByVal cbIconBits As Long, _
            ByVal fIcon As Long, ByVal dwVersion As Long, cxDesired As Long, cyDesired As Long, uFlags As Long) As Long
Private Declare Function CreateIconFromResource Lib "user32" (pbIconBits As Byte, ByVal cbIconBits As Long, _
            ByVal fIcon As Long, ByVal dwVersion As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function wine_get_version Lib "ntdll" () As Long

Public minIconWidth As Long, hwIcon As Long, lIconIndex As Long

Public WINE_DETECTED As Boolean

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
 
Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
 
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50


Public Sub Main()
    On Error Resume Next
    
Dim ret As Long

    ret = wine_get_version
    
    If Err.Number <> 0 Then
        Err.Clear
        WINE_DETECTED = False
    Else
        'MsgBox "WINE detected!!! Program may produce unexpected result!", vbCritical + vbOKOnly, "SuperStarter"
        WINE_DETECTED = True
    End If
    
    ret = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If ret <> 0 Then
        IsWow64Process GetCurrentProcess, ret
        If ret <> 0 Then
            MsgBox "This application is working ONLY on an 32-bit Windows, not on 64-bit!!! Bye!"
            End
        End If
    End If
    
    If (Not App.PrevInstance) And (Not WINE_DETECTED) Then
    
        minIconWidth = GetSystemMetrics(SM_CXSMICON)

        Select Case minIconWidth
            Case Is = 16
                lIconIndex = 28
            Case Is = 20
                lIconIndex = 27
            Case Is = 24
                lIconIndex = 26
            Case Is = 28
                lIconIndex = 25
            Case Else
                lIconIndex = 24
        End Select
        hwIcon = LoadIconFromMultiRES("TRAYICON", 101, lIconIndex, , , True)
    
        Set frmMain.FormSys = New FrmSysTray
        Load frmMain.FormSys
        Set frmMain.FormSys.FSys = frmMain
        frmMain.FormSys.Tooltip = "SuperStarter"
        frmMain.FormSys.Interval = 500
        If hwIcon = 0 Then
            frmMain.FormSys.TrayIcon = frmMain
        Else
            frmMain.FormSys.TrayIcon = hwIcon
        End If
        
    'Else
        'Load frmMain
    End If
    Load frmMain
    Err.Clear
    On Error GoTo 0
End Sub

Public Function LoadIconFromMultiRES(ResID, ResName, IconIndex, Optional PixelsX = 16, Optional PixelsY = 16, Optional bDefaultSize As Boolean = False) As Long
    Const ICRESVER As Long = &H30000
    Dim IconFile As Long
    Dim IconRes() As Byte
    Dim hIcon As Long
    Dim lDirPos As Long, lSize As Long, lBMPpos As Long
    
    'Load the icon from desired resource
    IconRes = LoadResData(ResName, ResID)
    
    'Grab the chosen icon from the file Index; 0 = 1st Icon
    lDirPos = 6 + (IconIndex) * 16

    CopyMemory lSize, IconRes(lDirPos + 8), 4&

    CopyMemory lBMPpos, IconRes(lDirPos + 12), 4&

    'Create the Icon File
    If Not bDefaultSize Then
        hIcon = CreateIconFromResourceEx(IconRes(lBMPpos), UBound(IconRes) - 21&, True, ICRESVER, PixelsX, PixelsY, 0&)
    Else
        hIcon = CreateIconFromResource(IconRes(lBMPpos), UBound(IconRes) - 21&, True, ICRESVER)
    End If
    
    'If there is data, set the icon to the desired source
    If hIcon > 0 Then LoadIconFromMultiRES = hIcon
End Function

Private Function GetSystemSerialNumber() As String
Dim mother_boards As Variant
Dim board As Variant
Dim wmi As Variant
Dim serial_numbers As String

    ' Get the Windows Management Instrumentation object.
    Set wmi = GetObject("WinMgmts:")

    ' Get the "base boards" (mother boards).
    Set mother_boards = wmi.InstancesOf("Win32_BaseBoard")
    For Each board In mother_boards
        serial_numbers = serial_numbers & ", " & _
            board.SerialNumber
    Next board
    If Len(serial_numbers) > 0 Then serial_numbers = _
        Mid$(serial_numbers, 3)

    GetSystemSerialNumber = serial_numbers
End Function

Private Function GetHDDSerialNumber() As String
Dim hdds As Variant
Dim disk As Variant
Dim wmi As Variant
Dim serial_numbers As String

    ' Get the Windows Management Instrumentation object.
    Set wmi = GetObject("WinMgmts:")

    ' Get the "base boards" (mother boards).
    Set hdds = wmi.InstancesOf("Win32_PhysicalMedia")
    For Each disk In hdds
        serial_numbers = serial_numbers & ", " & _
            Trim$(disk.SerialNumber)
    Next disk
    If Len(serial_numbers) > 0 Then serial_numbers = _
        Mid$(serial_numbers, 3)

    GetHDDSerialNumber = serial_numbers
End Function

Private Function GetDriveLetters() As String
    Dim ComputerName As String, wmiServices As Object, wmiDiskDrives As Object, wmiDiskDrive As Object
    Dim wmiDiskPartitions As Object, wmiDiskPartition As Object, query As String
    Dim wmiLogicalDisks As Object, wmiLogicalDisk As Object
ComputerName = "."
Set wmiServices = GetObject( _
    "winmgmts:" & _
        "{impersonationLevel=impersonate}!\\" & _
        ComputerName & "\root\cimv2")
' Get physical disk drive
Set wmiDiskDrives = wmiServices.ExecQuery( _
    "SELECT * FROM Win32_DiskDrive")

For Each wmiDiskDrive In wmiDiskDrives
     Debug.Print "Disk drive Caption: " _
        & wmiDiskDrive.Name _
        & vbNewLine & "DeviceID: " _
        & " (" & wmiDiskDrive.DeviceID & ")" _
        & vbNewLine & "DeviceID: " _
        & " (" & wmiDiskDrive.DeviceID & ")"

    'Use the disk drive device id to
    ' find associated partition
    query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" _
        & wmiDiskDrive.DeviceID & "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
    Set wmiDiskPartitions = wmiServices.ExecQuery(query)

    For Each wmiDiskPartition In wmiDiskPartitions
        'Use partition device id to find logical disk
        Set wmiLogicalDisks = wmiServices.ExecQuery _
            ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" _
             & wmiDiskPartition.DeviceID & "'} WHERE AssocClass = Win32_LogicalDiskToPartition")

        For Each wmiLogicalDisk In wmiLogicalDisks
            Debug.Print "Drive letter associated" _
                & " with disk drive = " _
                & wmiDiskDrive.Caption _
                & wmiDiskDrive.DeviceID _
                & vbNewLine & " Partition = " _
                & wmiDiskPartition.DeviceID _
                & vbNewLine & " is " _
                & wmiLogicalDisk.DeviceID
        Next
    Next
Next
End Function

 
Private Function GetCpuId() As String
Dim computer As String
Dim wmi As Variant
Dim processors As Variant
Dim cpu As Variant
Dim cpu_ids As String

    computer = "."
    Set wmi = GetObject("winmgmts:" & _
        "{impersonationLevel=impersonate}!\\" & _
        computer & "\root\cimv2")
    Set processors = wmi.ExecQuery("Select * from " & _
        "Win32_Processor")

    For Each cpu In processors
        cpu_ids = cpu_ids & ", " & cpu.ProcessorId
    Next cpu
    If Len(cpu_ids) > 0 Then cpu_ids = Mid$(cpu_ids, 3)

    GetCpuId = cpu_ids
End Function

Public Function GetEPSCreator(ByVal sFName As String, Optional ByRef frmOWNER As Form = Nothing) As String
   'Dim bBegin1K(1 To 1024 * 31) As Byte
   Dim vEPS_STRINGS As Variant
   Dim FN As Integer, sFile As String
   Dim sTmp As String, lInstr As Long, lInstr2 As Long, i As Long, lFLen As Long
   Dim nCount As Integer, nFirst As Long, nSecond As Long, sFirst As String, sSecond As String
   Dim j As Long
   
   sFile = sFName
   FN = FreeFile
   lFLen = FileLen(sFile)
   nCount = 0: nFirst = 0: nSecond = 0
   
' If Not (frmOWNER Is Nothing) Then frmOWNER.Show
' If Not (frmOWNER Is Nothing) Then frmOWNER.Caption = "Determining correct creator of EPS file, please wait..."
' If Not (frmOWNER Is Nothing) Then frmOWNER.ProgressBar1.Min = 0
' If Not (frmOWNER Is Nothing) Then frmOWNER.ProgressBar1.Max = lFLen
    
   'vEPS_STRINGS = ReadEpsFileToStringArray(sFName, frmOWNER, _
            "Determining correct creator of EPS file, please wait...")
    vEPS_STRINGS = vEPS_DATA
   'Open sFile For Binary As #FN
      'Do
         'Get #FN, , bBegin1K
         'sTmp = StrConv(bBegin1K, vbUnicode)
    For j = LBound(vEPS_STRINGS) To UBound(vEPS_STRINGS)
        
        sTmp = vEPS_STRINGS(j)
        
         lInstr = InStr(1, sTmp, "%%Creator:")
         lInstr2 = InStr(1, sTmp, "%AI9_PrivateDataBegin")
         If (lInstr > 0) And nCount = 0 Then
            nCount = 1
            nFirst = lInstr
            sFirst = sTmp
            If InStr(sTmp, "Adobe Illustrator 8") > 0 Then GoTo EPS8
         End If
         If lInstr2 > 0 Then
'            If Seek(FN) > 1024 * 31 Then
'                Seek FN, Seek(FN) - 1024 * 31 + lInstr2
'            Else
'                Seek FN, lInstr2
'            End If
'            Get #FN, , bBegin1K
'            sTmp = StrConv(bBegin1K, vbUnicode)
            lInstr = InStr(1, sTmp, "%%Creator:")
            If (lInstr > 0) Then
               nCount = 1
               nFirst = lInstr
               sFirst = sTmp
               'Exit Do
               Exit For
            End If
         End If
'         Seek FN, Seek(FN) - 100
'         If Seek(FN) > lFLen Then
'            If Not (frmOWNER Is Nothing) Then frmOWNER.ProgressBar1.Value = lFLen
'         Else
'            If Not (frmOWNER Is Nothing) Then frmOWNER.ProgressBar1.Value = Seek(FN)
'         End If
         DoEvents
'      Loop Until (Seek(FN) >= lFLen) ' Or (nCount = 2)
EPS8:
'   Close #FN
            'If frmOWNER.bEPSAnalizing = False Then GetEPSCreator = "Analyzing aborted!!!": Exit Function
    Next j
' If Not (frmOWNER Is Nothing) Then frmOWNER.Hide
   lInstr = 0
   If nFirst > 0 Then lInstr = nFirst: sTmp = sFirst
'   If nSecond > 0 Then lInstr = nSecond: sTmp = sSecond
   If lInstr > 0 Then
      sTmp = Right$(sTmp, Len(sTmp) - lInstr - 9)
      lInstr = InStr(1, sTmp, Chr$(13), vbBinaryCompare)
      If lInstr > 0 Then
         sTmp = Left$(sTmp, lInstr - 1)
      End If
      lInstr = InStr(1, sTmp, "x", vbBinaryCompare)
      If lInstr > 0 Then
         sTmp = Left$(sTmp, lInstr - 1)
      End If
      GetEPSCreator = Trim$(sTmp)
   Else
      GetEPSCreator = ""
   End If
End Function


Public Function GetQXDVersion(ByVal sFName As String) As String
   Dim bHeader(1 To 12) As Byte, FN As Integer, sFile As String
   Dim sTmp As String
   sFile = sFName
   FN = FreeFile
   Open sFile For Binary As #FN
      Get #FN, , bHeader
   Close #FN
   If bHeader(9) + bHeader(10) = &H41 Then
      sTmp = "Quark4"
   ElseIf bHeader(9) + bHeader(10) = &H42 Then
      sTmp = "Quark5"
   ElseIf bHeader(9) + bHeader(10) = &H3F Then
      sTmp = "Quark3"
   ElseIf bHeader(9) + bHeader(10) = &H43 Then
      sTmp = "Quark6"
   ElseIf bHeader(9) + bHeader(10) = &H44 Then
      sTmp = "Quark7"
   ElseIf bHeader(9) + bHeader(10) = &H5F Then
      sTmp = "Quark3 Passport"
   ElseIf (bHeader(9) + bHeader(10) >= &H61) And (bHeader(9) + bHeader(10) <= &H65) Then
      sTmp = "Quark" & CStr(bHeader(9) + bHeader(10) - &H5D) & " Passport"
   ElseIf (bHeader(9) + bHeader(10) >= &H67) Then
      sTmp = "Quark" & CStr(bHeader(9) + bHeader(10) - &H5E) & " Passport"
'   ElseIf bHeader(9) + bHeader(10) = &H62 Then
'      sTmp = "Quark5 Passport"
'   ElseIf bHeader(9) + bHeader(10) = &H63 Then
'      sTmp = "Quark6 Passport"
'   ElseIf bHeader(9) + bHeader(10) = &H64 Then
'      sTmp = "Quark7 Passport"
'   ElseIf bHeader(9) + bHeader(10) = &H65 Then
'      sTmp = "Quark8 Passport"
'   ElseIf bHeader(9) + bHeader(10) = &H66 Then
'      sTmp = "Quark9 Passport"
   Else
      sTmp = ""
   End If
   If Len(sTmp) > 0 And (bHeader(3) = &H4D) And (bHeader(4) = &H4D) Then sTmp = sTmp & " MAC"
   'If Len(sTmp) > 0 And (bHeader(3) = &H49) And (bHeader(4) = &H49) Then sTmp = sTmp & " PC"
   GetQXDVersion = sTmp
End Function

Public Function GetINDDVersion(ByVal sFName As String) As String
   Dim bHeader(0 To 47) As Byte, FN As Integer, sFile As String
   Dim sTmp As String, varResult As Variant
   sFile = sFName
   FN = FreeFile
   Open sFile For Binary As #FN
      Get #FN, , bHeader
   Close #FN
   
   'NOTE - bytes in bHeader counting from ZERO!!!!   -   from 0 to 47 (from &H0 to &H2F)
   
   If bHeader(&H19) = &H70 And bHeader(&H1A) = &HF Then ' PC indd   (bytes 25 & 26)
      If CStr(bHeader(&H21) + bHeader(&H22)) = 0 Then 'Indesign version like 5.0, 6.0, 7.0 etc  (bytes 33 & 34)
        sTmp = "Indesign" & CStr(bHeader(&H1C) + bHeader(&H1D)) & " PC"  ' bytes 28 & 29
      Else
        sTmp = "Indesign" & CStr(bHeader(&H1C) + bHeader(&H1D)) & "." & _
            CStr(bHeader(&H20) + bHeader(&H21)) & " PC" 'Indesign version like 5.2, 6.3, 7.5 etc  (bytes 32 & 33)
      End If
   ElseIf bHeader(&H1B) = &HF And bHeader(&H1C) = &H70 Then  ' MAC indd  (bytes 27 & 28)
      If CStr(bHeader(&H23) + bHeader(&H24)) = 0 Then 'Indesign version like 5.0, 6.0, 7.0 etc  (bytes 35 & 36)
        sTmp = "Indesign" & CStr(bHeader(&H20) + bHeader(&H21)) & " MAC"  '(bytes 32 & 33)
      Else
        sTmp = "Indesign" & CStr(bHeader(&H20) + bHeader(&H21)) & "." & _
            CStr(bHeader(&H23) + bHeader(&H24)) & " MAC" 'Indesign version like 5.2, 6.3, 7.5 etc (bytes 37 & 38)
      End If
   Else
      sTmp = ""
   End If
   
   frmMain.bEPSAnalizing = True
   varResult = ReadEpsFileToStringArray(sFName, , , 100, frmMain, "Analizing INDD file, please wait...", ".DFONT")
   frmMain.bEPSAnalizing = False
   If UBound(varResult) > 0 Then
      sTmp = Replace$(sTmp, " PC", " MAC")
   Else
      sTmp = Replace$(sTmp, " MAC", " PC")
   End If
   
   GetINDDVersion = sTmp
End Function

Public Function GetINXVersion(ByVal sFName As String) As String
    Dim xmlDoc As Object, currNode As Object, sTmp As String, gTmp As Single
    Dim ZZ As Cls_Zip, bBBB() As Variant, i As Long
    Dim bData() As Boolean, iFNN As Integer, bFileHeader(1 To 15) As Byte, bZIP As Boolean
    
    On Error Resume Next
    iFNN = FreeFile()
    Open sFName For Binary Access Read Shared As #iFNN
        Get #iFNN, , bFileHeader
    Close #iFNN

    bZIP = False
    If bFileHeader(1) = &H50 And bFileHeader(2) = &H4B Then ' here we have IDML (not INX) that is packed by ZIP algorithm
        bZIP = True
        Set ZZ = New Cls_Zip
        ZZ.Get_Contents sFName
        If ZZ.FileCount > 0 Then
            ReDim bData(1 To ZZ.FileCount)
            For i = 1 To ZZ.FileCount
                If UCase$(ZZ.FileName(i)) = "DESIGNMAP.XML" Then
                    bData(i) = True
                Else
                    bData(i) = False
                End If
            Next i
        End If
        ReDim bBBB(1 To 1)
        ZZ.UnPackToBuffer bData, bBBB
        Set ZZ = Nothing
        sTmp = StrConv(bBBB(2), vbUnicode)
    End If
    
    Set xmlDoc = CreateObject("MSXML.DOMDocument")
    xmlDoc.async = False
    If bZIP = False Then
        xmlDoc.Load sFName
    Else
        If sTmp <> vbNullString Then xmlDoc.LoadXML sTmp
    End If
    If (xmlDoc.parseError.errorCode <> 0) Then 'Error parsing XML - maybe not XML? :(
        GetINXVersion = vbNullString
        Exit Function
    Else
        Set currNode = xmlDoc.childNodes(1)
        sTmp = UCase$(currNode.Text)
        gTmp = Val(Right$(sTmp, Len(sTmp) - InStr(sTmp, "PRODUCT=") - 8))
        sTmp = "Adobe Indesign " & Format$(gTmp, "#0.0")
    End If
    Set currNode = Nothing
    Set xmlDoc = Nothing
    
    Err.Clear
    On Error GoTo 0
   
'   Dim FN As Integer, sTmp As String, lPos As Long
'   FN = FreeFile
'   Open sFName For Binary Access Read Shared As #FN
'      Do While Not EOF(FN)
'         sTmp = GetTextLineFromFile(FN)
'         If sTmp Like "*<???:CreatorTool>*InDesign*</???:CreatorTool>" Then
'            sTmp = Trim$(sTmp)
'            lPos = InStr(1, sTmp, ":CreatorTool>", vbTextCompare) + 13
'            sTmp = Mid$(sTmp, lPos, Len(sTmp))
'            lPos = InStr(1, sTmp, ":CreatorTool>", vbTextCompare) - 6
'            sTmp = Mid$(sTmp, 1, lPos)
'            Exit Do
'         End If
'         sTmp = ""
'      Loop
'   Close #FN
   GetINXVersion = sTmp
End Function

'some things make me to change type of this function from simple BOOLEAN to INTEGER, where INTEGER is a PDF version number (i.e. for PDF-1.4 this function returns 4)
Public Function Ensure_file_is_PDFcompatible(ByVal sFName As String) As Integer
   Dim bHeader(1 To 9) As Byte, FN As Integer
   FN = FreeFile
   Open sFName For Binary As #FN
      Get #FN, , bHeader
   Close #FN
   Ensure_file_is_PDFcompatible = 0
   If bHeader(1) = &H25 And bHeader(2) = &H50 And bHeader(3) = &H44 _
         And bHeader(4) = &H46 And bHeader(5) = &H2D And bHeader(6) = &H31 And bHeader(7) = &H2E Then
      Ensure_file_is_PDFcompatible = CInt(Val(Chr$(bHeader(8)) & Chr$(bHeader(9))))
   End If
End Function

Public Function GetTextLineFromFile(ByVal iFNum As Integer) As String
    Dim bChars As String, bBytik(1 To 1024) As Byte, lCurrPos As Long, i As Long
    If EOF(iFNum) Then GetTextLineFromFile = vbNullString: Exit Function
    bChars = vbNullString
    lCurrPos = Seek(iFNum)
    Get #iFNum, , bBytik
    For i = 1 To 256
        If bBytik(i) = &HA Or bBytik(i) = &HD Then
            Seek iFNum, lCurrPos + i
            Exit For
        End If
        If bBytik(i) >= &H20 Then bChars = bChars & Chr$(bBytik(i))
    Next i
    GetTextLineFromFile = bChars
End Function

Public Function MyGetInStrLine(ByVal lStart As Long, ByRef sSearchWhere As String, _
        ByVal sSearchWhat As String, Optional ByVal lOption As VbCompareMethod _
        = vbBinaryCompare) As String
    Dim lInStrBegin As Long, lInStrEnd As Long, sRes As String
    sRes = vbNullString
    lInStrBegin = InStr(lStart, sSearchWhere, sSearchWhat, lOption)
    If lInStrBegin > 0 Then
        lInStrEnd = InStr(lInStrBegin, sSearchWhere, vbLf, lOption)
        If lInStrEnd = 0 Then lInStrEnd = Len(sSearchWhere)
        sRes = Mid$(sSearchWhere, lInStrBegin, lInStrEnd - lInStrBegin)
    End If
    MyGetInStrLine = sRes
End Function

Public Function ReadEpsFileToStringArray(sFile As String, Optional ByVal lBeginOffset As Long = 1, _
        Optional ByVal lBlockSize As Long = &H200000, _
        Optional ByRef lBlocksToRead As Integer = 3, _
        Optional ByRef frmOWNER As Form = Nothing, _
        Optional ByVal sFormCaption As String, _
        Optional ByVal sFindThatStringInAnyFile As String = vbNullString) As Variant
    Dim bStep() As Byte, sStepString As String, varArray As Variant
    Dim lCurrSeek As Long, i As Long, lFileLength As Long, j As Long, nBytes As Long
    Dim vResultArray() As String, iFileNumber As Integer, boolRes As Boolean
    
    iFileNumber = FreeFile
    'get current file lengh
    lFileLength = FileLen(sFile)
    
 If Not (frmOWNER Is Nothing) Then frmOWNER.Show
 If Not (frmOWNER Is Nothing) Then frmOWNER.Caption = sFormCaption
 If Not (frmOWNER Is Nothing) Then frmOWNER.ProgressBar1.Min = 0
 If Not (frmOWNER Is Nothing) Then frmOWNER.ProgressBar1.Max = lFileLength
    
    'set read block size
    nBytes = lBlockSize
    'if file size less then read block, set read block to file size
    If lFileLength < nBytes Then nBytes = lFileLength
    'here is our buffer
    'ReDim bStep(1 To nBytes)
    sStepString = String$(nBytes, vbLf)
    
    ReDim vResultArray(0 To 0)
    'open file
    Open sFile For Binary Access Read Shared As iFileNumber
        'cyclic reading from file into buffer
        For j = lBeginOffset To lFileLength Step nBytes
            If lBlocksToRead = 0 Then Exit For
            lBlocksToRead = lBlocksToRead - 1
            'we have to remember current file position
            lCurrSeek = Seek(iFileNumber)
            'fill buffer
            Get #iFileNumber, , sStepString 'bStep
            'Get #iFileNumber, , bStep
            
'            'check buffer backward to find last symbol &HA (NEW_LINE)
'            If (Seek(iFileNumber) - 1) < lFileLength Then
'                For i = nBytes To 1 Step -1
'                    If bStep(i) = &HA Then
'                        'if &HA is found, we have to step back file pos to position of &HA
'                        Seek iFileNumber, lCurrSeek + i
'                        Exit For
'                    End If
'                    DoEvents
'                Next i
'            End If
            
            If Not (frmOWNER Is Nothing) Then frmOWNER.ProgressBar1.Value = lCurrSeek
'            'now we are in position of last &HA in buffer,
'            'so we can go on
            
            'convert buffer into unicode string
            'sStepString = StrConv(bStep, vbUnicode)
            
'i = InStrRev(sStepString, Chr$(&HA))
'Seek iFileNumber, lCurrSeek + i - 1
'sStepString = Left$(sStepString, Len(sStepString) - i)

            'remove symbol &HD (CARRIAGE RETURN)
            sStepString = Replace$(sStepString, vbCr, vbLf)
            sStepString = Replace$(sStepString, vbLf & vbLf, vbLf)
            'For i = 1 To nBytes
            '    If AscW(Mid$(sStepString, i, 1)) = &HD Then
            '        Mid$(sStepString, i, 1) = vbLf
            '    End If
            'Next i
            'split string into array of strings by delimiter &HA
            varArray = Split(sStepString, vbLf)
'            vResultArray = MyPrepareByteArray(bStep, frmOWNER)
            'analyzing strings
            For i = LBound(varArray) To UBound(varArray)
                sStepString = varArray(i)
                'we dont need strings less than 6 chars
                If Len(sStepString) > 6 Then
                    'only informative strings
                    If Len(sFindThatStringInAnyFile) > 0 Then
                        boolRes = (InStrRev(UCase$(sStepString), UCase$(sFindThatStringInAnyFile)) > 0)
                    Else
                        boolRes = (Left$(sStepString, 2) = "%%") Or (Left$(sStepString, 4) = "%%AI") _
                            Or (Left$(sStepString, 6) = "%Image") Or (Left$(sStepString, 3) = "%AI")
                    End If
                    If boolRes Then
                        'resize result array by 1
                        ReDim Preserve vResultArray(0 To UBound(vResultArray) + 1)
                        vResultArray(UBound(vResultArray)) = sStepString
                    End If
                End If
                DoEvents
                If frmOWNER.bEPSAnalizing = False Then
                    Close iFileNumber
'                    ReDim vResultArray(0 To 0)
'                    vResultArray(0) = "Analyzing aborted!!!"
                    ReadEpsFileToStringArray = vResultArray
                    Exit Function
                End If
            Next i
            DoEvents
            If frmOWNER.bEPSAnalizing = False Then
                Close iFileNumber
'                ReDim vResultArray(0 To 0)
'                vResultArray(0) = "Analyzing aborted!!!"
                ReadEpsFileToStringArray = vResultArray
                Exit Function
            End If
        Next j
    Close iFileNumber
    
    If Not (frmOWNER Is Nothing) Then frmOWNER.Hide
    
    ReadEpsFileToStringArray = vResultArray
End Function


Public Function MyPrepareByteArray(ByRef bArray() As Byte, ByRef fff As Form) As Variant
    Dim i As Long, LB As Long, UB As Long, sRes() As String, lPrevPos As Long
    Dim bTemp() As Byte, boolHaveSpace As Boolean
    LB = LBound(bArray): UB = UBound(bArray)
    ReDim sRes(1 To 1)
    ReDim bTemp(1 To 1): bTemp(1) = &H25
    lPrevPos = LB
    'sTemp = vbNullString
    boolHaveSpace = False
    For i = LB To UB
        If bArray(i) < &HA Then bArray(i) = &HA
        If bArray(i) = &HD Then bArray(i) = &HA
        If bArray(i) <> &HA Then
            If bTemp(1) = &H25 Then '%
                'sTemp = sTemp & ChrW$(bArray(i))
                bTemp(UBound(bTemp)) = bArray(i)
                ReDim Preserve bTemp(1 To UBound(bTemp) + 1)
                If bArray(i) = &H20 Then boolHaveSpace = True
            End If
        Else
            If i - lPrevPos > 6 Then
                'sRes(UBound(sRes)) = sTemp
                If bTemp(1) = &H25 And boolHaveSpace Then '% and Space
                    ReDim Preserve bTemp(1 To UBound(bTemp) - 1)
                    sRes(UBound(sRes)) = StrConv(bTemp, vbUnicode)
                    ReDim Preserve sRes(1 To UBound(sRes) + 1)
                End If
            End If
            'sTemp = vbNullString
            ReDim bTemp(1 To 1):  bTemp(1) = &H25
            lPrevPos = i + 1
            boolHaveSpace = False
        End If
        DoEvents
'            If fff.bEPSAnalizing = False Then
'                ReDim sRes(1 To 1)
'                sRes(1) = "Analyzing aborted!!!"
'                MyPrepareByteArray = sRes
'                Exit Function
'            End If
    Next i
    If UBound(sRes) > 1 Then ReDim Preserve sRes(1 To UBound(sRes) - 1)
    MyPrepareByteArray = sRes
End Function


Public Function GetCDRVersion(ByVal sFileName As String) As Integer
    Dim AA As String, ZZ As Cls_Zip, bBBB() As Variant, i As Long
    Dim bData() As Boolean, iFNN As Integer, bCDRHeader(1 To 15) As Byte
    On Error Resume Next
    AA = sFileName
    iFNN = FreeFile()
    Open sFileName For Binary Access Read Shared As #iFNN
        Get #iFNN, , bCDRHeader
    Close #iFNN
    GetCDRVersion = 0
    If bCDRHeader(1) = &H50 And bCDRHeader(2) = &H4B Then ' here we have CDR version 14 or more that is packed by ZIP algorithm
        Set ZZ = New Cls_Zip
        ZZ.Get_Contents AA
        If ZZ.FileCount > 0 Then
            ReDim bData(1 To ZZ.FileCount)
            For i = 1 To ZZ.FileCount
                If UCase$(ZZ.FileName(i)) = "CONTENT/ROOT.DAT" Or UCase$(ZZ.FileName(i)) = "CONTENT/RIFFDATA.CDR" Then
                    bData(i) = True
                Else
                    bData(i) = False
                End If
            Next i
        End If
        ReDim bBBB(1 To 1)
        ZZ.UnPackToBuffer bData, bBBB, 15
        GetCDRVersion = bBBB(UBound(bBBB))(11) - &H37
        Set ZZ = Nothing
    End If
    If (bCDRHeader(1) = &H52 And bCDRHeader(2) = &H49 And bCDRHeader(3) = &H46 And bCDRHeader(4) = &H46 _
                And bCDRHeader(9) = &H43 And bCDRHeader(10) = &H44 And bCDRHeader(11) = &H52) _
            Or _
                (bCDRHeader(1) = &H52 And bCDRHeader(2) = &H49 And bCDRHeader(3) = &H46 And bCDRHeader(4) = &H46 _
                And bCDRHeader(9) = &H63 And bCDRHeader(10) = &H64 And bCDRHeader(11) = &H72) _
                Then
        'sCDRHeader = bCDRHeader(12)
        If Not (bCDRHeader(12) >= &H31 And bCDRHeader(12) <= &H39) Then
            GetCDRVersion = bCDRHeader(12) - &H37
        Else
            GetCDRVersion = bCDRHeader(12) - &H30
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function


Public Function GetAIVersion(ByVal sFileName As String, Optional ByRef frmOWNER As Form) As String
    Dim AA As String, bBBB As Variant, i As Long, lTMP As Long, j As Long, lRecLen As Long, lStreamLen As Long
    Dim bData() As Boolean, iFNN As Integer, bPDFHeader() As Byte, sPDFHeader As String, lCurrSeek As Long, lXrefSeek As Long
    'On Error Resume Next
    AA = sFileName
    'MsgBox AA
    iFNN = FreeFile()
    ReDim bPDFHeader(1 To 48)
    Open sFileName For Binary Access Read Shared As #iFNN
        Get #iFNN, , bPDFHeader
    Close #iFNN
    sPDFHeader = StrConv(bPDFHeader, vbUnicode)
    'MsgBox sPDFHeader
    GetAIVersion = vbNullString
    '
    'If bPDFHeader(1) = &H25 And bPDFHeader(2) = &H50 And bPDFHeader(3) = &H44 And bPDFHeader(4) = &H46 _
            And bPDFHeader(5) = &H2D And bPDFHeader(6) = &H31 And bPDFHeader(7) = &H2E Then
    If Left$(sPDFHeader, 7) <> "%PDF-1." Then   'AI header is NOT PDF-compatible - we have versions 8 or belove
        'ReDim bBBB(1 To 1)
        sPDFHeader = GetEPSInfo(sFileName, frmOWNER).EPSCreator
        'lTMP = InStrRev(UCase$(sPDFHeader), " ")
        'sPDFHeader = Mid$(sPDFHeader, lTMP + 1)
        'If sPDFHeader = "X" Then
        '    GetAIVersion = "Adobe Illustrator(R) 10.0"
        'Else
        '    GetAIVersion = "Adobe Illustrator(R) " & sPDFHeader
        'End If
        GetAIVersion = Trim$(sPDFHeader)
        Exit Function
        Err.Clear
        On Error GoTo 0
    Else 'Here we have AI version 9 or above. Must parse PDF :(((
        'at first we read END of file and search for "XREF" array begin
        iFNN = FreeFile()
        Open sFileName For Binary Access Read Shared As #iFNN
            Seek #iFNN, FileLen(sFileName) - (UBound(bPDFHeader) - LBound(bPDFHeader) + 1)
            Get #iFNN, , bPDFHeader
            sPDFHeader = StrConv(bPDFHeader, vbUnicode)
            lTMP = InStr(UCase$(sPDFHeader), "STARTXREF")
            If lTMP > 0 Then ' get xref offset here
                sPDFHeader = UCase$(sPDFHeader)
                sPDFHeader = Replace$(sPDFHeader, vbCr, vbLf)
                sPDFHeader = Replace$(sPDFHeader, vbLf & vbLf, vbLf)
                bBBB = Split(sPDFHeader, vbLf)
                For i = LBound(bBBB) To UBound(bBBB)
                    If bBBB(i) = "STARTXREF" Then
                        lTMP = Val(bBBB(i + 1))
                        Exit For
                    End If
                Next i
            End If
            If lTMP > 0 Then 'here we jump to XREF table and check objects offsets for correct object
                Seek #iFNN, lTMP + 1
                Get #iFNN, , bPDFHeader
                sPDFHeader = StrConv(bPDFHeader, vbUnicode)
                lTMP = InStr(6, sPDFHeader, " ")
                sPDFHeader = Mid$(sPDFHeader, lTMP + 1)
                For i = 1 To Len(sPDFHeader)
                    If Not IsNumeric(Mid$(sPDFHeader, i, 1)) Then
                        If Mid$(sPDFHeader, i, 2) = vbCrLf Then
                            lRecLen = 2 ' RecLen is used for correcting objects record length
                        Else
                            lRecLen = 1
                        End If
                        Exit For
                    End If
                Next i
                sPDFHeader = Replace$(sPDFHeader, vbCr, vbLf)
                sPDFHeader = Replace$(sPDFHeader, vbLf & vbLf, vbLf)
                bBBB = Split(sPDFHeader, vbLf)
                j = Val(bBBB(LBound(bBBB)))
                'j = Val(sPDFHeader) ' count of objects in PDF-like document
                lTMP = Seek(iFNN) - 49 + lTMP + i + lRecLen
                Seek #iFNN, lTMP
                ReDim bPDFHeader(1 To 20)
                lRecLen = 2
                For i = 1 To j
                    
                    Get #iFNN, , bPDFHeader
                    lXrefSeek = Seek(iFNN)
                    sPDFHeader = StrConv(bPDFHeader, vbUnicode)
                    sPDFHeader = Left$(sPDFHeader, 10)
                    If sPDFHeader <> "0000000000" Then ' go thru the objects and search EPS header
                        lTMP = Val(sPDFHeader)
                        Seek #iFNN, lTMP + 1
                        lCurrSeek = Seek(iFNN)
                        sPDFHeader = Space$(255)
                        Get #iFNN, , sPDFHeader
                        lTMP = InStr(UCase$(sPDFHeader), "ENDOBJ")
                        If lTMP > 0 Then sPDFHeader = Left$(sPDFHeader, lTMP - 1)
                        lTMP = InStr(UCase$(sPDFHeader), "LENGTH ")
                        If lTMP > 0 Then
                            lStreamLen = Val(Mid$(sPDFHeader, lTMP + 6))
                            If lStreamLen > 0 Then
                                'lCurrSeek = lTmp
                                lTMP = InStr(lTMP + 6, UCase$(sPDFHeader), ">>STREAM")
                                If lTMP > 0 Then
                                    lCurrSeek = lCurrSeek + lTMP + lRecLen + 7
                                    sPDFHeader = Mid$(sPDFHeader, lTMP + 8 + lRecLen)
                                    If Len(sPDFHeader) > 30 Then
                                        If Left$(sPDFHeader, 14) = "%!PS-Adobe-3.0" Then
                                            Seek #iFNN, lCurrSeek
                                            sPDFHeader = Space$(255)
                                            Get #iFNN, , sPDFHeader
                                            lTMP = InStr(UCase$(sPDFHeader), "%%CREATOR: ")
                                            If lTMP > 0 Then
                                                sPDFHeader = Mid$(sPDFHeader, lTMP + 11)
                                                lTMP = InStr(UCase$(sPDFHeader), "%%")
                                                sPDFHeader = Left$(sPDFHeader, lTMP - lRecLen)
                                                If Right$(sPDFHeader, 1) = vbCr Then sPDFHeader = Left$(sPDFHeader, Len(sPDFHeader) - 1)
                                                'lTMP = InStrRev(UCase$(sPDFHeader), " ")
                                                'sPDFHeader = Mid$(sPDFHeader, lTMP + 1)
                                                'If sPDFHeader = "X" Then
                                                '    GetAIVersion = "Adobe Illustrator(R) 10.0"
                                                'Else
                                                '    GetAIVersion = "Adobe Illustrator(R) " & sPDFHeader
                                                'End If
                                                GetAIVersion = Trim$(sPDFHeader)
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    Seek #iFNN, lXrefSeek
                Next i
            End If
        Close #iFNN
    End If
    Err.Clear
    On Error GoTo 0
End Function



