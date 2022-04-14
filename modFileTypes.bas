Attribute VB_Name = "modFileTypes"
Option Explicit

Private Type bForBytes
    bArr(1 To 4) As Byte
End Type

Private Type bTwoBytes
    bArr(1 To 2) As Byte
End Type

Public Enum eTIFBytesOrder
    tifPC
    tifMAC
    tifNonTIFImage
End Enum

Public Enum eBytesPerChannel
    bpc1 = 1
    bpc2 = 2
    bpc4 = 4
    bpc8 = 8
    bpc16 = 16
End Enum

Public Enum ePhotoshopColorMode
    pcmGrayscale = 1
    pcmLAB = 2
    pcmRGB = 3
    pcmCMYK = 4
    pcmMultichannel = 5
End Enum

Public Enum eTIFFColorMode
    tifBW = 0
    tifGrayscale = 1
    tifRGB = 2
    tifRGBPalette = 3
    tifTransparencyMask = 4
    tifCMYK = 5
    tifYCbCr = 6
    tifLAB = 8
End Enum

Public Enum eTIFFCompression
    tifUncompressed = 1
    tifCCIT_1D = 2
    tifCCIT_Group3Fax = 3
    tifCCIT_Group4Fax = 4
    tifLZW = 5
    tifJPEG = 6
    tifJPEG_Photoshop = 7
    tifZIP = 8
    tifPackBits = 9 '32773
End Enum

Public Enum eTIFFUnits
    tifUnspecified = 1
    tifINCH = 2
    tifCM = 3
End Enum

Public Enum eEPSType
    epsVector
    epsDCS1
    epsDCS2
    epsPhotoshop
    epsNonEPSImage
End Enum

Public Type ftEPSHeader
    bFileHeader As bForBytes
    bOffset As bForBytes
End Type

Public Type ftTIFFHeader
    bTIFFMarker As bForBytes
    bFakeData1(5 To 50) As Byte
    bChannels1 As bForBytes
    bFakeData2(55 To 66) As Byte
    bCompression1 As bTwoBytes
    bFakeData3(69 To 78) As Byte
    bColorMode As bTwoBytes
    bFakeData4(81 To 86) As Byte
    bCompression2 As bForBytes
    bFakeData5(91 To 102) As Byte
    bChannels2 As bTwoBytes
End Type

Public Type ftTIFFInfo
    TIFFBytesOrder As eTIFBytesOrder
    TIFFColorMode As eTIFFColorMode
    TIFFCompression As eTIFFCompression
    TIFFChannels As Long
    TIFFAlfaChannels As Long
    TIFFWidth As Long
    TIFFHeight As Long
    TIFFXRes As Double
    TIFFYRes As Double
    TIFFBitsPerSample As Long
    TIFFUnits As eTIFFUnits
    TIFFProfile As String
    TIFFProfileData() As Byte
End Type

Public Type ftTIFFTag
    tag1word_ID As Long
    tag2word_Datatype As Long
    tag3long_NumValues As Long
    tag4long_Data As Long
End Type

Public Type ftEPSInfo
    EPSCreator As String
    EPS_AI8_Creator_found As Boolean
    EPSDocProcessColors As String
    EPSDocCustomColors As String
    EPS_DCSPlates() As String
    EPSType As eEPSType
    EPS_BPS As eBytesPerChannel
    EPS_PhotoshopMode As ePhotoshopColorMode
End Type

Public Type tFindResult
   bNoErrors As Boolean
   lResultPosition As Long
   sResultString As String
End Type

Public sEPSFormat(0 To 3) As String, sPhotoshopColorMode(0 To 5) As String
Public sTIFFormat(0 To 1) As String, sTIFFColorMode(0 To 8) As String, sTIFFCompression(0 To 9) As String
Public sTIFFUnits(0 To 3) As String

Public Function GetFullLong(ByRef arrOff As bForBytes, Optional ByVal bMACBytesOrder As Boolean = False) As Currency
    If bMACBytesOrder Then
        GetFullLong = CCur(arrOff.bArr(1)) * &H1000000 + CCur(arrOff.bArr(2)) * &H10000 + CCur(arrOff.bArr(3)) * &H100 + CCur(arrOff.bArr(4))
    Else
        GetFullLong = CCur(arrOff.bArr(4)) * &H1000000 + CCur(arrOff.bArr(3)) * &H10000 + CCur(arrOff.bArr(2)) * &H100 + CCur(arrOff.bArr(1))
    End If
End Function

Public Function GetShortLong(ByRef arrOff As bTwoBytes, Optional ByVal bMACBytesOrder As Boolean = False) As Long
    If bMACBytesOrder Then
        GetShortLong = CLng(arrOff.bArr(1)) * &H100 + CLng(arrOff.bArr(2))
    Else
        GetShortLong = CLng(arrOff.bArr(2)) * &H100 + CLng(arrOff.bArr(1))
    End If
End Function

Public Function GetEPSInfo(ByVal sFileName As String, Optional ByVal frmOWNER As Form, Optional ByVal bAnalizeWholeFile As Boolean = False) As ftEPSInfo
    Dim iFN As Integer, tEPSHead As ftEPSHeader, FS As Object
    Dim lOffs As Currency, bHEAD As bForBytes, tEPSi As ftEPSInfo, sLine As Variant
    Dim vArr As Variant, lSeek As Long, lPos As Long
    Dim bOKCreator As Boolean, bOKDocProcColors As Boolean, bOKDocCustColors As Boolean
    Dim bOKDCS1Plates As Boolean, bOKDCS2Plates As Boolean, bOKImageData As Boolean
    Dim j As Long, bOK_AI8_Creator As Boolean
    
    iFN = FreeFile
    Set FS = CreateObject("Scripting.FileSystemObject")
    If Not FS.FileExists(sFileName) Then Exit Function
    If FileLen(sFileName) < 50 Then Exit Function
    bHEAD.bArr(1) = &HC5
    bHEAD.bArr(2) = &HD0
    bHEAD.bArr(3) = &HD3
    bHEAD.bArr(4) = &HC6
    Open sFileName For Binary Access Read Shared As iFN
        Get #iFN, , tEPSHead
    Close iFN
    If tEPSHead.bFileHeader.bArr(1) = &H25 Or _
        tEPSHead.bFileHeader.bArr(2) = &H21 Or _
        tEPSHead.bFileHeader.bArr(3) = &H50 Or _
        tEPSHead.bFileHeader.bArr(4) = &H53 Then
        lOffs = 1
    ElseIf tEPSHead.bFileHeader.bArr(1) <> bHEAD.bArr(1) Or _
        tEPSHead.bFileHeader.bArr(2) <> bHEAD.bArr(2) Or _
        tEPSHead.bFileHeader.bArr(3) <> bHEAD.bArr(3) Or _
        tEPSHead.bFileHeader.bArr(4) <> bHEAD.bArr(4) Then
      GetEPSInfo.EPSType = epsNonEPSImage: Exit Function
    Else
        lOffs = GetFullLong(tEPSHead.bOffset) + 1
    End If
    
    ReDim tEPSi.EPS_DCSPlates(1 To 1)

    vEPS_DATA = Array()
    vEPS_DATA = ReadEpsFileToStringArray(sFileName, lOffs, , IIf(bAnalizeWholeFile, 100, 5), frmOWNER, "Analyzing file, please wait ...")
'    Open sFileName For Binary Access Read As iFN
'        Seek iFN, lOffs
 
 'If Not (frmOWNER Is Nothing) Then frmOWNER.Show
 'If Not (frmOWNER Is Nothing) Then frmOWNER.Caption = "Analyzing EPS file, please wait ..."
 'If Not (frmOWNER Is Nothing) Then frmOWNER.ProgressBar1.Min = 0
 'If Not (frmOWNER Is Nothing) Then frmOWNER.ProgressBar1.Max = FileLen(sFileName)
        
'        Do While Not EOF(iFN)
        
'    If Not (frmOWNER Is Nothing) Then frmOWNER.ProgressBar1.Value = Seek(iFN)
For j = LBound(vEPS_DATA) To UBound(vEPS_DATA)
            'sLine = GetTextLineFromFile(iFN)
            sLine = vEPS_DATA(j)
            If Not bOKCreator Then
                 If InStr(sLine, "%%Creator:") > 0 Then
                    bOKCreator = True
                    sLine = Trim$(Replace(sLine, "%%Creator:", ""))
                    tEPSi.EPSCreator = sLine
                    If InStr(UCase$(sLine), "PHOTOSHOP") > 0 Then
                        tEPSi.EPSType = epsPhotoshop
                    Else
                        tEPSi.EPSType = epsVector
                    End If
                    tEPSi.EPS_AI8_Creator_found = False
                End If
            End If
            
            If Not bOK_AI8_Creator Then
                 If InStr(sLine, "%%AI8_CreatorVersion:") > 0 Then
                    bOK_AI8_Creator = True
                    bOKCreator = True
                    sLine = Trim$(Replace(sLine, "%%AI8_CreatorVersion:", "Adobe Illustrator(R)"))
                    tEPSi.EPSCreator = sLine
                    tEPSi.EPSType = epsVector
                    tEPSi.EPS_AI8_Creator_found = True
                End If
            End If
            
            If Not bOKDocProcColors Then
                'sLine = MyGetInStrLine(1, sOneLine, "%%DocumentProcessColors:")
                If InStr(sLine, "%%DocumentProcessColors:") > 0 Then
                    bOKDocProcColors = True
                    sLine = Trim$(Replace(sLine, "%%DocumentProcessColors:", ""))
                    sLine = Replace(sLine, "Cyan Magenta Yellow Black", "CMYK")
                    tEPSi.EPSDocProcessColors = sLine
                End If
            End If
            
            If Not bOKDocCustColors Then
                'sLine = MyGetInStrLine(1, sOneLine, "%%DocumentCustomColors:")
                If InStr(sLine, "%%DocumentCustomColors:") > 0 Then
                    bOKDocCustColors = True
'                    lSeek = Seek(iFN)
                    lSeek = j
                    sLine = Trim$(Replace(sLine, "%%DocumentCustomColors:", ""))
                    sLine = Replace(sLine, ")", "")
                    sLine = Replace(sLine, "(", "")
                    If Len(tEPSi.EPSDocCustomColors) > 0 Then
                        tEPSi.EPSDocCustomColors = tEPSi.EPSDocCustomColors & " + " & Trim$(sLine)
                    Else
                        tEPSi.EPSDocCustomColors = Trim$(sLine)
                    End If
Rep1:
'                    sLine = GetTextLineFromFile(iFN)
                    j = j + 1
                    sLine = vEPS_DATA(j)
                    If InStr(sLine, "%%+") > 0 Then
'                        lSeek = Seek(iFN)
                        lSeek = j
                        sLine = Trim$(Replace(sLine, "%%+", ""))
                        sLine = Replace(sLine, ")", "")
                        sLine = Replace(sLine, "(", "")
                        tEPSi.EPSDocCustomColors = tEPSi.EPSDocCustomColors & " + " & Trim$(sLine)
                        GoTo Rep1
                    Else
                        j = lSeek
                    End If
                End If
            End If
            
            
            If (tEPSi.EPSType = epsVector) And bOKCreator And bOKDocProcColors And bOKDocCustColors Then Exit For
            '%%BlackPlate:
            If (Not bOKDCS1Plates) And tEPSi.EPSType <> epsVector Then
                If InStr(sLine, "Plate: ") > 2 And InStr(sLine, "%%") = 1 Then
                    bOKDCS1Plates = True
                    tEPSi.EPSType = epsDCS1
                    'Exit Do
                End If
            End If
            
            '%%PlateFile:
            If (Not bOKDCS2Plates) And (Not bOKDCS1Plates) And tEPSi.EPSType <> epsVector Then
                If InStr(sLine, "%%PlateFile:") > 0 Then
                    bOKDCS2Plates = True
                    tEPSi.EPSType = epsDCS2
                    sLine = Trim$(Replace(sLine, "%%PlateFile:", ""))
                    tEPSi.EPS_DCSPlates(UBound(tEPSi.EPS_DCSPlates)) = sLine
                    ReDim Preserve tEPSi.EPS_DCSPlates(1 To UBound(tEPSi.EPS_DCSPlates) + 1)
                    'Exit Do
                End If
            End If
            
            '%ImageData:
            If (Not bOKImageData) And tEPSi.EPSType <> epsVector Then
                If InStr(sLine, "%ImageData:") > 0 Then
                    bOKImageData = True
                    sLine = Trim$(Replace(sLine, "%ImageData:", ""))
                    If InStr(sLine, " ") > 1 Then
    '                    vArr = Nothing
                        vArr = Split(sLine, " ")
                        If IsArray(vArr) Then
                            If (UBound(vArr) - LBound(vArr)) > 6 Then
                                tEPSi.EPS_BPS = vArr(LBound(vArr) + 2)
                                tEPSi.EPS_PhotoshopMode = vArr(LBound(vArr) + 3)
                            End If
                        End If
                    End If
                    Exit For
                End If
            End If

            If (InStr(sLine, "%%BeginBinary:") > 0) Or (InStr(sLine, "%%BeginData:") > 0) Then Exit For
            If InStr(sLine, "%%EOF") > 0 Then Exit For
            DoEvents
'        Loop
'    Close iFN
Next j

'If Not (frmOWNER Is Nothing) Then frmOWNER.Hide
    If UBound(tEPSi.EPS_DCSPlates) > 1 Then ReDim Preserve tEPSi.EPS_DCSPlates(1 To UBound(tEPSi.EPS_DCSPlates) - 1)
    GetEPSInfo = tEPSi
End Function

Public Function GetTIFInfo_OLD(ByVal sFileName As String) As ftTIFFInfo
    Dim iFN As Integer, tTIFFHead As ftTIFFHeader, FS As Object
    Dim bHEAD_PC As bForBytes, bHEAD_MAC As bForBytes, sLine As String
    Dim tifRes As ftTIFFInfo
    iFN = FreeFile
    Set FS = CreateObject("Scripting.FileSystemObject")
    If Not FS.FileExists(sFileName) Then Exit Function
    If FileLen(sFileName) < 200 Then Exit Function
    bHEAD_PC.bArr(1) = &H49
    bHEAD_PC.bArr(2) = &H49
    bHEAD_PC.bArr(3) = &H2A
    bHEAD_PC.bArr(4) = &H0
    
    bHEAD_MAC.bArr(1) = &H4D
    bHEAD_MAC.bArr(2) = &H4D
    bHEAD_MAC.bArr(3) = &H0
    bHEAD_MAC.bArr(4) = &H2A
    
    Open sFileName For Binary Access Read Shared As iFN
        Get #iFN, , tTIFFHead
    Close #iFN
    With tTIFFHead
        If .bTIFFMarker.bArr(1) = bHEAD_PC.bArr(1) And _
                .bTIFFMarker.bArr(2) = bHEAD_PC.bArr(2) And _
                .bTIFFMarker.bArr(3) = bHEAD_PC.bArr(3) And _
                .bTIFFMarker.bArr(4) = bHEAD_PC.bArr(4) Then
            tifRes.TIFFBytesOrder = tifPC
        ElseIf .bTIFFMarker.bArr(1) = bHEAD_MAC.bArr(1) And _
                .bTIFFMarker.bArr(2) = bHEAD_MAC.bArr(2) And _
                .bTIFFMarker.bArr(3) = bHEAD_MAC.bArr(3) And _
                .bTIFFMarker.bArr(4) = bHEAD_MAC.bArr(4) Then
            tifRes.TIFFBytesOrder = tifMAC
        Else
            tifRes.TIFFBytesOrder = tifNonTIFImage
        End If
        If tifRes.TIFFBytesOrder <> tifNonTIFImage Then
            tifRes.TIFFColorMode = GetShortLong(.bColorMode, -(tifRes.TIFFBytesOrder))
'            If GetShortLong(.bCompression1, -(tifRes.TIFFBytesOrder)) = 1 And _
                    GetFullLong(.bCompression2, -(tifRes.TIFFBytesOrder)) = 1 Then
'                tifRes.TIFFNoCompressed = True
'            Else
'                tifRes.TIFFNoCompressed = False
'            End If
            tifRes.TIFFChannels = CLng(GetFullLong(.bChannels1, -(tifRes.TIFFBytesOrder)))
            Select Case tifRes.TIFFColorMode
                Case 0, 1
                    tifRes.TIFFAlfaChannels = tifRes.TIFFChannels - 1
                Case 2, 8
                    tifRes.TIFFAlfaChannels = tifRes.TIFFChannels - 3
                Case 5
                    tifRes.TIFFAlfaChannels = tifRes.TIFFChannels - 4
                Case Else
                    tifRes.TIFFAlfaChannels = tifRes.TIFFChannels
            End Select
        End If
    End With
    'GetTIFInfo = tifRes
End Function

Public Function GetTIFInfo(ByVal sFileName As String) As ftTIFFInfo
    Dim iFN As Integer, tTIFFHead As ftTIFFHeader, FS As Object
    Dim bHEAD_PC As bForBytes, bHEAD_MAC As bForBytes, sLine As String
    Dim tifRes As ftTIFFInfo, lTMP(1 To 4) As Long
    
    Dim arTag(0 To 11) As Byte, tTAG As ftTIFFTag, lFirst_IFD_offset As Currency
    Dim bbQuaziLong As bForBytes, bbQuaziWord As bTwoBytes, iNumOfEntries As Integer
    Dim ii As Integer, lOffs1 As Long
    Dim lRatioFirst As Currency, lRatioSecond As Currency, sRationalFull As String
    Dim bProfile() As Byte, sProfile As String, lPoss As Long
    
    iFN = FreeFile
    Set FS = CreateObject("Scripting.FileSystemObject")
    If Not FS.FileExists(sFileName) Then Exit Function
    If FileLen(sFileName) < 200 Then Exit Function
    bHEAD_PC.bArr(1) = &H49
    bHEAD_PC.bArr(2) = &H49
    bHEAD_PC.bArr(3) = &H2A
    bHEAD_PC.bArr(4) = &H0
    
    bHEAD_MAC.bArr(1) = &H4D
    bHEAD_MAC.bArr(2) = &H4D
    bHEAD_MAC.bArr(3) = &H0
    bHEAD_MAC.bArr(4) = &H2A
    
    Open sFileName For Binary Access Read Shared As iFN
        Get #iFN, , tTIFFHead
    
    With tTIFFHead
        If .bTIFFMarker.bArr(1) = bHEAD_PC.bArr(1) And _
                .bTIFFMarker.bArr(2) = bHEAD_PC.bArr(2) And _
                .bTIFFMarker.bArr(3) = bHEAD_PC.bArr(3) And _
                .bTIFFMarker.bArr(4) = bHEAD_PC.bArr(4) Then
            tifRes.TIFFBytesOrder = tifPC
        ElseIf .bTIFFMarker.bArr(1) = bHEAD_MAC.bArr(1) And _
                .bTIFFMarker.bArr(2) = bHEAD_MAC.bArr(2) And _
                .bTIFFMarker.bArr(3) = bHEAD_MAC.bArr(3) And _
                .bTIFFMarker.bArr(4) = bHEAD_MAC.bArr(4) Then
            tifRes.TIFFBytesOrder = tifMAC
        Else
            tifRes.TIFFBytesOrder = tifNonTIFImage
        End If
    End With
    
        If tifRes.TIFFBytesOrder <> tifNonTIFImage Then
            Seek #iFN, 5
            Get #iFN, , bbQuaziLong
            lFirst_IFD_offset = GetFullLong(bbQuaziLong, -(tifRes.TIFFBytesOrder))
            
            Seek #iFN, lFirst_IFD_offset + 1
            Get #iFN, , bbQuaziWord
            iNumOfEntries = GetShortLong(bbQuaziWord, -(tifRes.TIFFBytesOrder))
            ReDim arIFD_Entries(1 To iNumOfEntries)
            
            For ii = 1 To iNumOfEntries
                Get #iFN, , arTag
                bbQuaziWord.bArr(1) = arTag(0)
                bbQuaziWord.bArr(2) = arTag(1)
                tTAG.tag1word_ID = GetShortLong(bbQuaziWord, -(tifRes.TIFFBytesOrder))
                 'If tTAG.tag1word_ID > 300 Then Exit For
                 If tTAG.tag1word_ID <> 256 And tTAG.tag1word_ID <> 257 And tTAG.tag1word_ID <> 258 And _
                    tTAG.tag1word_ID <> 259 And tTAG.tag1word_ID <> 262 And tTAG.tag1word_ID <> 277 And _
                    tTAG.tag1word_ID <> 282 And tTAG.tag1word_ID <> 283 And tTAG.tag1word_ID <> 296 _
                    And tTAG.tag1word_ID <> 34675 _
                    Then GoTo NEXXXT            '34675 - color profile
                 
                bbQuaziWord.bArr(1) = arTag(2)
                bbQuaziWord.bArr(2) = arTag(3)
                tTAG.tag2word_Datatype = GetShortLong(bbQuaziWord, -(tifRes.TIFFBytesOrder))
                bbQuaziLong.bArr(1) = arTag(4)
                bbQuaziLong.bArr(2) = arTag(5)
                bbQuaziLong.bArr(3) = arTag(6)
                bbQuaziLong.bArr(4) = arTag(7)
                tTAG.tag3long_NumValues = CLng(GetFullLong(bbQuaziLong, -(tifRes.TIFFBytesOrder)))
                
                bbQuaziLong.bArr(1) = arTag(8)
                bbQuaziLong.bArr(2) = arTag(9)
                bbQuaziLong.bArr(3) = arTag(10)
                bbQuaziLong.bArr(4) = arTag(11)
                
                If tifRes.TIFFBytesOrder = tifMAC Then
                    If Not (tTAG.tag2word_Datatype = 5 Or tTAG.tag3long_NumValues > 1) Then
                        bbQuaziLong.bArr(1) = arTag(10)
                        bbQuaziLong.bArr(2) = arTag(11)
                        bbQuaziLong.bArr(3) = arTag(8)
                        bbQuaziLong.bArr(4) = arTag(9)
                    End If
                End If
                tTAG.tag4long_Data = CLng(GetFullLong(bbQuaziLong, -(tifRes.TIFFBytesOrder)))

                
                Select Case tTAG.tag1word_ID
                Case 256 '0x0100 ImageWidth
                    tifRes.TIFFWidth = tTAG.tag4long_Data
                Case 257 '0x0101 ImageHeight
                    tifRes.TIFFHeight = tTAG.tag4long_Data
                Case 258 '0x0102 BitsPerSample
                    lOffs1 = Seek(iFN)
                    'If tifRes.TIFFBytesOrder = tifMAC Then
                    If tTAG.tag3long_NumValues = 1 Then
                    '    Seek #iFN, tTAG.tag4long_Data + 2
                        tifRes.TIFFBitsPerSample = tTAG.tag4long_Data
                    Else
                        Seek #iFN, tTAG.tag4long_Data
                    'End If
                        Get #iFN, , bbQuaziWord
                        Seek #iFN, lOffs1
                        tifRes.TIFFBitsPerSample = bbQuaziWord.bArr(1) + bbQuaziWord.bArr(2)
                    End If
                Case 259 '0x0103 Compression
                    If tTAG.tag4long_Data > 8 Then
                        tifRes.TIFFCompression = tifPackBits
                    Else
                        tifRes.TIFFCompression = tTAG.tag4long_Data
                    End If
                Case 262 '0x0106 ColorModel
                    tifRes.TIFFColorMode = tTAG.tag4long_Data
                Case 277 '0x0115 SamplesPerPixel (channels)
                    tifRes.TIFFChannels = tTAG.tag4long_Data
                Case 282 '0x011A XResolution (offset)
                    lOffs1 = Seek(iFN)
                    If tifRes.TIFFBytesOrder = tifMAC Then
                        Seek #iFN, tTAG.tag4long_Data + 1
                    Else
                        Seek #iFN, tTAG.tag4long_Data
                    End If
                    Get #iFN, , bbQuaziLong
                    lRatioFirst = GetFullLong(bbQuaziLong, -(tifRes.TIFFBytesOrder))
                    Get #iFN, , bbQuaziLong
                    lRatioSecond = GetFullLong(bbQuaziLong, -(tifRes.TIFFBytesOrder))
                    tifRes.TIFFXRes = CDbl(lRatioFirst) / CDbl(lRatioSecond)
                    Seek #iFN, lOffs1
                Case 283 '0x011B YResolution (offset)
                    lOffs1 = Seek(iFN)
                    If tifRes.TIFFBytesOrder = tifMAC Then
                        Seek #iFN, tTAG.tag4long_Data + 1
                    Else
                        Seek #iFN, tTAG.tag4long_Data
                    End If
                    Get #iFN, , bbQuaziLong
                    lRatioFirst = GetFullLong(bbQuaziLong, -(tifRes.TIFFBytesOrder))
                    Get #iFN, , bbQuaziLong
                    lRatioSecond = GetFullLong(bbQuaziLong, -(tifRes.TIFFBytesOrder))
                    tifRes.TIFFYRes = CDbl(lRatioFirst) / CDbl(lRatioSecond)
                    Seek #iFN, lOffs1
                Case 296 ' Res UNIT  - 1 -no, 2 - inch, 3 - cm
                    tifRes.TIFFUnits = tTAG.tag4long_Data
                    'Exit For
                Case 34675 ' color profile
                    Seek #iFN, tTAG.tag4long_Data + 1
                    ReDim bProfile(1 To tTAG.tag3long_NumValues)
                    Get #iFN, , bProfile
                    sProfile = StrConv(bProfile, vbUnicode)
                    lPoss = InStr(1, sProfile, "desc" & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0), vbBinaryCompare)
                    sProfile = Mid$(sProfile, lPoss + 12, bProfile(lPoss + 11) - 1)
                    If Len(sProfile) > 0 Then
                            tifRes.TIFFProfile = sProfile
                            ReDim tifRes.TIFFProfileData(1 To tTAG.tag3long_NumValues)
                            tifRes.TIFFProfileData = bProfile
                    End If
                'Case Is > 300
                    'Exit For
                End Select
NEXXXT:
            Next ii
            
        End If
        
Close #iFN

           Select Case tifRes.TIFFColorMode
                Case 0, 1
                    tifRes.TIFFAlfaChannels = tifRes.TIFFChannels - 1
                Case 2, 3, 6, 8
                    tifRes.TIFFAlfaChannels = tifRes.TIFFChannels - 3
                Case 5
                    tifRes.TIFFAlfaChannels = tifRes.TIFFChannels - 4
                Case Else
                    tifRes.TIFFAlfaChannels = tifRes.TIFFChannels
            End Select
    
    GetTIFInfo = tifRes
End Function

Public Function GetGflaxImagePreview(ByVal sFile As String) As Boolean
    On Error Resume Next
    'Dim GFLX As GflAx.GflAx
    'Set GFLX = New GflAx.GflAx
    
    Dim GFLX As Object
    Set GFLX = CreateObject("GflAx.GflAx")
    GFLX.EnableLZW = True
    
    If sFile Like "*.tif" Then
      Dim aTIFdata As ftTIFFInfo, iW As Integer, iH As Integer
      aTIFdata = GetTIFInfo(sFile)
      If aTIFdata.TIFFWidth > 0 And aTIFdata.TIFFHeight > 0 Then
         If aTIFdata.TIFFWidth > aTIFdata.TIFFHeight Then
            iW = 255
            iH = 255 * aTIFdata.TIFFHeight / aTIFdata.TIFFWidth
            If iH > 255 Then iH = 255
         Else
            iH = 255
            iW = 255 * aTIFdata.TIFFWidth / aTIFdata.TIFFHeight
            If iW > 255 Then iW = 255
         End If
      End If
      GFLX.LoadThumbnail sFile, iW, iH
    Else
      GFLX.LoadThumbnail sFile, 255, 255
    End If
    frmPreview.Hide
    If Err.Number = 0 Then
      GFLX.ChangeColorDepth &H100
      GFLX.SaveFormat = 19 'Const AX_ICO = 19 (&H13)
      GFLX.SaveBitmap App.Path & "\~preview.ico"
      Set frmPreview.Image1.Picture = Nothing
      Set frmPreview.Image1.Picture = VB.LoadPicture(App.Path & "\~preview.ico")
    Else
      Err.Clear
    End If
    frmPreview.Move 1000, 1000, frmPreview.Image1.Width, frmPreview.Image1.Height
    'frmPreview.Show
    If Err.Number = 0 Then GetGflaxImagePreview = True Else GetGflaxImagePreview = False
    Set GFLX = Nothing
    Err.Clear
    On Error GoTo 0
End Function

Private Function FindStringInFile(ByVal sFileName As String, ByRef bByteArray() As Byte, _
            Optional ByVal lStartPosition As Long = 1) As tFindResult
            
   Const BLOCK_SIZE As Long = 1048576
   Dim FSO As Object, iFN As Integer, lFLen As Long, lNumBlocks As Long, lLastBlockSize As Long
   Dim i As Long, j As Long, k As Long, bReadBuff(1 To BLOCK_SIZE) As Byte, lFindLen As Long
   Dim lLowBoundOfByteArray As Long, lUpBoundOfByteArray As Long, boCompRes As Boolean
   Dim lPosOfBegin As Long, lPosOfEnd As Long, bCurrByte As Byte, sResString As String
   
   If Not IsArray(bByteArray) Then GoTo ERR_HANDLER ' nothing to search!
   lLowBoundOfByteArray = LBound(bByteArray)
   lUpBoundOfByteArray = UBound(bByteArray)
   lFindLen = lUpBoundOfByteArray - lLowBoundOfByteArray + 1
   If lFindLen = 0 Then GoTo ERR_HANDLER ' nothing to search!
   
   On Error Resume Next
   Set FSO = CreateObject("Scripting.FileSystemObject")
   If Err.Number <> 0 Then GoTo ERR_HANDLER ' error accessing ActiveX object!
   If Not FSO.FileExists(sFileName) Then GoTo ERR_HANDLER ' no input file!
   Set FSO = Nothing
   On Error GoTo 0
   
   lFLen = FileLen(sFileName)
   If lStartPosition >= lFLen Then GoTo ERR_HANDLER
   
   lNumBlocks = (lFLen - lStartPosition + 1) \ BLOCK_SIZE ' we need do +1, because first byte number in file is 1, not 0
   lLastBlockSize = (lFLen - lStartPosition + 1) Mod BLOCK_SIZE
   
   iFN = FreeFile()
   Open sFileName For Binary Access Read Shared As iFN
      
      ' lLastBlockSize - we need to expand real file length to read all data
      ' without any additional steps. In fact, exceeded bytes beyond file length
      ' just be filled with ZERO
      For i = lStartPosition To lFLen + (BLOCK_SIZE - lLastBlockSize) Step BLOCK_SIZE
         Erase bReadBuff
         Get #iFN, i, bReadBuff
         For j = 1 To BLOCK_SIZE - lFindLen + 1
            boCompRes = True
            For k = lLowBoundOfByteArray To lUpBoundOfByteArray
               If bReadBuff(j + k - lLowBoundOfByteArray) <> bByteArray(k) Then
                  boCompRes = False
                  Exit For
               End If
'               If k = lUpBoundOfByteArray And boCompRes = True Then
'               End If
            Next k
            If boCompRes Then ' Bingo!!! :)
               lPosOfBegin = i + j ' - lLowBoundOfByteArray
               Exit For
            End If
         Next j
         If boCompRes Then Exit For
         ' then we must shift variable i on lFindLen value
         i = i - lFindLen + 1
      Next i
   
   If boCompRes Then
      FindStringInFile.bNoErrors = True
      
      ' here we step backward until found ZERO, LF or CR symbol
      For i = lPosOfBegin To lStartPosition Step -1
         Get #iFN, i, bCurrByte
         If bCurrByte = &H0 Or bCurrByte = &HA Or bCurrByte = &HD Then
            lPosOfBegin = i + 1
            Exit For
         End If
      Next i
      If i = lStartPosition - 1 Then lPosOfBegin = lStartPosition
      
      ' here we step forward until found ZERO, LF or CR symbol
      For i = lPosOfBegin To lFLen
         Get #iFN, i, bCurrByte
         If bCurrByte = &H0 Or bCurrByte = &HA Or bCurrByte = &HD Then
            lPosOfEnd = i - 1
            Exit For
         End If
      Next i
      If i = lFLen + 1 Then lPosOfEnd = lFLen
      
      sResString = vbNullString
      For i = lPosOfBegin To lPosOfEnd
         Get #iFN, i, bCurrByte
         sResString = sResString & Chr$(bCurrByte)
      Next i
      
      FindStringInFile.lResultPosition = lPosOfBegin
      FindStringInFile.sResultString = sResString
   Else
      FindStringInFile.bNoErrors = False
      FindStringInFile.lResultPosition = 0
      FindStringInFile.sResultString = vbNullString
   End If
   Close #iFN
   Exit Function

ERR_HANDLER:
      Err.Clear
      On Error GoTo 0
      FindStringInFile.bNoErrors = False
      FindStringInFile.lResultPosition = 0
      FindStringInFile.sResultString = vbNullString
End Function

Public Function PostScriptMadeByQuark(ByVal sFileName As String, Optional ByVal frmOWNER As Form = Nothing, Optional ByVal bAnalizeWholeFile As Boolean = False) As Boolean
    Const QRKMarks As String = "QuarkXPress_Reg_Marks"
    Erase vEPS_DATA
    vEPS_DATA = Array()
    vEPS_DATA = ReadEpsFileToStringArray(sFileName, 1, , , frmOWNER, "Looking for QuarkXPress marks, please wait ...", QRKMarks)
End Function

