Attribute VB_Name = "modStdDialog"
Option Explicit

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias _
    "GetSaveFileNameA" (pOpenfilename As OpenFilename) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
      "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long

Private Type OpenFilename
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    iFilterIndex As Long
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

Public Enum OFNFlagsEnum
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
End Enum

' Show the common dialog to select a file to save. Returns the path of the
' selected file or a null string if the dialog is canceled
' Parameters:
'  - sFilter is used to specify what type(s) of files will be shown
'  - sDefExt is the default extension associated to a file name if no one is
' specified by the user
'  - sInitDir is the directory that will be open when the dialog is shown
'  - lFlag is a combination of Flags for the dialog. Look at the Common
' Dialogs' Help for more informations
'  - hParent is the handle of the parent form

' Example:
'    Dim sFilter As String
'    'set the filter: show text files and all the files
'    sFilter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
'    'let the user select a file, an ask for confirmation if the file already
' exists
'    MsgBox "File selected: " & ShowOpenFileDialog(sFilter, "txt",
'  "C:\Documents", OFN_OVERWRITEPROMPT)

Function ShowSaveFileDialog(ByVal sFilter As String, Optional ByVal sDefExt As _
    String, Optional ByVal sInitDir As String, Optional ByVal lFlags As OFNFlagsEnum, _
    Optional ByVal hParent As Long) As String
    Dim OFN As OpenFilename
    On Error Resume Next
    
    ' set the values for the OpenFileName struct
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hParent
        .lpstrFilter = Replace(sFilter, "|", vbNullChar) & vbNullChar
        .lpstrFile = Space$(255) & vbNullChar & vbNullChar
        .nMaxFile = Len(.lpstrFile)
        .flags = lFlags
        .lpstrInitialDir = sInitDir
        .lpstrDefExt = sDefExt
    End With
    
    ' show the dialog
    If GetSaveFileName(OFN) Then
        ' extract the selected file (including the path)
        ShowSaveFileDialog = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, _
            vbNullChar) - 1)
    End If
End Function

Function ShowOpenFileDialog(ByVal sFilter As String, Optional ByVal sDefExt As _
    String, Optional ByVal sInitDir As String, Optional ByVal lFlags As OFNFlagsEnum, _
    Optional ByVal hParent As Long) As String
    Dim OFN As OpenFilename
    On Error Resume Next
    
    ' set the values for the OpenFileName struct
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hParent
        .lpstrFilter = Replace(sFilter, "|", vbNullChar) & vbNullChar
        .lpstrFile = Space$(255) & vbNullChar & vbNullChar
        .nMaxFile = Len(.lpstrFile)
        .flags = lFlags
        .lpstrInitialDir = sInitDir
        .lpstrDefExt = sDefExt
    End With
    
    ' show the dialog
    If GetOpenFileName(OFN) Then
        ' extract the selected file (including the path)
        ShowOpenFileDialog = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, _
            vbNullChar) - 1)
    End If
End Function

