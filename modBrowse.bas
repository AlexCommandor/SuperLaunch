Attribute VB_Name = "modBrowse"
Option Explicit

' Maximun long filename path length
Public Const MAX_PATH = 260

' An item identifier is defined by the variable-length SHITEMID structure.
' The first two bytes of this structure specify its size, and the format of
' the remaining bytes depends on the parent folder, or more precisely
' on the software that implements the parent folder’s IShellFolder interface.
' Except for the first two bytes, item identifiers are not strictly defined, and
' applications should make no assumptions about their format.


Type SHITEMID   ' mkid
    cb As Long       ' Size of the ID (including cb itself)
    abID() As Byte  ' The item ID (variable length)
End Type

Type SHFILEINFO   ' shfi
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Type ITEMIDLIST   ' idl
    mkid As SHITEMID
End Type
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                              (ByVal pIdl As Long, ByVal pszPath As String) As Long

' Frees memory allocated by SHBrowseForFolder()
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

' Displays a dialog box that enables the user to select a shell folder.
' Returns a pointer to an item identifier list that specifies the location
' of the selected folder relative to the root of the name space. If the user
' chooses the Cancel button in the dialog box, the return value is NULL.
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                              (lpBrowseInfo As BROWSEINFO) As Long ' ITEMIDLIST

' Contains parameters for the the SHBrowseForFolder function and receives
' information about the folder selected by the user.
Public Type BROWSEINFO   ' bi
    
    ' Handle of the owner window for the dialog box.
    hOwner As Long
    
    ' Pointer to an item identifier list (an ITEMIDLIST structure) specifying the location
    ' of the "root" folder to browse from. Only the specified folder and its subfolders
    ' appear in the dialog box. This member can be NULL, and in that case, the
    ' name space root (the desktop folder) is used.
    pidlRoot As Long
    
    ' Pointer to a buffer that receives the display name of the folder selected by the
    ' user. The size of this buffer is assumed to be MAX_PATH bytes.
    pszDisplayName As String
    
    ' Pointer to a null-terminated string that is displayed above the tree view control
    ' in the dialog box. This string can be used to specify instructions to the user.
    lpszTitle As String
    
    ' Value specifying the types of folders to be listed in the dialog box as well as
    ' other options. This member can include zero or more of the following values below.
    ulFlags As Long
    
    ' Address an application-defined function that the dialog box calls when events
    ' occur. For more information, see the description of the BrowseCallbackProc
    ' function. This member can be NULL.
    lpfn As Long
    
    ' Application-defined value that the dialog box passes to the callback function
    ' (if one is specified).
    lParam As Long
    
    ' Variable that receives the image associated with the selected folder. The image
    ' is specified as an index to the system image list.
    iImage As Long

End Type

' BROWSEINFO ulFlags values:
' Value specifying the types of folders to be listed in the dialog box as well as
' other options. This member can include zero or more of the following values:

' Only returns file system directories. If the user selects folders
' that are not part of the file system, the OK button is grayed.
Public Const BIF_RETURNONLYFSDIRS = &H1

' Does not include network folders below the domain level in the tree view control.
' For starting the Find Computer
Public Const BIF_DONTGOBELOWDOMAIN = &H2

' Includes a status area in the dialog box. The callback function can set
' the status text by sending messages to the dialog box.
Public Const BIF_STATUSTEXT = &H4


'// get the options
'Private Function GetReturnType() As Long
'  Dim dwRtn As Long
'  If chkRtnType(0) Then dwRtn = dwRtn Or BIF_RETURNONLYFSDIRS
'  If chkRtnType(1) Then dwRtn = dwRtn Or BIF_DONTGOBELOWDOMAIN
'  If chkRtnType(3) Then dwRtn = dwRtn Or BIF_RETURNFSANCESTORS
'  If chkRtnType(4) Then dwRtn = dwRtn Or BIF_BROWSEFORCOMPUTER
'  If chkRtnType(5) Then dwRtn = dwRtn Or BIF_BROWSEFORPRINTER
'  GetReturnType = dwRtn
'End Function

Public Function BrowseForFolder(ByVal hWnd As Long, Optional ByVal sTitle = "Browse for Folder") As String
  Dim BI As BROWSEINFO
  Dim nFolder As Long
  Dim IDL As ITEMIDLIST
  Dim pIdl As Long
  Dim sPath As String
  Dim SHFI As SHFILEINFO
  Dim txtPath As String
  Dim txtDisplayName As String
  
  With BI
    '// The dialog'//s owner window...
    .hOwner = hWnd
    
    '// Initialize the buffer that rtns the display name of the selected folder
    .pszDisplayName = String$(MAX_PATH, 0)
    
    '// Set the dialog'//s banner text
    .lpszTitle = sTitle
    
    '// Set the type of folders to display & return
    '// -play with these option constants to see what can be returned
    .ulFlags = BIF_RETURNONLYFSDIRS
    
  End With
  
  '// Clear previous return vals before the
  '// dialog is shown (it might be cancelled)
  txtPath = ""
  txtDisplayName = ""
  '// if you stop code execution between here and the
  '// end of this sub, you will be wasting memory.
  '// you need to call CoTaskMemFree pIdl to free the
  '// memory used by SHBrowseForFolder
  
  '// Show the Browse dialog
  pIdl = SHBrowseForFolder(BI)
  
  '// If the dialog was cancelled...
  If pIdl = 0 Then BrowseForFolder = vbNullString: Exit Function
    
  '// Fill sPath w/ the selected path from the id list
  '// (will rtn False if the id list can'//t be converted)
  sPath = String$(MAX_PATH, 0)
  SHGetPathFromIDList ByVal pIdl, ByVal sPath

  '// Display the path and the name of the selected folder
  txtPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
  txtDisplayName = Left$(BI.pszDisplayName, _
                             InStr(BI.pszDisplayName, vbNullChar) - 1)
  
  '// Frees the memory SHBrowseForFolder()
  '// allocated for the pointer to the item id list
  CoTaskMemFree pIdl
  BrowseForFolder = txtPath
End Function






