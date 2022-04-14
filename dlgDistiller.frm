VERSION 5.00
Begin VB.Form dlgDistiller 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Distiller options"
   ClientHeight    =   4740
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "dlgDistiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    End
End Sub

Private Sub Form_Load()
  Dim sKeyName As String  ' Holds Key Name in registry.
  Dim sKeyValue As String ' Holds Key Value in registry.
  Dim ret As Long         ' Holds error status if any from
  ' API calls.
  Dim lphKey As Long      ' Holds created key handle from
  ' RegCreateKey.
    RegOpenKey(HKEY_CLASSES_ROOT
    'GetSpecFolder
End Sub

Private Function GetSpecFolder(ByVal nFolder As Long) As String

  Dim wIdx As Integer, nFolder As Long
  Dim sPath As String * MAX_PATH   ' 260
  Dim IDL As ITEMIDLIST
  
  ' Loads the labels with the respective
  ' system folder's path (if found)
  'For wIdx = 1 To 17
  wIdx = nFolder
    nFolder = GetFolderValue(wIdx)

    ' Fill the item id list with the pointer of each folder item, rtns 0 on success
    If SHGetSpecialFolderLocation(Me.hwnd, nFolder, IDL) = NOERROR Then
      
      ' Get the path from the item id list pointer, rtns True on success
      If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
    
        ' Display the path in the respective label
        GetSpecFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1)
      
      End If
    
    Else
      ' The folder item doesn't exist, disable it's checkbox
      GetSpecFolder = ""
    
    End If
  'Next
  
End Function
