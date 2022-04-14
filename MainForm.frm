VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Raster EPS Start"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "MainForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "What you want to do?"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      Begin VB.OptionButton optCDR 
         Caption         =   "&Import in CorelDRAW"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton optAI 
         Caption         =   "&Open with Adobe Illustrator"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton optPSD 
         Caption         =   "Open with &PhotoShop"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Adobe Photoshop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Determined raster EPS format created with program"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "FRMmAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum epsOpenWith
  voNothing = 0
  voAI8 = 8
  voAI10 = 10
  voAI11 = 11
  voPSD5 = 5
  voPSD9 = 9
End Enum

Public Enum qxdOpenWith
  voNothing = 0
  voQX3 = 13
  voQX4 = 14
  voQX5 = 15
  voQX6 = 16
End Enum

Dim sEPSFormat(0 To 3) As String, sPhotoshopColorMode(0 To 5) As String
Dim sTIFFormat(0 To 1) As String, sTIFFColorMode(0 To 8) As String

Dim tEPS As ftEPSInfo, tTIF As ftTIFFInfo

Dim EPSopen As epsOpenWith, CreatorName As String, ShopEPS As Boolean, ss$, ThisIsEPS As Boolean
Dim QXDopen As qxdOpenWith
'Public iniFile As String
'Dim prgPath(0 To 2) As String


Private Sub CancelButton_Click()
  Unload Me
  End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    Unload Me
    End
  End If
End Sub

Private Sub Form_Load()
  Dim i&, j&, sstr$
  Dim B1 As Byte, B2 As Byte, B3 As Byte, B4 As Byte
  Dim PSoffset As Long, CmdFile
  Dim Res As VbMsgBoxResult
  On Error GoTo StartError
    'CmdFile = GetCommandLine
    CmdFile = Command()
    If Len(CmdFile) = 0 Or Not ProgramInit Then
'      CreateAssociation "dtp_file", "SuperLauncher", ".eps", _
                        Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)
'      CreateAssociation "dtp_file", "SuperLauncher", ".qxd", _
                        Chr$(34) & App.Path & "\" & App.EXEName & ".exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)
'      CreateAssociation "dtp_file", "SuperLauncher", ".tif", _
                        App.Path & "\" & App.EXEName & ".exe %1"
      frmOptions.Show 1
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
    sTIFFColorMode(2) = "RGB"
    sTIFFColorMode(5) = "CMYK"
    sTIFFColorMode(8) = "LAB"
    
    sPhotoshopColorMode(0) = ""
    sPhotoshopColorMode(1) = "Grayscale"
    sPhotoshopColorMode(2) = "LAB"
    sPhotoshopColorMode(3) = "RGB"
    sPhotoshopColorMode(4) = "CMYK"
    sPhotoshopColorMode(5) = "Multichannel"
    
    
    ss$ = CmdFile ' берем первое имя файла
    Set FI = FS.GetFile(ss$)
    
    tEPS = GetEPSInfo(FI.Path)
    'If tEPS.EPSType <> epsNonEPSImage Then
    
    
    
    CreatorName = Replace(sstr$, "%%Creator: ", "")
    If InStr(sstr$, "Photoshop") <> 0 Then ShopEPS = True Else ShopEPS = False
    If Not ShopEPS Then
      Me.Label1.Caption = "Determined vector EPS format created with program"
      Me.Caption = "Vector EPS Start"
      If InStr(sstr$, "Corel") > 0 Then Me.optCDR.Value = True Else Me.optAI.Value = True
    End If
    Me.Label2.Caption = CreatorName
    Exit Sub
StartError:
  If Err.Number = 62 Then ' 62 - попытка прочитать с позиции за концом файла
    MsgBox "Error in EPS file!", vbCritical, "EPS Start"
    End
  End If
  MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, "EPS Start ERROR"
  End
End Sub

Private Sub OKButton_Click()
'  If Me.optCDR = True Then EPSopen = voCDR
'  If Me.optAI = True Then EPSopen = voAI
'  If Me.optPSD = True Then EPSopen = voPSD
'  Select Case EPSopen
'    Case voPSD
'      Call Shell(prgPath(2) & " " & ss$, vbMaximizedFocus)
'    Case voAI
'      Call Shell(prgPath(1) & " " & ss$, vbMaximizedFocus)
'  End Select
'  End
End Sub

Private Sub optAI_DblClick()
  OKButton_Click
End Sub

Private Sub optCDR_DblClick()
  OKButton_Click
End Sub

Private Sub optPSD_DblClick()
  OKButton_Click
End Sub

Private Function ProgramInit() As Boolean
  Dim NN As Integer
'  iniFile = App.Path & "\EPSStart.ini"
'  If Not frmOptions.FileExists(iniFile) Then
'    ProgramInit = False
'  Else
'    NN = FreeFile
'    Open iniFile For Input As NN
'    On Error GoTo InitErr
'    Line Input #NN, prgPath(0) ' путь к Corel Draw
'    Line Input #NN, prgPath(1) ' путь к Adobe Illustrator
'    Line Input #NN, prgPath(2) ' путь к Adobe Photoshop
'    Close #NN
'    frmOptions.txtPath(0).Text = prgPath(0)
'    frmOptions.txtPath(1).Text = prgPath(1)
'    frmOptions.txtPath(2).Text = prgPath(2)
'    ProgramInit = True
'  End If
'  Exit Function
'InitErr:
'  Close
'  Err.Clear
'  ProgramInit = False
End Function
