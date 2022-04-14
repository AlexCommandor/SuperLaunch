VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4590
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5745
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5535
      Begin VB.OptionButton optQXD 
         Caption         =   "Option2"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optEPS 
         Caption         =   "Option1"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   3975
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3975
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2010
      TabIndex        =   0
      Top             =   3975
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
  Dim NN As Integer, Res As Integer
'  On Error GoTo lblErr
'  If Not FileExists(Me.txtPath(0).Text) Or Not FileExists(Me.txtPath(1).Text) _
'    Or Not FileExists(Me.txtPath(2).Text) Then
'    Res = MsgBox("One or more paths incorrect! Continue?", vbYesNo, "EPS Start")
'    If Res = vbNo Then Exit Sub
'  End If
'  NN = FreeFile
'  Open MainForm.iniFile For Output As NN
'    Print #NN, txtPath(0).Text
'    Print #NN, txtPath(1).Text
'    Print #NN, txtPath(2).Text
'  Close #NN
'  'MsgBox "Parameters saved.", vbOKOnly, "EPS Start"
'  Me.cmdApply.Enabled = False
'  Exit Sub
'lblErr:
'  MsgBox Err.Description
'  End
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
  On Error Resume Next
'  With Me.ComDial
'    Select Case Index
'      Case 0
'        .DialogTitle = "Locate file CORELDRW.EXE"
'      Case 1
'        .DialogTitle = "Locate file ILLUSTRATOR.EXE"
'      Case 2
'        .DialogTitle = "Locate file PHOTOSHP.EXE"
'    End Select
'    .ShowOpen
'    If Err.Number <> 0 Then Exit Sub
'    If FileExists(.FileName) Then
'      Me.txtPath(Index).Text = .FileName
'      Me.cmdApply.Enabled = True
'    End If
'  End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
  cmdApply_Click
  Unload Me
End Sub

Public Function FileExists(strFileName As String) As Boolean
Dim nFile As Integer
On Error Resume Next
nFile = FreeFile
Open strFileName For Input As nFile
If Err.Number <> 0 Then
  Err.Clear
  FileExists = False
Else
  Close nFile
  FileExists = True
End If
End Function
