VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3030
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   550
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   5655
   End
   Begin VB.TextBox txtFooter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtMain 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   3  'Not Merge Pen
      FillColor       =   &H8000000F&
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   240
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   3
      Top             =   240
      Width           =   1920
   End
   Begin VB.CommandButton buttonS 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton buttonS 
      Caption         =   "&No"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton buttonS 
      Caption         =   "&Yes"
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public msgRes As VbMsgBoxResult

Private Sub buttonS_Click(Index As Integer)
    Select Case Index
        Case 0 ' First Button
            msgRes = vbYes
        Case 1 ' Second Button
            msgRes = vbNo
        Case 2 ' Third Button
            msgRes = vbCancel
    End Select
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Hide
    Me.txtFooter.FontBold = True
    Me.txtHeader.FontBold = True
End Sub
