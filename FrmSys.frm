VERSION 5.00
Begin VB.Form FrmSysTray 
   BorderStyle     =   0  'None
   ClientHeight    =   735
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2055
   Icon            =   "FrmSys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Flash2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   720
      Picture         =   "FrmSys.frx":0742
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox Flash1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "FrmSys.frx":0CCC
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Timer TmrFlash 
      Interval        =   1000
      Left            =   1440
      Top             =   120
   End
   Begin VB.Menu mPopupMenu 
      Caption         =   "&PopupMenu"
      Begin VB.Menu mOpenMacFilesOnPC 
         Caption         =   "&Open MAC files on PC"
      End
      Begin VB.Menu mUseVMWare 
         Caption         =   "&Use VMWare MacOSX virtual mashine"
      End
      Begin VB.Menu mSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mUse2Dist 
         Caption         =   "Use two &Distillers - main and portable"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mAssoc 
         Caption         =   "Edit &INI file manually"
      End
      Begin VB.Menu mMaximize 
         Caption         =   "Ma&ximize"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mRestore 
         Caption         =   "&Restore"
         Visible         =   0   'False
      End
      Begin VB.Menu mMinimize 
         Caption         =   "&Minimize"
         Visible         =   0   'False
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mCloseMenu 
         Caption         =   "&Close this menu"
      End
      Begin VB.Menu mSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public WithEvents FSys As Form
Attribute FSys.VB_VarHelpID = -1
Public Event Click(ClickWhat As String)
Public Event TIcon(F As Form)

Private nid As NOTIFYICONDATA
Private LastWindowState As Integer, hwndIcon As Long

Public Property Let Tooltip(Value As String)
        
        On Error GoTo Tooltip_Err

        nid.szTip = Value & vbNullChar

        
        Exit Property

Tooltip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Tooltip " & _
               "at line " & Erl
        End
        
End Property

Public Property Get Tooltip() As String
        
        On Error GoTo Tooltip_Err
        

        Tooltip = nid.szTip

        
        Exit Property

Tooltip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Tooltip " & _
               "at line " & Erl
        End
        
End Property

Public Property Let Interval(Value As Integer)
        
        On Error GoTo Interval_Err
        

        TmrFlash.Interval = Value
        UpdateIcon NIM_MODIFY

        
        Exit Property

Interval_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Interval " & _
               "at line " & Erl
        End
        
End Property

Public Property Get Interval() As Integer
        
        On Error GoTo Interval_Err
        

        Interval = TmrFlash.Interval

        
        Exit Property

Interval_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Interval " & _
               "at line " & Erl
        End
        
End Property

Public Property Let TrayIcon(Value)
        
        On Error GoTo TrayIcon_Err
        

        TmrFlash.Enabled = False
        On Error Resume Next
        ' Value can be a picturebox, image, form or string

        Select Case TypeName(Value)

            Case "PictureBox", "Image"
                Me.Icon = Value.Picture
                TmrFlash.Enabled = False
                RaiseEvent TIcon(Me)

            Case "String"

                If (UCase(Value) = "DEFAULT") Then

                    TmrFlash.Enabled = True
                    Me.Icon = Flash2.Picture
                    RaiseEvent TIcon(Me)

                Else

                    ' Sting is filename; load icon from picture file.
                    TmrFlash.Enabled = True
                    Me.Icon = LoadPicture(Value)
                   RaiseEvent TIcon(Me)

                End If
            Case "Long"
                hwndIcon = Value
                RaiseEvent TIcon(Me)
            Case Else
                ' It's a form ?
                Me.Icon = Value.Icon
                RaiseEvent TIcon(Me)

        End Select

        If Err.Number <> 0 Then TmrFlash.Enabled = True

        UpdateIcon NIM_MODIFY, hwndIcon

        
        Exit Property

TrayIcon_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.TrayIcon " & _
               "at line " & Erl
        End
        
End Property

Private Sub Form_Load()
        
        On Error GoTo Form_Load_Err
        

        Me.Icon = Flash1
        RaiseEvent TIcon(Me)
        Me.Visible = False
        TmrFlash.Enabled = True
        Tooltip = App.EXEName
        mAbout.Caption = "About " & App.EXEName
        UpdateIcon NIM_ADD

        
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Form_Load " & _
               "at line " & Erl
        End
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        On Error GoTo Form_MouseMove_Err
        

        Dim result As Long
        Dim msg As Long
   
        ' The Form_MouseMove is intercepted to give systray mouse events.

        If Me.ScaleMode = vbPixels Then

            msg = X

        Else

            msg = X / Screen.TwipsPerPixelX

        End If
      
        Select Case msg

            Case WM_RBUTTONDBLCLK
                RaiseEvent Click("RBUTTONDBLCLK")

            Case WM_RBUTTONDOWN
                RaiseEvent Click("RBUTTONDOWN")

            Case WM_RBUTTONUP
                ' Popup menu: selectively enable items dependent on context.

'                Select Case FSys.Visible

'                    Case True

'                        Select Case FSys.WindowState

'                            Case vbMaximized
'                                mMaximize.Enabled = False
'                                mMinimize.Enabled = True
'                                mRestore.Enabled = False

'                            Case vbNormal
'                                mMaximize.Enabled = False
'                                mMinimize.Enabled = True
'                                mRestore.Enabled = False

'                            Case vbMinimized
'                                mMaximize.Enabled = False
'                                mMinimize.Enabled = False
'                                mRestore.Enabled = True

'                            Case Else
'                                mMaximize.Enabled = False
'                                mMinimize.Enabled = True
'                                mRestore.Enabled = True

'                        End Select

'                    Case Else
'                        mRestore.Enabled = True
'                        mMaximize.Enabled = False
'                        mMinimize.Enabled = False

'                End Select
         
                RaiseEvent Click("RBUTTONUP")
                PopupMenu mPopupMenu

            Case WM_LBUTTONDBLCLK
                RaiseEvent Click("LBUTTONDBLCLK")
                'mRestore_Click

            Case WM_LBUTTONDOWN
                RaiseEvent Click("LBUTTONDOWN")

            Case WM_LBUTTONUP
                RaiseEvent Click("LBUTTONUP")

            Case WM_MBUTTONDBLCLK
                RaiseEvent Click("MBUTTONDBLCLK")

            Case WM_MBUTTONDOWN
                RaiseEvent Click("MBUTTONDOWN")

            Case WM_MBUTTONUP
                RaiseEvent Click("MBUTTONUP")

            Case WM_MOUSEMOVE
                RaiseEvent Click("MOUSEMOVE")

            Case Else
                RaiseEvent Click("OTHER....: " & Format$(msg))

        End Select

        
        Exit Sub

Form_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.Form_MouseMove " & _
               "at line " & Erl
        End
        
End Sub

Private Sub FSys_Resize()
    
    'On Error Resume Next
    

    ' Event generated my main form. WindowState is stored in LastWindowState, so that
    ' it may be re- set when the menu item "Restore" is selected.

    'If (FSys.WindowState <> vbMinimized) Then LastWindowState = FSys.WindowState

End Sub

Private Sub FSys_Unload(Cancel As Integer)
        
        On Error GoTo FSys_Unload_Err
        

        ' Important: remove icon from tray, and unload this form when
        ' the main form is unloaded.
        UpdateIcon NIM_DELETE
        Unload Me

        
        Exit Sub

FSys_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.FSys_Unload " & _
               "at line " & Erl
        End
        
End Sub

Private Sub mAbout_Click()
        
        On Error GoTo mAbout_Click_Err
        

        MsgBox "SuperStarter project  v." & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
                "Selective and flexible launcher for many types of DesktopPublishing documents." & vbCrLf & vbCrLf & _
                "© Copyright 2008-2012, Alex Commandor (alex.commandor@gmail.com) ;)", vbInformation, "About SuperStarter"

        
        Exit Sub

mAbout_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mAbout_Click " & _
               "at line " & Erl
        End
        
End Sub

Private Sub mAssoc_Click()
    On Error Resume Next
    Call Shell("notepad.exe " & Chr$(34) & sIniFile & Chr$(34), vbMaximizedFocus)
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub mMaximize_Click()
        
        On Error GoTo mMaximize_Click_Err
        

        'FSys.WindowState = vbMaximized
        'FSys.Show

        
        Exit Sub

mMaximize_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mMaximize_Click " & _
               "at line " & Erl
        End
        
End Sub

Private Sub mMinimize_Click()
        
        On Error GoTo mMinimize_Click_Err
        

        'FSys.WindowState = vbMinimized

        
        Exit Sub

mMinimize_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mMinimize_Click " & _
               "at line " & Erl
        End
        
End Sub

Public Sub mExit_Click()
        
        On Error GoTo mExit_Click_Err
        

        UpdateIcon NIM_DELETE
        Unload FSys
        End

        
        Exit Sub

mExit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mExit_Click " & _
               "at line " & Erl
        End
        
End Sub

Private Sub mOpenMacFilesOnPC_Click()
    mOpenMacFilesOnPC.Checked = Not mOpenMacFilesOnPC.Checked
    If mOpenMacFilesOnPC.Checked Then
        frmMain.Ini1.Writing "MAIN", "OPEN_MAC_ON_PC", "YES", sIniFile
    Else
        frmMain.Ini1.Writing "MAIN", "OPEN_MAC_ON_PC", "NO", sIniFile
    End If
End Sub

Private Sub mRestore_Click()
        
        On Error GoTo mRestore_Click_Err
        

        ' Don't "restore"  FSys is visible and not minimized.

        'If (FSys.Visible And FSys.WindowState <> vbMinimized) Then Exit Sub

        ' Restore LastWindowState
        'FSys.WindowState = LastWindowState
        'FSys.Visible = True
        'SetForegroundWindow FSys.hwnd

        
        Exit Sub

mRestore_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mRestore_Click " & _
               "at line " & Erl
        End
        
End Sub

Private Sub UpdateIcon(Value As Long, Optional ByVal hndlIcon As Long = 0)
        
        On Error GoTo UpdateIcon_Err
        

        ' Used to add, modify and delete icon.

        With nid

            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uID = vbNull
            .uFlags = NIM_DELETE Or NIF_TIP Or NIM_MODIFY
            .uCallbackMessage = WM_MOUSEMOVE
            If hndlIcon = 0 Then
                .hIcon = Me.Icon
            Else
                .hIcon = hndlIcon
            End If

        End With

        Shell_NotifyIcon Value, nid

        
        Exit Sub

UpdateIcon_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.UpdateIcon " & _
               "at line " & Erl
        End
        
End Sub

Public Sub MeQueryUnload(ByRef F As Form, Cancel As Integer, UnloadMode As Integer)
        
        On Error GoTo MeQueryUnload_Err
        

'        If UnloadMode = vbFormControlMenu Then'

            ' Cancel by setting Cancel = 1, minimize and hide main window.
'            Cancel = 1
'            F.WindowState = vbMinimized
'            F.Hide

'        End If

        
        Exit Sub

MeQueryUnload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.MeQueryUnload " & _
               "at line " & Erl
        End
        
End Sub

Public Sub MeResize(ByRef F As Form)
        
        On Error GoTo MeResize_Err
        

'        Select Case F.WindowState

'            Case vbNormal, vbMaximized
                ' Store LastWindowState
'                LastWindowState = F.WindowState

'            Case vbMinimized
'                F.Hide

'        End Select

        
        Exit Sub

MeResize_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.MeResize " & _
               "at line " & Erl
        End
        
End Sub

Private Sub mStart_Click()
        
        On Error GoTo mStart_Click_Err
        

        'Call FSys.btnStart_Click

        
        Exit Sub

mStart_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in SuperStarter.FrmSysTray.mStart_Click " & _
               "at line " & Erl
        End
        
End Sub

Private Sub mUse2Dist_Click()
    mUse2Dist.Checked = Not mUse2Dist.Checked
    If mUse2Dist.Checked Then
        frmMain.Ini1.Writing "MAIN", "USE_2_DISTILLERS", "YES", sIniFile
    Else
        frmMain.Ini1.Writing "MAIN", "USE_2_DISTILLERS", "NO", sIniFile
    End If
End Sub

Private Sub mUseVMWare_Click()
    mUseVMWare.Checked = Not mUseVMWare.Checked
    If mUseVMWare.Checked Then
        frmMain.Ini1.Writing "MAIN", "USE_VMWARE", "YES", sIniFile
    Else
        frmMain.Ini1.Writing "MAIN", "USE_VMWARE", "NO", sIniFile
    End If
    MsgBox "This option still in development! Sorry :(", vbExclamation, "SuperStarter"
End Sub

Private Sub TmrFlash_Timer()
    
    On Error Resume Next
    

    ' Change icon.
    Static LastIconWasFlash1 As Boolean
    LastIconWasFlash1 = Not LastIconWasFlash1

    Select Case LastIconWasFlash1

        Case True
            Me.Icon = Flash2

        Case Else
            Me.Icon = Flash1

    End Select

    RaiseEvent TIcon(Me)
    UpdateIcon NIM_MODIFY

End Sub

