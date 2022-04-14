Attribute VB_Name = "mExtendedMsgBox"
Option Explicit

Private Const WH_CBT As Long = &H5
Private Const HCBT_ACTIVATE As Long = &H5
Private Const STM_SETICON As Long = &H170
Private Const MODAL_WINDOW_CLASSNAME As String = "#32770"
Private Const SS_ICON As Long = &H3
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const STM_SETIMAGE As Long = &H172
Private Const IMAGE_CURSOR As Long = &H2

Private Const DCB_DISABLE = &H8
Private Const DCB_ENABLE = &H4
Private Const DCB_RESET = &H1
Private Const DCB_ACCUMULATE = &H2
Private Const DCB_SET = (DCB_RESET Or DCB_ACCUMULATE)

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Boolean
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Type ANICURSOR
   m_hCursor As Long
   m_hWnd As Long
End Type

Private pHook As Long
Private phIcon As Long
Private pAniIcon As String

Public Function XMsgBox(ByVal Message As String, _
               Optional ByVal MBoxStyle As VbMsgBoxStyle = vbOKOnly, _
               Optional ByVal Title As String = "", _
               Optional ByVal hIcon As Long = 0&, _
               Optional ByVal AniIcon As String = "") As VbMsgBoxResult
   
   ' Hook the msgbox with the function usual arguments,
   ' redirecting messages to MsgBoxHookProc.
   pHook = SetWindowsHookEx(WH_CBT, _
          AddressOf MsgBoxHookProc, _
                     App.hInstance, _
                 GetCurrentThreadId())
                 
   ' Save other arguments for use in MsgBoxHookProc
   phIcon = hIcon
   pAniIcon = AniIcon
   
   ' If a custom icon (animated or otherwise) is required
   ' make sure the msgbox makes room for it by setting the
   ' style to vbInformation; ensure other icon styles are set
   ' off, as if more than one are set no icon can be displayed.
   If Len(AniIcon) <> 0 Or phIcon <> 0 Then
      MBoxStyle = MBoxStyle And Not (vbCritical)
      MBoxStyle = MBoxStyle And Not (vbExclamation)
      MBoxStyle = MBoxStyle And Not (vbQuestion)
      MBoxStyle = MBoxStyle Or vbInformation
   End If
   
   ' Invoke the Msgbox; MsgBoxHookProc will take over from here.
   XMsgBox = MsgBox(Message, MBoxStyle Or vbMsgBoxSetForeground, Title)
End Function

Private Function MsgBoxHookProc(ByVal CodeNo As Long, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long) As Long
   Dim ClassNameSize As Long
   Dim sClassName As String
   Dim hIconWnd As Long
   Dim M As ANICURSOR
   
   ' Call the next hook; this is standard stuff.
   MsgBoxHookProc = CallNextHookEx(pHook, CodeNo, wParam, lParam)
   ' Only interfere if the msgbox activate message is being dealt with:
   If CodeNo = HCBT_ACTIVATE Then
      ' Check the classname; exit if not a standard msgbox.
      sClassName = Space$(32)
      ClassNameSize = GetClassName(wParam, sClassName, 32)
      If Left$(sClassName, ClassNameSize) <> MODAL_WINDOW_CLASSNAME Then Exit Function
   
      ' If displaying custom icon (animated or not), get icon window handle.
      If phIcon <> 0 Or Len(pAniIcon) <> 0 Then _
         hIconWnd = FindWindowEx(wParam, 0&, "Static", vbNullString)
      
      ' If custom (non- animated) icon, set here:
      If phIcon <> 0 Then SendMessage hIconWnd, STM_SETICON, phIcon, ByVal 0&
      
      ' If custom (animated) icon, set here: (animated takes precidence)
      If Len(pAniIcon) Then AniCreate M, pAniIcon, hIconWnd, 0, 0
      
      'unhook.
      UnhookWindowsHookEx pHook
   End If
End Function

Public Sub AniCreate(ByRef m_AniStuff As ANICURSOR, sAniName As String, hwndParent As Long, X As Long, Y As Long)
   ' Creates an animated cursor on hwndParent at x,y
   
   ' First destroy previous ani if m_AniStuff refers to one.
   AniDestroy m_AniStuff
   With m_AniStuff
      ' Get cursor.
      .m_hCursor = LoadCursorFromFile(sAniName)
      If .m_hCursor Then
         ' Create cursor window.
         .m_hWnd = CreateWindowEx(0, "Static", "", WS_CHILD Or WS_VISIBLE Or SS_ICON, ByVal 20, ByVal 20, 0, 0, hwndParent, 0, App.hInstance, ByVal 0)
         If .m_hWnd Then
            ' Place cursor in window & position
            SendMessage .m_hWnd, STM_SETIMAGE, IMAGE_CURSOR, ByVal .m_hCursor
            SetWindowPos .m_hWnd, 0, X, Y, 0, 0, SWP_NOZORDER Or SWP_NOSIZE
         Else
            ' Clean up.
            DestroyCursor .m_hCursor
         End If
      End If
   End With
End Sub

Public Sub AniDestroy(ByRef m_AniStuff As ANICURSOR)
   ' Destroy animated cursor referenced by m_AniStuff
   With m_AniStuff
      If .m_hCursor Then _
         If DestroyCursor(.m_hCursor) Then .m_hCursor = 0
      If IsWindow(.m_hWnd) Then _
         If DestroyWindow(.m_hWnd) Then .m_hWnd = 0
   End With
End Sub


Public Function ZMsgBox(ByVal sCaption As String, _
                        ByVal sMainMessage As String, _
                        ByVal sMessageHeader As String, _
                        ByVal sMessageFooter As String, _
                        ByRef piPicture As IPictureDisp, _
                        Optional ByVal bDefaultFirstButton As Boolean = False) As VbMsgBoxResult
    With frmMsgBox
        .msgRes = vbCancel
        .Caption = sCaption
        .txtHeader = sMessageHeader
        .txtFooter = sMessageFooter
        .txtMain = sMainMessage
        .buttonS(0).Default = bDefaultFirstButton
        .buttonS(1).Default = Not bDefaultFirstButton
        Set .Picture1 = piPicture
        .Show 1
        ZMsgBox = .msgRes
        .Hide
    End With
End Function
