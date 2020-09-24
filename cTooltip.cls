VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Custom Tooltip Class

Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharset As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rectangle) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nMultiplier As Long, ByVal nDivisor As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Enum Consts
    SWP_NOSIZE = &H1
    SWP_NOMOVE = &H2
    SWP_NOACTIVATE = &H10
    SWP_FLAGS = SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    SWP_TOPMOST = -1
    CW_USEDEFAULT = &H80000000
    TTS_STANDARD = 0
    TTS_BALLOON = &H40
    TTS_ALWAYSTIP = 1               'display even if parent window is inactive
    TTS_NOPREFIX = 2                'does not remove "&" from text
    TTDT_AUTOPOP = 2
    WM_USER = &H400
    TTM_SETDELAYTIME = WM_USER + 3
    TTM_ADDTOOL = WM_USER + 4
    TTM_GETTEXT = WM_USER + 11
    TTM_SETTIPBKCOLOR = WM_USER + 19
    TTM_SETTIPTEXTCOLOR = WM_USER + 20
    TTM_SETMAXTIPWIDTH = WM_USER + 24
    TTM_SETTITLE = WM_USER + 32
    WM_SETFONT = &H30
    TTF_CENTERTIP = 2               'center tool on parent
    TTF_SUBCLASS = &H10             'use implicit subclassing
    fwDontCare = 0
    fwBold = 700
    DEFAULT_CHARSET = 1
    LOGPIXELSY = 90
    TTDT_INITIAL = 3
End Enum
#If False Then ':) Line inserted by Formatter
Private SWP_NOSIZE, SWP_NOMOVE, SWP_NOACTIVATE, SWP_FLAGS, SWP_TOPMOST, CW_USEDEFAULT, TTS_STANDARD, TTS_BALLOON, TTS_ALWAYSTIP, TTS_NOPREFIX, _
        TTDT_AUTOPOP, WM_USER, TTM_SETDELAYTIME, TTM_ADDTOOL, TTM_GETTEXT, TTM_SETTIPBKCOLOR, TTM_SETTIPTEXTCOLOR, TTM_SETMAXTIPWIDTH, _
        TTM_SETTITLE, WM_SETFONT, TTF_CENTERTIP, TTF_SUBCLASS, fwDontCare, fwBold, DEFAULT_CHARSET, LOGPIXELSY, TTDT_INITIAL ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Public Enum TTStyle
    TTStandardIfActive = TTS_STANDARD                   'suppress if parent form is not active
    TTBalloonIfActive = TTS_BALLOON                     'suppress if parent form is not active
    TTStandardAlways = TTS_STANDARD Or TTS_ALWAYSTIP    'display even if parent form is not active
    TTBalloonAlways = TTS_BALLOON Or TTS_ALWAYSTIP      'display even if parent form is not active
    TTNone = -1                                         'kill tooltip (this is simply treated as illegal, so after killing the current tip no new one is created)
End Enum
#If False Then ':) Line inserted by Formatter
Private TTStandardIfActive, TTBalloonIfActive, TTStandardAlways, TTBalloonAlways, TTNone ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Public Enum TTIcon
    TTIconNone = 0
    TTIconInfo = 1         'i in white balloon
    TTIconWarning = 2      '! in yellow triangle
    TTIconError = 3        'x in red circle
    'all have a light gray shadow so be careful when selecting the ToolTip BackColor
End Enum
#If False Then ':) Line inserted by Formatter
Private TTIconNone, TTIconInfo, TTIconWarning, TTIconError ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Private Type Rectangle
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type tToolInfo
    ttSize      As Long
    myFlags     As Long
    ttParhWnd   As Long
    ttId        As Long
    ParentRect  As Rectangle
    hInstance   As Long
    pText       As String
    lParam      As Long
End Type
Private ToolInfo                    As tToolInfo

Private Const ToolTipWindowClassName As String = "Tooltips_Class32"
Private Const defFontName           As String = "Tahoma"
Private Const defFontSize           As Long = 8

Private TThWnd                      As Long     'Tooltip window handle
Private TThDC                       As Long     'Tooltip devive context
Private TThFont                     As Long     'Tooltip font handle

'my properties
Private myStyle                     As TTStyle
Private myIcon                      As TTIcon
Private myForeColor                 As Long
Private myBackColor                 As Long
Private myText                      As String
Private myTitle                     As String   'has the current title
Private myHoverTime                 As Long     'time im millisecs (-1 = use default)
Private myPopupTime                 As Long     'time im millisecs (-1 = use default)
Private myInitialText               As Variant  'has the initial text
Private myInitialTitle              As Variant  'has the initial title
Private myFontName                  As String
Private myFontSize                  As Long
Private myFontBold                  As Boolean
Private myFontItalic                As Boolean

Public Property Get BackCol() As Long

  'this returns the current tooltip backcolor

    BackCol = myBackColor

End Property

Public Property Get Centered() As Boolean

  'this returns the current tooltip alignment

    Centered = CBool(ToolInfo.myFlags And TTF_CENTERTIP)

End Property

Private Sub Class_Initialize()

    InitCommonControls 'doesn't matter that this is called for every class instance
    myStyle = TTNone

End Sub

Private Sub Class_Terminate()

  'kill tooltip window if one exists

    KillFont 'kill the font object if one exists
    If TThWnd Then
        ReleaseDC TThWnd, TThDC 'release the popup window's device context
        TThDC = 0
        DestroyWindow TThWnd 'and destroy the window itself
        TThWnd = 0
    End If
    myStyle = TTNone

End Sub

Public Function Create(Parent As Control, _
                       Text As String, _
                       Optional ByVal Style As TTStyle = TTBalloonAlways, _
                       Optional ByVal Centered As Boolean = False, _
                       Optional ByVal Icon As TTIcon = TTIconNone, _
                       Optional Title As String = vbNullString, _
                       Optional ByVal ForeColor As Long = vbButtonText, _
                       Optional ByVal BackColor As Long = vbInfoBackground, _
                       Optional ByVal HoverTime As Long = -1, _
                       Optional ByVal PopupTime As Long = -1) As Long

  'Create the tooltip window for the tooltip's parent control if that has an hWnd
  'This can now also create custom tooltips for hWnd-less controls (one at a time)
  'just supply a fake hWnd (normally the containing form.hWnd) for windowless controls

    Class_Terminate 'kill font and tooltip window if one exists
    With ToolInfo
        On Error Resume Next
            .ttParhWnd = Parent.hWnd 'the control's hWnd
            If Err Then 'has no hWnd
                Err.Clear
                .ttParhWnd = Parent.Parent.hWnd 'so use the control's parent's hWnd temporarily
            End If
            If (Err = 0) And _
                (Style = TTBalloonAlways Or Style = TTStandardAlways Or Style = TTBalloonIfActive Or Style = TTStandardIfActive) And _
                (Icon = TTIconError Or Icon = TTIconInfo Or Icon = TTIconNone Or Icon = TTIconWarning) Then
                'the tooltip's parent control (or the parent of the tooltip's parent control) has an hWnd and the params are acceptable
                .ttSize = Len(ToolInfo)
                .myFlags = TTF_SUBCLASS Or IIf(Centered, TTF_CENTERTIP, 0&)
                GetClientRect .ttParhWnd, .ParentRect
                .hInstance = App.hInstance
                myTitle = Title
                If myInitialTitle = Empty Then
                    myInitialTitle = myTitle
                End If
                myText = Replace$(Text, "|", vbCrLf) 'the vertical bar is used as line break character
                .pText = myText
                If myInitialText = Empty Then
                    myInitialText = myText
                End If
                If ForeColor < 0 Then
                    ForeColor = GetSysColor(ForeColor And &H7FFFFFFF)
                End If
                If BackColor < 0 Then
                    BackColor = GetSysColor(BackColor And &H7FFFFFFF)
                End If
                If ForeColor = BackColor Then
                    ForeColor = vbButtonText
                    BackColor = vbInfoBackground
                End If
                myForeColor = ForeColor
                myBackColor = BackColor
                myStyle = Style
                myIcon = Icon
                myHoverTime = HoverTime
                If myHoverTime >= 0 And myHoverTime < 50 Then 'at least 50 millisecs
                    myHoverTime = 50
                End If
                If PopupTime < 0 Then 'autotime depending on length of text
                    myPopupTime = Len(myText) * 45 + 1500
                  Else 'NOT POPUPTIME...
                    myPopupTime = PopupTime
                End If
                'create tooltip window and set it's properties
                TThWnd = CreateWindowEx(0&, ToolTipWindowClassName, vbNullString, TTS_NOPREFIX Or Style, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, .ttParhWnd, 0&, .hInstance, 0&)
                TThDC = GetDC(TThWnd)

                SetWindowPos TThWnd, SWP_TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS
                With Screen
                    SendMessage TThWnd, TTM_SETMAXTIPWIDTH, 0, ByVal .Width / .TwipsPerPixelX / 3
                End With 'SCREEN
                SendMessage TThWnd, TTM_ADDTOOL, 0&, ToolInfo
                SendMessage TThWnd, TTM_SETTITLE, Icon, ByVal myTitle
                SendMessage TThWnd, TTM_SETTIPTEXTCOLOR, myForeColor, ByVal 0&
                SendMessage TThWnd, TTM_SETTIPBKCOLOR, myBackColor, ByVal 0&
                SendMessage TThWnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal myHoverTime
                SendMessage TThWnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal myPopupTime
                SubstituteFont 'to fill myFontName, myFontSize, myFontBold, and myFontItalic
                Create = TThWnd 'return the tooltip window handle
            End If
        On Error GoTo 0
    End With 'TOOLINFO

End Function

Public Property Get ForeCol() As Long

  'this returns the current tooltip forecolor

    ForeCol = myForeColor

End Property

Public Property Get HoverTime() As Long

  'this returns the current mouse hover time time in millicecs (-1 for default)

    HoverTime = myHoverTime

End Property

Public Property Get Icon() As TTIcon

  'this returns the current tooltip icon

    Icon = myIcon

End Property

Public Property Get InitialText() As String

  'this returns the inital tooltip text, ie the one that was supplied on creation

    InitialText = myInitialText

End Property

Public Property Get InitialTitle() As String

  'this returns the inital tooltip title, ie the one that was supplied on creation

    InitialTitle = myInitialTitle

End Property

Private Sub KillFont()

    If TThFont Then
        DeleteObject TThFont
        TThFont = 0
    End If

End Sub

Public Property Get PopupTime() As Long

  'this returns the current max PopupTime time in millisecs (-1 for default)

    PopupTime = myPopupTime

End Property

Public Property Get Style() As TTStyle

  'this returns the current tooltip style

    Style = myStyle

End Property

Public Function SubstituteFont(Optional FontName As String = defFontName, _
                                                             Optional ByVal FontSize As Long = defFontSize, _
                                                             Optional ByVal Bold As Boolean = False, _
                                                             Optional ByVal Italic As Boolean = False) As Long

  'modify tooltip font

    If TThWnd Then 'we have a tooltip window
        Select Case FontSize 'limit fontsize to reasonable values
          Case Is < 8
            FontSize = 8
          Case Is > 36
            FontSize = 36
        End Select
        myFontName = Trim$(FontName)
        myFontSize = FontSize
        myFontBold = Bold
        myFontItalic = Italic
        KillFont 'kill any previous font and create a new one
        TThFont = CreateFont(-MulDiv(myFontSize, GetDeviceCaps(TThDC, LOGPIXELSY), 72), 0, 0, 0, IIf(myFontBold, fwBold, fwDontCare), myFontItalic, 0, 0, DEFAULT_CHARSET, 0, 0, 0, 0, myFontName)
        If TThFont Then                                 'font was successfully created...
            SendMessage TThWnd, WM_SETFONT, TThFont, 0  'so send it to the tooltip window...
            SubstituteFont = TThFont                    'and return font handle
        End If
    End If

End Function

Public Property Get Text() As String

  'this returns the current tooltip text

    Text = ToolInfo.pText

End Property

Public Property Get Title() As String

  'this returns the current tooltip Title

    Title = myTitle

End Property

':) Ulli's VB Code Formatter V2.24.21 (2009-Apr-01 09:11)  Decl: 117  Code: 235  Total: 352 Lines
':) CommentOnly: 23 (6,5%)  Commented: 51 (14,5%)  Filled: 278 (79%)  Empty: 74 (21%)  Max Logic Depth: 4
