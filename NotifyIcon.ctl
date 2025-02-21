VERSION 5.00
Begin VB.UserControl NotifyIcon 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00C0C0C0&
   PaletteMode     =   4  'None
   Picture         =   "NotifyIcon.ctx":0000
   ScaleHeight     =   480
   ScaleMode       =   0  'User
   ScaleWidth      =   451.765
   ToolboxBitmap   =   "NotifyIcon.ctx":00F1
End
Attribute VB_Name = "NotifyIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========
'NotifyIcon
'==========
'
'Implements a single Notification Area ("tray") icon raising events
'for common actions.  Supports Balloon Info messages in Windows 2000
'and later, but does not implement full WinXP or Vista and later
'behaviors and features and has reduced functionality prior to
'Windows 2000.
'
'Uses Subclasser module for safer subclassing.
'
'This is VERSION 5 of the NotifyIcon.ctl UserControl.
'
'*************
'NOW IDE-SAFE!
'*************
'
'   The Stop button should not crash the IDE.
'
'
'PROPERTIES
'
'   Icon As StdPicture (R/W)
'
'       Used to set the Notification Area icon.  If changed at runtime
'       the new icon is displayed replacing the old one if .Shown is
'       True.  Should be an .ico icon.  CAN BE SET DURING DESIGN MODE.
'
'   Shown As Boolean (RO)
'
'       True when Show has been called but Hide has not yet been called
'       (i.e. when the icon is active).
'
'   ToolTip As String (R/W)
'
'       Text for the Notification Area icon's tooltip (max 127 chars
'       starting in Windows 2000 or 63 if earlier).
'       CAN BE SET DURING DESIGN MODE.
'
'METHODS
'
'   BalloonHide()
'
'       Hides any displayed Balloon Info.
'
'   BalloonShow(ByVal Title As String, _
'               ByVal Text As String, _
'               Optional ByVal BalloonIcon As BalloonIcons = NIIF_NONE, _
'               Optional ByVal UserIcon As StdPicture) As Boolean
'
'       Shows Balloon Info with Title (max 63 chars) and Text (max 255
'       chars) and an optional standard or user icon.  User icon
'       (BalloonIcon = NIIF_USER) will be the Tray Icon image if an
'       alternate is not provided.  Only shows if Shown is True.
'
'       No balloon at all on Windows prior to Windows 2000, returns
'       False.  Windows 2000 and later return True.
'
'   Hide()
'
'       Hides/removes Notification Area icon, clears subclassing.
'
'   MinimizeToTray(ByVal Form As Form)
'
'       Minimize Form and hide Taskbar Icon, show Tray Icon.
'
'   Restore(Optional ByVal HideIcon As Boolean = False)
'
'       Undo MinimizeToTray() call, optionally hiding Tray Icon.
'
'   SetForeground(ByVal Form As Form)
'
'       Brings the Form to the top and gives it focus.
'
'   Show()
'
'       Sets up subclassing, adds/shows Notification Area icon.
'
'EVENTS
'
'   Activate()
'
'       Left double-click, Enter key.  Normally used to mean "restore
'       Form."
'
'   BalloonClick() [Requires XP or later.]
'
'       Raised when user clicks on balloon (not its close box).
'
'   BalloonDismissed() [Requires XP or later.]
'
'       Raised when balloon times out or when user clicks balloon
'       close box.
'
'   Click()
'
'       User left-clicked icon, rarely required.
'
'   ContextMenu()
'
'       Right-click or Menu key, normally used to mean "popup menu."
'
'
'ENUMS
'
'   BalloonIcons
'
'       Constants for standard Balloon icon and user-defined icon
'       choices.  NIIF_NOSOUND can be Or'ed with other values to
'       suppress sound on XP or earlier.   NIIF_LARGE_ICON can be
'       Or'ed to select a large icon on Vista or later.
'

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Private Const NOTIFYICONDATA_V1_SIZE_A = 88 'ANSI V1 size.

'Private Type NOTIFYICONDATA_V1
'    cbSize As Long
'    hWnd As Long
'    uId As Long
'    uFlags As Long
'    uCallBackMessage As Long
'    hIcon As Long
'    szTip(63) As Byte
'End Type

Private Const API_FALSE As Long = 0

Private Const NIM_ADD As Long = &H0&
Private Const NIM_MODIFY As Long = &H1&
Private Const NIM_DELETE As Long = &H2&
Private Const NIF_MESSAGE As Long = &H1&
Private Const NIF_ICON As Long = &H2&
Private Const NIF_TIP As Long = &H4&
Private Const NIF_INFO As Long = &H10&

Private Const WM_CONTEXTMENU As Long = &H7B&
Private Const WM_USER As Long = &H400&
Private Const WM_APP As Long = &H8000&
Private Const WM_APP_NIF As Long = WM_APP + &H1FFF&
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206

Private Const NIN_SELECT As Long = WM_USER + 0&
Private Const NINF_KEY As Long = &H1&
Private Const NIN_KEYSELECT As Long = NIN_SELECT Or NINF_KEY
Private Const NIN_BALLOONTIMEOUT As Long = WM_USER + 4&
Private Const NIN_BALLOONUSERCLICK As Long = WM_USER + 5&

Private Declare Function DllGetVersion Lib "shell32" (ByVal pdvi As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" ( _
    ByVal hWnd As Long) As Long

Private Declare Function Shell_NotifyIconA Lib "shell32" ( _
    ByVal dwMessage As Long, _
    ByVal pnid As Long) As Long

Private Declare Function Shell_NotifyIconW Lib "shell32" ( _
    ByVal dwMessage As Long, _
    ByVal pnid As Long) As Long

Private nid As NOTIFYICONDATA
Private mpicIcon As StdPicture
Private mblnShown As Boolean
Private mstrToolTip As String
Private mfrmMinimizedToTray As Form

Public Enum BalloonIcons
    NIIF_NONE = 0&
    NIIF_INFO = 1&
    NIIF_WARNING = 2&
    NIIF_ERROR = 3&
    NIIF_USER = 4&
    NIIF_NOSOUND = &H10& 'Can be Or'ed to suppress sound on XP or earlier.
    NIIF_LARGE_ICON = &H20& 'Or'ed for a large Balloon icon in Vista and later.
End Enum

Public Event Activate()               'Often used to mean "restore window."
Attribute Activate.VB_Description = "Left double-click, Enter key.  Normally used to mean ""restore Form"""
Attribute Activate.VB_MemberFlags = "200"
Public Event BalloonClick()           'User clicked on balloon (not its close box).
Attribute BalloonClick.VB_Description = "Raised when user clicks on balloon (not its close box)"
Public Event BalloonDismissed()       'Balloon timed out or user clicked balloon close box.
Attribute BalloonDismissed.VB_Description = "Raised when balloon times out or when user clicks balloon close box"
Public Event Click()                  'User left-clicked icon, rarely required.
Attribute Click.VB_Description = "User left-clicked icon, rarely required"
Public Event ContextMenu()            'Often used to mean "show popup menu."
Attribute ContextMenu.VB_Description = "Right-click or Menu key, normally used to mean ""popup menu"""

Public Property Get Icon() As StdPicture
Attribute Icon.VB_Description = "Icon (.ico) image displayed by the NotifyIcon"
Attribute Icon.VB_ProcData.VB_Invoke_Property = "StandardPicture"
    Set Icon = mpicIcon
End Property

Public Property Set Icon(ByVal Icon As StdPicture)
    Set mpicIcon = Icon
    If mblnShown Then
        nid.hIcon = mpicIcon.Handle
        Shell_NotifyIcon NIM_MODIFY, VarPtr(nid)
    End If
    PropertyChanged "Icon"
End Property

Public Property Get Shown() As Boolean
Attribute Shown.VB_Description = "True when the NotifyIcon is visible"
    Shown = mblnShown
End Property

Public Property Get ToolTip() As String
Attribute ToolTip.VB_Description = "Tooltip displayed when the user hovers the mouse over the NotifyIcon."
    If nid.cbSize = NOTIFYICONDATA_V1_SIZE_A Then
        ToolTip = Left$(mstrToolTip, 63)
    Else
        ToolTip = Left$(mstrToolTip, 127)
    End If
End Property

Public Property Let ToolTip(ByVal ToolTip As String)
    If nid.cbSize = NOTIFYICONDATA_V1_SIZE_A Then
        mstrToolTip = Left$(ToolTip, 63)
    Else
        mstrToolTip = Left$(ToolTip, 127)
    End If
    If mblnShown Then
        With nid
            .szTip = mstrToolTip & vbNullChar
        End With
        Shell_NotifyIcon NIM_MODIFY, VarPtr(nid)
    End If
    PropertyChanged "ToolTip"
End Property

Public Sub BalloonHide()
Attribute BalloonHide.VB_Description = "Remove the current Balloon Info if any"
    If nid.cbSize = NOTIFYICONDATA_V1_SIZE_A Then Exit Sub
    
    If mblnShown Then
        Shell_NotifyIcon NIM_MODIFY, VarPtr(nid)
    End If
End Sub

Public Function BalloonShow( _
    ByVal Title As String, _
    ByVal Text As String, _
    Optional ByVal BalloonIcon As BalloonIcons = NIIF_NONE, _
    Optional ByVal UserIcon As StdPicture) As Boolean
    'Returns False if OS is earlier than Windows 2000.
    
    If nid.cbSize = NOTIFYICONDATA_V1_SIZE_A Then Exit Function
    BalloonShow = True
    
    If mblnShown Then
        With nid
            .szInfoTitle = Left$(Title, 63) & vbNullChar
            .szInfo = Left$(Text, 255) & vbNullChar
            .dwInfoFlags = BalloonIcon
            If BalloonIcon = NIIF_USER Then nid.hIcon = UserIcon.Handle
            Shell_NotifyIcon NIM_MODIFY, VarPtr(nid)
            If BalloonIcon = NIIF_USER Then nid.hIcon = mpicIcon.Handle
            .szInfo = vbNullChar
        End With
    End If
End Function

Public Sub Hide()
Attribute Hide.VB_Description = "Removes the Notification Area Icon"
    If mblnShown Then
        Shell_NotifyIcon NIM_DELETE, VarPtr(nid)
        RemoveMe UserControl.hWnd, Me
        mblnShown = False
    End If
End Sub

Public Sub MinimizeToTray(ByVal Form As Form)
Attribute MinimizeToTray.VB_Description = "Minimize Form and hide Taskbar Icon, show Tray Icon"
    Set mfrmMinimizedToTray = Form
    Show
    Form.Hide
End Sub

Public Sub Restore(Optional ByVal HideIcon As Boolean = False)
Attribute Restore.VB_Description = "Undo MinimizeToTray() call, optionally hiding Tray Icon"
    If Not mfrmMinimizedToTray Is Nothing Then
        If HideIcon Then Hide
        mfrmMinimizedToTray.WindowState = vbNormal
        SetForegroundWindow mfrmMinimizedToTray.hWnd
        mfrmMinimizedToTray.Show
        Set mfrmMinimizedToTray = Nothing
    End If
End Sub

Public Sub SetForeground(ByVal Form As Form)
Attribute SetForeground.VB_Description = "Brings the Form to the top and gives it focus"
    SetForegroundWindow Form.hWnd
End Sub

Public Sub Show()
    SubclassMe UserControl.hWnd, Me
    With nid
        .hWnd = UserControl.hWnd
        .uId = UserControl.hWnd
        If .cbSize = NOTIFYICONDATA_V1_SIZE_A Then
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        Else
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO
        End If
        .hIcon = mpicIcon.Handle
        .uCallBackMessage = WM_APP_NIF
        .szTip = mstrToolTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, VarPtr(nid)
    mblnShown = True
End Sub

Private Function Shell_NotifyIcon( _
    ByVal dwMessage As Long, _
    ByVal pnid As Long) As Boolean
    
    With nid
        If .cbSize = NOTIFYICONDATA_V1_SIZE_A Then
            .szTip = StrConv(.szTip, vbFromUnicode)
            Shell_NotifyIcon = Shell_NotifyIconA(dwMessage, pnid) <> API_FALSE
        Else
            Shell_NotifyIcon = Shell_NotifyIconW(dwMessage, pnid) <> API_FALSE
        End If
    End With
End Function

'We mark this hidden here.  Code in the container should not invoke it.
Public Function SubclassProc( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long, _
    ByVal dwRefData As Long) As Long
Attribute SubclassProc.VB_Description = "Use with a subclassing technique here, not for use by container code"
Attribute SubclassProc.VB_MemberFlags = "40"
    
    If uMsg = WM_APP_NIF Then
        Select Case lParam
            Case NIN_SELECT, NIN_KEYSELECT, WM_LBUTTONDBLCLK
                RaiseEvent Activate
            Case NIN_BALLOONTIMEOUT
                RaiseEvent BalloonDismissed
            Case NIN_BALLOONUSERCLICK
                RaiseEvent BalloonClick
            Case WM_LBUTTONUP
                RaiseEvent Click
            Case WM_CONTEXTMENU, WM_RBUTTONUP
                RaiseEvent ContextMenu
        End Select
        SubclassProc = 0
    Else
        SubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

Private Sub UserControl_Initialize()
    Dim dviShell32 As DLLVERSIONINFO
    
    With dviShell32
        .cbSize = LenB(dviShell32)
        DllGetVersion VarPtr(dviShell32)
        If PackVersion(.dwMajorVersion, .dwMinorVersion, .dwBuildNumber) < PackVersion(5, 0, 0) Then
            'Oops, before Win2K.
            nid.cbSize = NOTIFYICONDATA_V1_SIZE_A
        Else
            nid.cbSize = LenB(nid)
        End If
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set Icon = .ReadProperty("Icon", Nothing)
        ToolTip = .ReadProperty("ToolTip", "")
    End With
End Sub

Private Sub UserControl_Resize()
    Width = ScaleX(32, vbPixels, ScaleMode)
    Height = ScaleY(32, vbPixels, ScaleMode)
End Sub

Private Sub UserControl_Terminate()
    Hide
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Icon", Icon, Nothing
        .WriteProperty "ToolTip", ToolTip, ""
    End With
End Sub
