Attribute VB_Name = "Subclasser"
Option Explicit
'
'See: http://msdn.microsoft.com/en-us/library/bb762102(VS.85).aspx
'and related items.
'
'This code is based on Karl E. Peterson's: HookXP
'
'http://vb.mvps.org/samples/HookXP/
'
'However he will not support the modifications here.
'
'Additional changes to support NotifyIcon.ctl here.
'

Public Type DLLVERSIONINFO
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

'---- NotifyIcon.ctl support-----
Private Type NOTIFYICONDATA_V1
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_DELETE As Long = &H2

Private Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconW" ( _
    ByVal dwMessage As Long, _
    ByVal pnid As Long) As Boolean
'--------------------------------

Private Const WM_NCDESTROY As Long = &H82&
Private Const WM_UAHDESTROYWINDOW As Long = &H90& '??? Undocumented ???

Private Declare Function DllGetVersion Lib "comctl32" (ByVal pdvi As Long) As Long

Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" ( _
    ByVal hWnd As Long, _
    ByVal pfnSubclass As Long, _
    ByVal uIdSubclass As Long, _
    ByVal dwRefData As Long) As Long

Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" ( _
    ByVal hWnd As Long, _
    ByVal pfnSubclass As Long, _
    ByVal uIdSubclass As Long) As Long

Public Declare Function DefSubclassProc Lib "comctl32" Alias "#413" ( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Function PackVersion(ByVal Major As Long, ByVal Minor As Long, ByVal Build As Long) As Long
    'Allows for Major 0-127, Minor 0-512, Build 0-32767.
    PackVersion = Major * &H1000000 + Minor * &H8000& + Build
End Function

Public Function SubclassMe( _
    ByVal hWnd As Long, _
    ByVal pfnSubclass As Object, _
    Optional ByVal dwRefData As Long) As Boolean
    
    Dim dviComCtl32 As DLLVERSIONINFO
    
    With dviComCtl32
        .cbSize = LenB(dviComCtl32)
        DllGetVersion VarPtr(dviComCtl32)
        If PackVersion(.dwMajorVersion, .dwMinorVersion, .dwBuildNumber) < PackVersion(4, 72, 0) Then
            'Uh oh.  Win95 w/o the IE 4.x Integrated Shell.
            Err.Raise &H8004F000, "Subclasser.SubclassMe", "Requires COMCTL32.DLL 4.72 or later"
        End If
    End With
    
    SubclassMe = SetWindowSubclass(hWnd, _
                                   AddressOf SubclassProxy, _
                                   ObjPtr(pfnSubclass), _
                                   dwRefData)
End Function

Public Function RemoveMe( _
    ByVal hWnd As Long, _
    ByVal pfnSubclass As Object) As Boolean
    
    RemoveMe = RemoveWindowSubclass(hWnd, _
                                    AddressOf SubclassProxy, _
                                    ObjPtr(pfnSubclass))
End Function

Private Function SubclassProxy( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long, _
    ByVal uIdSubclass As Object, _
    ByVal dwRefData As Long) As Long
   
    If uMsg = WM_NCDESTROY Or uMsg = WM_UAHDESTROYWINDOW Then
        'Just in case the client fails to clean up.
        '---- NotifyIcon.ctl support-----
        Dim nid As NOTIFYICONDATA_V1
        
        With nid
            .cbSize = LenB(nid)
            .hWnd = hWnd
            .uId = hWnd
        End With
        Shell_NotifyIcon NIM_DELETE, VarPtr(nid)
        '--------------------------------
        RemoveMe hWnd, uIdSubclass
    Else
        SubclassProxy = uIdSubclass.SubclassProc(hWnd, uMsg, wParam, lParam, dwRefData)
    End If
End Function
