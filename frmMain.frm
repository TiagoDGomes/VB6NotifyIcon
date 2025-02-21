VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NotifyIcon Demo"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSpecialOffer 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   330
      Top             =   2460
   End
   Begin Demo.NotifyIcon NotifyIcon 
      Left            =   3765
      Top             =   2475
      _ExtentX        =   847
      _ExtentY        =   900
      Icon            =   "frmMain.frx":08CA
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmMain.frx":11A4
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuSys 
      Caption         =   "System"
      Visible         =   0   'False
      Begin VB.Menu mnuSysUseFormIcon 
         Caption         =   "Use Form Icon"
      End
      Begin VB.Menu mnuSysRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuSysDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSysExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnSpecialOffer As Boolean

Private Sub Restore()
    tmrSpecialOffer.Enabled = False
    NotifyIcon.Restore HideIcon:=True
End Sub

Private Sub Form_Load()
    NotifyIcon.ToolTip = Caption
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then
        NotifyIcon.MinimizeToTray Me
        If NotifyIcon.BalloonShow(App.EXEName, _
                                  "I am in the tray!", _
                                  NIIF_WARNING Or NIIF_LARGE_ICON) Then
            tmrSpecialOffer.Enabled = True
            blnSpecialOffer = False
        Else
            MsgBox "Sorry, running on Windows older than" & vbNewLine _
                 & "Windows 200 so no baloon tips and" & vbNewLine _
                 & "no ""Special offer"" from this demo."
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NotifyIcon.Hide 'Cleans up TrayIcon and Subclasser.
End Sub

Private Sub mnuSysExit_Click()
    Unload Me
End Sub

Private Sub mnuSysRestore_Click()
    Restore
End Sub

Private Sub mnuSysUseFormIcon_Click()
    'Just an example of a ContextMenu action.
    Set NotifyIcon.Icon = Icon
    mnuSysUseFormIcon.Enabled = False
End Sub

Private Sub NotifyIcon_Activate()
    Restore
End Sub

Private Sub NotifyIcon_BalloonClick()
    If blnSpecialOffer Then
        'Just another demonstration.
        tmrSpecialOffer.Enabled = False
        NotifyIcon.SetForeground Me
        MsgBox "Congratulations!" & vbNewLine & vbNewLine _
             & "You have accepted our special offer.", _
               vbOKOnly Or vbExclamation, _
               App.EXEName
        blnSpecialOffer = False
    End If
End Sub

Private Sub NotifyIcon_BalloonDismissed()
    If blnSpecialOffer Then
        NotifyIcon.ToolTip = "Offer timed out or rejected"
        tmrSpecialOffer.Enabled = True
    End If
End Sub

Private Sub NotifyIcon_ContextMenu()
    NotifyIcon.SetForeground Me
    PopupMenu mnuSys
End Sub

Private Sub tmrSpecialOffer_Timer()
    'Just another demonstration.
    tmrSpecialOffer.Enabled = False
    blnSpecialOffer = True
    NotifyIcon.BalloonShow "Special Offer", _
                            "Click within the text to accept our special offer!", _
                            NIIF_USER, _
                            Icon
End Sub
