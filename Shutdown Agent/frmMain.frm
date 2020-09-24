VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shutdown Agent"
   ClientHeight    =   4335
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3780
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3120
      Top             =   3000
   End
   Begin VB.CommandButton cmdTimer 
      Caption         =   "&Enable Timer"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   3660
      Width           =   3435
   End
   Begin VB.Frame frame 
      Caption         =   " Configuration "
      Height          =   3315
      Left            =   180
      TabIndex        =   8
      Top             =   180
      Width           =   3435
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   150
         ScaleHeight     =   825
         ScaleWidth      =   3075
         TabIndex        =   10
         Top             =   930
         Width           =   3075
         Begin VB.TextBox txtPeriod 
            Height          =   315
            Left            =   750
            MaxLength       =   3
            TabIndex        =   20
            Text            =   "1"
            Top             =   0
            Width           =   495
         End
         Begin VB.ComboBox cmbUnit 
            Height          =   315
            ItemData        =   "frmMain.frx":08CA
            Left            =   1530
            List            =   "frmMain.frx":08D7
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   0
            Width           =   1515
         End
         Begin VB.OptionButton opt 
            Caption         =   "In"
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   18
            Top             =   60
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton opt 
            Caption         =   "At"
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   17
            Top             =   540
            Width           =   675
         End
         Begin VB.VScrollBar vsPeriod 
            Height          =   315
            Left            =   1230
            Max             =   1
            Min             =   999
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Value           =   1
            Width           =   165
         End
         Begin VB.VScrollBar vsHour 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1230
            Max             =   1
            Min             =   12
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   480
            Value           =   1
            Width           =   165
         End
         Begin VB.TextBox txtHour 
            Enabled         =   0   'False
            Height          =   315
            Left            =   750
            MaxLength       =   2
            TabIndex        =   14
            Text            =   "1"
            Top             =   480
            Width           =   495
         End
         Begin VB.VScrollBar vsMin 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2010
            Max             =   0
            Min             =   59
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   480
            Width           =   165
         End
         Begin VB.TextBox txtMin 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "00"
            Top             =   480
            Width           =   495
         End
         Begin VB.ComboBox cmbT 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":08F4
            Left            =   2310
            List            =   "frmMain.frx":08FE
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.CheckBox chkForce 
         Caption         =   "Force processes to terminate"
         Height          =   255
         Left            =   210
         TabIndex        =   1
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2355
      End
      Begin VB.VScrollBar vsBlink 
         Height          =   315
         Left            =   1380
         Max             =   1
         Min             =   999
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2760
         Value           =   30
         Width           =   165
      End
      Begin VB.TextBox txtBlink 
         Height          =   315
         Left            =   900
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "30"
         Top             =   2760
         Width           =   495
      End
      Begin VB.CheckBox chkBlink 
         Caption         =   "Icon blink before execution by"
         Height          =   255
         Left            =   210
         TabIndex        =   2
         Top             =   2340
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.ComboBox cmbAction 
         Height          =   315
         ItemData        =   "frmMain.frx":090A
         Left            =   900
         List            =   "frmMain.frx":091D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Seconds"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   2820
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Action"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   480
         Width           =   450
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Timer Disabled"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   4140
      Width           =   3795
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************'
'                                                       '
'   By:         Waleed A. Aly                           '
'   ASL:        [21 M Egypt]                            '
'   eMail:      wa_aly@tdcspace.dk                      '
'   On:         29 Jan, 2003                            '
'                                                       '
'     Please eMail me any Comments and|or Suggestions.  '
'   I hope you like my work and think is usefull !  :)  '
'   I'd love to know how many people are using my Code  '
'   so you can always eMail me if you are goin' to use  '
'   it :)                                               '
'                                      Thanks.          '
'                                                       '
'*******************************************************'

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Note that the picture control 'pic' is not actually used but it does  '
'  serve a purpose. It is used as a container for the radio buttons bec-  '
'  ause if they are directly placed within the frame they will be drawn   '
'  incorrectly when XP Style is applied: an ugly black box will show.     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private sAction As String
Private bEnabled As Boolean
Private bAllowExit As Boolean
Private bIconBlink As Boolean
Private bForceExecution As Boolean
Private SecondsToBlink As Long
Private ExecutionTime As Date
Private TrayIcon As NOTIFYICONDATA

Private Sub chkBlink_Click()
    txtBlink.Enabled = chkBlink.Value
    vsBlink.Enabled = chkBlink.Value
End Sub

Private Sub cmdTimer_Click()

    Dim sInterval As String
    
    If bEnabled Then
        lblStatus = sTD
        cmdTimer.Caption = sEnable
        TrayIcon.hIcon = Me.Icon.Handle
        TrayIcon.szTip = sNoAction & vbNullChar
        Shell_NotifyIcon NIM_MODIFY, TrayIcon
        frame.Enabled = True
    Else
        frame.Enabled = False
        cmdTimer.Caption = sDisable
        sAction = cmbAction
        bForceExecution = chkForce.Value
        bIconBlink = chkBlink.Value
        SecondsToBlink = Val(txtBlink)
        If opt(0).Value Then
            Select Case cmbUnit
                Case "Hours"
                    sInterval = "h"
                Case "Minutes"
                    sInterval = "n"
                Case "Seconds"
                    sInterval = "s"
            End Select
            ExecutionTime = DateAdd(sInterval, Val(txtPeriod), Now)
        Else
            ExecutionTime = Date & " " & txtHour & ":" & txtMin & cmbT
            If ExecutionTime < Now Then ExecutionTime = DateAdd("d", "1", ExecutionTime)
        End If
        TrayIcon.szTip = sAction & " On " & ExecutionTime & vbNullChar
        Shell_NotifyIcon NIM_MODIFY, TrayIcon
    End If
    
    bEnabled = Not bEnabled
    Timer.Enabled = bEnabled

End Sub

Private Sub Form_Initialize()

    InitCommonControls  'Add XP Style Support
    
    If App.PrevInstance Then
        MsgBox sAlreadyRunning, vbInformation, sBY
        End
    End If
    
    If System = "Windows 2000" Or System = "Windows XP" Then cmbAction.AddItem "Lock Computer"
    
    cmbAction = "Shutdown"
    cmbUnit = "Hours"
    cmbT = "AM"
    
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hIcon = Me.Icon.Handle
        .hWnd = Me.hWnd
        .szTip = sNoAction & vbNullChar
        .uCallBackMessage = WM_MOUSEMOVE
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uId = vbNull
    End With
    
    Shell_NotifyIcon NIM_ADD, TrayIcon

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim Message As Long
    
    Message = X / Screen.TwipsPerPixelX
    
    Select Case Message
        Case WM_LBUTTONDBLCLK
            If frmAbout.Visible Then Unload frmAbout
            Me.Show
        Case WM_RBUTTONUP
            SetForegroundWindow Me.hWnd
            PopupMenu mnuPopup
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If bAllowExit Then
        Shell_NotifyIcon NIM_DELETE, TrayIcon
        Exit Sub
    Else
        Cancel = True
        Me.Hide
    End If

End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
    bAllowExit = True
    Unload Me
End Sub

Private Sub opt_Click(Index As Integer)

    Dim b As Boolean
    b = Index And True
    txtHour.Enabled = b
    vsHour.Enabled = b
    txtMin.Enabled = b
    vsMin.Enabled = b
    cmbT.Enabled = b
    txtPeriod.Enabled = Not b
    vsPeriod.Enabled = Not b
    cmbUnit.Enabled = Not b

End Sub

Private Sub Timer_Timer()

    Static bShow As Boolean
    Dim RemainingSeconds As Long
    
    RemainingSeconds = DateDiff("s", Now, ExecutionTime)
    lblStatus = RemainingSeconds & sTE
    
    If bIconBlink And RemainingSeconds <= SecondsToBlink Then
        If bShow Then TrayIcon.hIcon = Me.Icon.Handle Else TrayIcon.hIcon = 0
        Shell_NotifyIcon NIM_MODIFY, TrayIcon
        bShow = Not bShow
    End If
    
    If ExecutionTime > Now Then Exit Sub
    
    Select Case sAction
        Case "Shutdown"
            ShutDown waPOWEROFF, bForceExecution
        Case "Restart"
            ShutDown waREBOOT, bForceExecution
        Case "Log Off"
            ShutDown waLOGOFF, bForceExecution
        Case "Suspend"
            SetSystemPowerState waSUSPEND, bForceExecution, False
        Case "Hibernate"
            SetSystemPowerState waHIBERNATE, bForceExecution, False
        Case "Lock Computer"
            LockComputer
    End Select
    
    bAllowExit = True
    Unload Me

End Sub

Private Sub txtBlink_Change()
    If Val(txtBlink) > vsBlink.Min Then txtBlink = vsBlink.Min
    If Val(txtBlink) < vsBlink.Max Then txtBlink = vsBlink.Max
    vsBlink.Value = Val(txtBlink)
End Sub

Private Sub txtBlink_KeyPress(KeyAscii As Integer)
    If IsDigit(KeyAscii) Then Exit Sub Else KeyAscii = 0
End Sub

Private Sub txtHour_Change()
    If Val(txtHour) > vsHour.Min Then txtHour = vsHour.Min
    If Val(txtHour) < vsHour.Max Then txtHour = vsHour.Max
    vsHour.Value = Val(txtHour)
End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)
    If IsDigit(KeyAscii) Then Exit Sub Else KeyAscii = 0
End Sub

Private Sub txtMin_Change()
    If Val(txtMin) > vsMin.Min Then txtMin = vsMin.Min
    If Val(txtMin) < vsMin.Max Then txtMin = vsMin.Max
    vsMin.Value = Val(txtMin)
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
    If IsDigit(KeyAscii) Then Exit Sub Else KeyAscii = 0
End Sub

Private Sub txtPeriod_Change()
    If Val(txtPeriod) > vsPeriod.Min Then txtPeriod = vsPeriod.Min
    If Val(txtPeriod) < vsPeriod.Max Then txtPeriod = vsPeriod.Max
    vsPeriod.Value = Val(txtPeriod)
End Sub

Private Sub txtPeriod_KeyPress(KeyAscii As Integer)
    If IsDigit(KeyAscii) Then Exit Sub Else KeyAscii = 0
End Sub

Private Sub vsBlink_Change()
    txtBlink = vsBlink.Value
End Sub

Private Sub vsHour_Change()
    txtHour = vsHour.Value
End Sub

Private Sub vsMin_Change()
    txtMin = Format(vsMin.Value, "00")
End Sub

Private Sub vsPeriod_Change()
    txtPeriod = vsPeriod.Value
End Sub

Private Function IsDigit(KeyAscii As Integer) As Boolean
    If KeyAscii > 47 And KeyAscii < 58 Then IsDigit = True
End Function
