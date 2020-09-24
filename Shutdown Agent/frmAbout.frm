VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   2175
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblEMail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "E-Mail ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3420
      TabIndex        =   2
      ToolTipText     =   "wa_aly@tdcspace.dk"
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblWebSite 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "WebSite:   http://ebrain.8m.net/"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   2580
      TabIndex        =   1
      Top             =   1380
      Width           =   2595
   End
   Begin VB.Label lblME 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "By:"
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   2100
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    lblME = "By: Waleed A. Aly, PTS, Egypt." & vbCrLf & vbCrLf & _
        "If you like this program, you are morally obliged to send me an eMail."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEMail.Font.Underline = False
    lblWebSite.Font.Underline = False
End Sub

Private Sub lblEMail_Click()
    ShellExecute Me.hWnd, "open", "mailto:wa_aly@tdcspace.dk?subject=Shutdown Agent", vbNullString, "C:\", 5
End Sub

Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEMail.Font.Underline = True
End Sub

Private Sub lblME_Click()
    Unload Me
End Sub

Private Sub lblWebSite_Click()
    ShellExecute Me.hWnd, "open", "http://ebrain.8m.net/", vbNullString, "C:\", 5
End Sub

Private Sub lblWebSite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblWebSite.Font.Underline = True
End Sub
