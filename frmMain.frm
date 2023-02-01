VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shell Starter v.0.2"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Optional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   8055
      Begin VB.CheckBox chkASK 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Ask before execute commands."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CommandButton cmdExplorer 
         Caption         =   "Start Explorer ( Default Shell )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   7815
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Text            =   "shutdown /s /t 0"
         Top             =   1320
         Width           =   5895
      End
      Begin VB.OptionButton optEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Other"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton optEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Restart PC"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Shutdown PC"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.TextBox txtSplashTimer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Text            =   "3"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.OptionButton optREG 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "CURRENT_USER"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   15
         Top             =   3000
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optREG 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "LOCAL_MACHINE"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   14
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmdSetCurrentShell 
         Caption         =   "Set Current Shell"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   3855
      End
      Begin VB.CommandButton cmdSetCurrentShellDefault 
         Caption         =   "Set Current Shell To Default"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   3855
      End
      Begin VB.TextBox txtParam 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   10
         Top             =   1560
         Width           =   7575
      End
      Begin VB.CommandButton cmdExe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Browse"
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtExe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         Top             =   600
         Width           =   7575
      End
      Begin VB.Label lblSplash 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Splash Timer"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblParam 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CL Text"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblExe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Executable Path"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExe_Click()
    Dim strFileName As String
    Dim CDlg As New cFileDialog
    On Error GoTo Error_Exit
    With CDlg
        .CancelError = True
        .InitDir = PATH
        .flags = DialogFlags.OFN_FILEMUSTEXIST + DialogFlags.OFN_EXPLORER
        .Filename = ""
        .Filter = "Executable Files(*.exe)|*.exe|All Files(*.*)|*.*|"
        .ShowOpen
        strFileName = .Filename
    End With
    On Error GoTo 0
        txtExe.Text = strFileName
        Set CDlg = Nothing
    Exit Sub
Error_Exit:
    Set CDlg = Nothing
End Sub

Private Sub cmdExplorer_Click()
    Shell "explorer.exe", vbNormalFocus
End Sub

Private Sub cmdSetCurrentShell_Click()
    If optREG(0).Value = True Then
        If MsgBox("Set LOCAL MACHINE SHELL?", vbYesNo + vbExclamation, "What to do?") = vbYes Then SetCurrentShell_LOCAL_MACHINE PATH & App.EXEName & ".exe"
    Else
        If MsgBox("Set CURRENT USER SHELL?", vbYesNo + vbExclamation, "What to do?") = vbYes Then SetCurrentShell_CURRENT_USER PATH & App.EXEName & ".exe"
    End If
    optREG(0).ToolTipText = GetCurrentShell_LOCAL_MACHINE
    optREG(1).ToolTipText = GetCurrentShell_CURRENT_USER
End Sub

Private Sub cmdSetCurrentShellDefault_Click()
    If optREG(0).Value = True Then
        If MsgBox("Set LOCAL MACHINE SHELL TO DEFAULT?", vbYesNo + vbExclamation, "What to do?") = vbYes Then SetCurrentShell_LOCAL_MACHINE_Default
    Else
        If MsgBox("Set CURRENT USER SHELL TO DEFAULT?", vbYesNo + vbExclamation, "What to do?") = vbYes Then SetCurrentShell_CURRENT_USER_Default
    End If
    optREG(0).ToolTipText = GetCurrentShell_LOCAL_MACHINE
    optREG(1).ToolTipText = GetCurrentShell_CURRENT_USER
End Sub

Private Sub Form_Load()
    'optREG(0).ToolTipText = GetCurrentShell_LOCAL_MACHINE
    'optREG(1).ToolTipText = GetCurrentShell_CURRENT_USER
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSS_Settings PATH & "Shell_Starter.ini"
End Sub

Private Sub txtExe_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    txtExe.Text = Data.Files(1)
    On Error GoTo 0
End Sub

