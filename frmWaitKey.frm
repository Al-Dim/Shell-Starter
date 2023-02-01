VERSION 5.00
Begin VB.Form frmWaitKey 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Wait for key..."
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1980
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   480
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Press < SPACE > key, if you want to show application window !!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Label lblTimer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   60
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmWaitKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CurrentTime As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        Me.Hide
        Load frmMain
        frmMain.Show
        Mode_Edit = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    CurrentTime = SS_Splash_Timer
    lblTimer.Caption = Trim$(Str$(CurrentTime))
End Sub

Private Sub Timer1_Timer()
    CurrentTime = CurrentTime - 1
    lblTimer.Caption = Trim$(Str$(CurrentTime))
    If CurrentTime <= 0 Then
        Unload Me
    End If
End Sub
