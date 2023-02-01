Attribute VB_Name = "mdlMain"
Option Explicit

'Registry Declarations
'=====================
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

'Misc Declarations
'=================
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const SYNCHRONIZE As Long = &H100000
Private Const INFINITE As Long = &HFFFFFFFF

'Shell Starter Declarations
'==========================
Public PATH As String, Mode_Edit As Boolean
Public SS_Executable_Path As String, SS_Executable_CL As String, SS_Option_Choice As Integer, SS_Option_Other_Command As String, SS_REG_KEY As Integer, SS_Splash_Timer As Integer, SS_ASK As Integer

Public Sub Main()
    PATH = App.PATH
    If Right$(PATH, 1) <> "\" Then PATH = PATH & "\"
    Mode_Edit = False
    If Dir(PATH & "Shell_Starter.ini") <> "" Then
        LoadSS_Settings (PATH & "Shell_Starter.ini")
        If SS_Splash_Timer > 0 Then frmWaitKey.Show vbModal
        If Mode_Edit = False Then RunApp SS_Executable_Path
        If Mode_Edit = False Then RunAfter
        SaveSS_Settings (PATH & "Shell_Starter.ini")
    Else
        Load frmMain
        frmMain.Show
        Mode_Edit = True
    End If
End Sub

Public Sub LoadSS_Settings(ByVal strSettingsPath As String)
Restart:
    If Dir(Trim$(strSettingsPath)) <> "" Then
        Dim FF As Integer, strTemp As String
        FF = FreeFile
        Open strSettingsPath For Input As #FF
            Do While Not EOF(FF)
                Line Input #FF, strTemp
                If Left$(strTemp, 15) = "ExecutablePath=" Then
                    SS_Executable_Path = Trim$(Mid$(strTemp, 16))
                    frmMain.txtExe.Text = SS_Executable_Path
                ElseIf Left$(strTemp, 13) = "ExecutableCL=" Then
                    SS_Executable_CL = Trim$(Mid$(strTemp, 14))
                    frmMain.txtParam.Text = SS_Executable_CL
                ElseIf Left$(strTemp, 13) = "OptionalMode=" Then
                    SS_Option_Choice = Int(Trim$(Mid$(strTemp, 14)))
                    frmMain.optEnd(SS_Option_Choice).Value = True
                ElseIf Left$(strTemp, 18) = "OptionalModeOther=" Then
                    SS_Option_Other_Command = Trim$(Mid$(strTemp, 19))
                    frmMain.txtOther.Text = SS_Option_Other_Command
                ElseIf Left$(strTemp, 8) = "REG_KEY=" Then
                    SS_REG_KEY = Int(Trim$(Mid$(strTemp, 9)))
                    frmMain.optREG(SS_REG_KEY).Value = True
                ElseIf Left$(strTemp, 11) = "SplashTime=" Then
                    SS_Splash_Timer = Int(Trim$(Mid$(strTemp, 12)))
                    frmMain.txtSplashTimer.Text = Trim$(Str$(SS_Splash_Timer))
                ElseIf Left$(strTemp, 4) = "ASK=" Then
                    SS_ASK = Int(Trim$(Mid$(strTemp, 5)))
                    frmMain.chkASK.Value = SS_ASK
                End If
            Loop
        Close #FF
    Else
        SaveSS_Settings (PATH & "Shell_Starter.ini")
        GoTo Restart
    End If
End Sub

Public Sub SaveSS_Settings(ByVal strSettingsPath As String)
    Dim FF As Integer, i As Integer
    FF = FreeFile
    Open strSettingsPath For Output As #FF
        SS_Executable_Path = frmMain.txtExe.Text
        Print #FF, "ExecutablePath=" & SS_Executable_Path
        SS_Executable_CL = frmMain.txtParam.Text
        Print #FF, "ExecutableCL=" & SS_Executable_CL
        For i = frmMain.optEnd.LBound To frmMain.optEnd.uBound
            If frmMain.optEnd.item(i).Value = True Then Exit For
        Next
        SS_Option_Choice = i
        Print #FF, "OptionalMode=" & Str$(SS_Option_Choice)
        SS_Option_Other_Command = frmMain.txtOther.Text
        Print #FF, "OptionalModeOther=" & SS_Option_Other_Command
        If frmMain.optREG(0).Value = True Then Print #FF, "REG_KEY=0" Else Print #FF, "REG_KEY=1"
        If IsNumeric(Trim$(frmMain.txtSplashTimer.Text)) = True Then
            SS_Splash_Timer = Int(Trim$(frmMain.txtSplashTimer.Text))
            If SS_Splash_Timer < 0 Then SS_Splash_Timer = 3
        Else
            SS_Splash_Timer = 3
        End If
        Print #FF, "SplashTime=" & Str$(SS_Splash_Timer)
        SS_ASK = frmMain.chkASK.Value
        Print #FF, "ASK=" & Str$(SS_ASK)
    Close #FF
End Sub

Public Sub RunApp(ByVal strExecutable As String)
    If Dir(Trim$(strExecutable)) <> "" And strExecutable <> "" Then
        ShellAndWait SS_Executable_Path, vbNormalFocus
    Else
        MsgBox "Can't find file!" & vbCrLf & "Loading Shell Start application!", vbExclamation + vbOKOnly, "Error..."
        Load frmMain
        frmMain.Show
        Mode_Edit = True
    End If
End Sub

Public Sub RunAfter()
    If frmMain.optEnd(0).Value = True Then
        If SS_ASK = 1 Then
            If MsgBox("Shutdown?", vbYesNo + vbExclamation, "What to do?") = vbYes Then Shell ("shutdown -s")
        Else
            Shell ("shutdown -s")
        End If
    ElseIf frmMain.optEnd(1).Value = True Then
        If SS_ASK = 1 Then
            If MsgBox("Restart?", vbYesNo + vbExclamation, "What to do?") = vbYes Then Shell ("shutdown -r")
        Else
            Shell ("shutdown -r")
        End If
    ElseIf frmMain.optEnd(2).Value = True Then
        If SS_ASK = 1 Then
            If MsgBox("Other?", vbYesNo + vbExclamation, "What to do?") = vbYes Then RunComKernel SS_Option_Other_Command, True
        Else
            RunComKernel SS_Option_Other_Command, True
        End If
    Else
        MsgBox "Error with mode selection!" & vbCrLf & "Starting default windows shell!", vbExclamation + vbOKOnly, "Error..."
        Shell "explorer.exe", vbNormalFocus
    End If
    End
End Sub

Public Function GetString(ByVal Hkey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim lValueType As Long
    Dim R As Long
    R = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveString(ByVal Hkey As HKeyTypes, ByVal strPath As String, ByVal strValue As String, ByVal strdata As String)
    Dim keyhand As Long
    Dim R As Long
    R = RegCreateKey(Hkey, strPath, keyhand)
    R = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    R = RegCloseKey(keyhand)
End Sub

Public Function GetCurrentShell_LOCAL_MACHINE()
    Dim strTemp As String
    strTemp = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell")
    GetCurrentShell_LOCAL_MACHINE = strTemp
End Function

Public Function GetCurrentShell_CURRENT_USER()
    Dim strTemp As String
    strTemp = GetString(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell")
    GetCurrentShell_CURRENT_USER = strTemp
End Function

Public Function SetCurrentShell_LOCAL_MACHINE_Default()
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"
End Function

Public Function SetCurrentShell_LOCAL_MACHINE(ByVal strPath As String)
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", strPath
End Function

Public Function SetCurrentShell_CURRENT_USER_Default()
    SaveString HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"
End Function

Public Function SetCurrentShell_CURRENT_USER(ByVal strPath As String)
    SaveString HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", strPath
End Function

Public Sub RunComKernel(ByVal ComCommand As String, Optional ByVal Continue As Boolean = False)
    If Continue = False Then
        Shell Environ("comspec") & " /c" & ComCommand, vbNormalFocus
    ElseIf Continue = True Then
        Shell Environ("comspec") & " /k" & ComCommand, vbNormalFocus
    End If
End Sub

Public Sub ErrorShow(ByVal ErrDesc As String, ByVal ErrNumber As Integer)
    MsgBox ErrDesc & " ( " & ErrNumber & " )", vbOKOnly + vbCritical, "Error"
End Sub

Public Sub ShellAndWait(ByVal Program_Name As String, ByVal Window_Style As VbAppWinStyle)
    Dim Process_ID As Long
    Dim Process_Handle As Long
    On Error GoTo ShellError
    Process_ID = Shell(Program_Name, Window_Style)
    On Error GoTo 0
    Process_Handle = OpenProcess(SYNCHRONIZE, 0, Process_ID)
    If Process_Handle <> 0 Then
        WaitForSingleObject Process_Handle, INFINITE
        CloseHandle Process_Handle
    End If
    Exit Sub
ShellError:
    ErrorShow Err.Description, Err.Number
End Sub


