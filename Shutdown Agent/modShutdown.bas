Attribute VB_Name = "modShutdown"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   This module provides actually more information about the operating   '
'  system than is needed to perform the shutdown action where the only   '
'  important thing to identify is the windows platform.                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

'Shutdown:
Public Enum eShutDownAction
    waLOGOFF = 0
    waPOWEROFF = &H8
    waSHUTDOWN = 1
    waREBOOT = 2
End Enum

'System Power State
Public Enum eSystemPowerState
    waSUSPEND
    waHIBERNATE
End Enum

'OS:
Private Enum eAllPlatforms
    wapWIN_32
    wapWIN_9x_ME
    wapWIN_NT
End Enum
Private Enum eAllSystems
    wasWIN_32
    wasWIN_95
    wasWIN_98
    wasWIN_ME
    wasWIN_NT
    wasWIN_2000
    wasWIN_XP
End Enum
Private Type tOS
    Platform As eAllPlatforms
    System As eAllSystems
    BuildNumber As Long
    WindowsVersion As String
End Type

'OS: Identify the Operating System|Platform|BuildNumber|WindowsVersion
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'Privilege Processing
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const TOKEN_QUERY = &H8
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLUID As LARGE_INTEGER
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges As LUID_AND_ATTRIBUTES
End Type

'Shutdown
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Const waFORCE = 4

'Suspend|Hibernate
Private Declare Function SetSuspendState Lib "Powrprof" (ByVal Hibernate As Long, ByVal ForceCritical As Long, ByVal DisableWakeEvent As Long) As Long

'Lock Computer
Private Declare Function LockWorkStation Lib "user32.dll" () As Long

'Change Privilege
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

Public Function ShutDown(mAction As eShutDownAction, mForceAppsClosed As Boolean) As Boolean
'Supports: ShutDown, Soft PowerOFF, Reboot, and LogOFF
'Platforms: ALL

    Dim flags As Long
    If mForceAppsClosed Then flags = (mAction Or waFORCE) Else flags = mAction
    
    If OS.Platform = wapWIN_NT Then
        EnablePrivilege True
        If ExitWindowsEx(flags, 0) Then ShutDown = True
        EnablePrivilege False
    Else
        If ExitWindowsEx(flags, 0) Then ShutDown = True
    End If

End Function

Public Function SetSystemPowerState(mAction As eSystemPowerState, mForceSuspension As Boolean, mDisableWakeEvents As Boolean) As Boolean
'Supports: Suspend(Stand By), Hibernate
'Platforms: Only Windows 98 or later, Windows 2000 or later
'If Hibernation is not enabled on the target system, it will Suspend instead

    Dim SYS As eAllSystems
    SYS = OS.System
    
    If Not (SYS = wasWIN_32 Or SYS = wasWIN_95 Or SYS = wasWIN_NT) Then
        If SetSuspendState(mAction, mForceSuspension, mDisableWakeEvents) Then SetSystemPowerState = True
    End If

End Function

Public Function LockComputer() As Boolean
'Platforms: Only Windows 2000 or later

    Dim SYS As eAllSystems
    SYS = OS.System
    
    If SYS = (wasWIN_2000) Or (SYS = wasWIN_XP) Then
        If LockWorkStation Then LockComputer = True
    End If

End Function

Public Property Get Platform() As String

    Select Case OS.Platform
        Case wapWIN_32
            Platform = "Windows 32"
        Case wapWIN_9x_ME
            Platform = "Windows 9x|ME"
        Case wapWIN_NT
            Platform = "Windows NT"
    End Select

End Property

Public Property Get System() As String

    Select Case OS.System
        Case wasWIN_32
            System = "Windows 32"
        Case wasWIN_95
            System = "Windows 95"
        Case wasWIN_98
            System = "Windows 98"
        Case wasWIN_ME
            System = "Windows ME"
        Case wasWIN_NT
            System = "Windows NT"
        Case wasWIN_2000
            System = "Windows 2000"
        Case wasWIN_XP
            System = "Windows XP"
    End Select

End Property

Public Property Get BuildNumber() As Long
    BuildNumber = OS.BuildNumber
End Property

Public Property Get WindowsVersion() As String
    WindowsVersion = OS.WindowsVersion
End Property

Private Function OS() As tOS
    
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    
    'Get Version information
    Call GetVersionExA(OSInfo)
    
    'Save Build Number
    OS.BuildNumber = OSInfo.dwBuildNumber
    
    'Save Windows Version
    OS.WindowsVersion = CStr(OSInfo.dwMajorVersion) & "." & CStr(OSInfo.dwMinorVersion)
    
    'Identify the Operating System
    With OSInfo
        Select Case .dwPlatformId
            Case 0
                OS.Platform = wapWIN_32
                OS.System = wasWIN_32
            Case 1
                OS.Platform = wapWIN_9x_ME
                Select Case .dwMinorVersion
                    Case 0
                        OS.System = wasWIN_95
                    Case 10
                        OS.System = wasWIN_98
                    Case 90
                        OS.System = wasWIN_ME
                End Select
            Case 2
                OS.Platform = wapWIN_NT
                Select Case .dwMajorVersion
                    Case Is < 5
                        OS.System = wasWIN_NT
                    Case 5 And .dwMinorVersion = 0
                        OS.System = wasWIN_2000
                    Case 5 And .dwMinorVersion = 1
                        OS.System = wasWIN_XP
                End Select
        End Select
    End With

End Function

Private Sub EnablePrivilege(ByVal State As Boolean)
'Only Windows NT or later

    Dim hProc As Long
    Dim OldTokenStuffLen As Long
    Dim NewTokenStuffLen As Long
    Dim OldTokenStuff As TOKEN_PRIVILEGES
    Dim NewTokenStuff As TOKEN_PRIVILEGES
    
    'Get Token Handle
    OpenProcessToken GetCurrentProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hProc
    'Get LUID representing the Shutdown Privilege Name
    LookupPrivilegeValue "", SE_SHUTDOWN_NAME, OldTokenStuff.Privileges.pLUID
    NewTokenStuff = OldTokenStuff
    NewTokenStuff.PrivilegeCount = 1
    NewTokenStuffLen = Len(NewTokenStuff)

    If State Then
        'Enable Shutdown Privilege
        NewTokenStuff.Privileges.Attributes = SE_PRIVILEGE_ENABLED
        AdjustTokenPrivileges hProc, False, NewTokenStuff, NewTokenStuffLen, OldTokenStuff, OldTokenStuffLen
    Else
        'Disable ShutDown Privilege
        NewTokenStuff.Privileges.Attributes = 0
        AdjustTokenPrivileges hProc, False, NewTokenStuff, NewTokenStuffLen, OldTokenStuff, OldTokenStuffLen
    End If

End Sub
