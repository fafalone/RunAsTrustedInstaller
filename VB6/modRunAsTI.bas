Attribute VB_Name = "modRunAsTI"
Option Explicit

'************************************************************************
'modRunAsTI - Run As TrustedInstaller
'Version: 2.0 (Feb 24 2022)
'Author: Jon Johnson (fafalone)   (see project thread for full credits)
'Project thread: https://www.vbforums.com/showthread.php?895287
'
'This module allows starting a process running as NT AUTHORITY\SYSTEM,
'giving it full system privileges. You must first be a normal admin.
'
'
'Main function:
'
'LaunchAsTI - The main function to call with the command line to launch
'             as impersonating TrustedInstaller.
'
'             This function will first enable the SeDebugPrivilege and
'             SeImpersonatePrivilege, then have the app impersonate the
'             system via winlogon.exe's token. It then starts the
'             TrustedInstaller service, impersonates it's thread, and
'             opens it's token. This should only need to be run once.
'
'             From there, it duplicates the token and starts a process
'             with it, which will run as SYSTEM.
'
'Other public functions:
'
'SetPrivilege - For convenience, this is public. It will attempt to
'               enable any of the privileges, all of which are defined
'               below.
'
'ReleaseToken - Calls CloseHandle on the TrustedInstaller token. If you
'               call this, the initialization procedure will run again
'               the next time LaunchAsTI is called.
'
'
'Version 2.0 changes:
'
'-Added support for command line arguments.
'
'-Error messages now have their descriptions looked up.
'
'-Replaced the 3-second wait for the TrustedInstaller service to start
' with continuous monitoring with system-suggested wait. This will avoid
' false errors on a busy system or if e.g. a hard drive spinup pauses
' the service launch.
' This also saved us from having to search for the process id as the
' QueryServiceStatusEx call returns that information.
'
'************************************************************************

Private hTiToken As Long
Private hAppThread As Long
Private sDesktop As String
Private bInit As Boolean
Private hNtDll As Long

'////////////////////////////////////////////////////////////////////////
'//                                                                    //
'//               NT Defined Privileges                                //
'//                                                                    //
'////////////////////////////////////////////////////////////////////////
Private Const SE_CREATE_TOKEN_NAME              As String = "SeCreateTokenPrivilege"
Private Const SE_ASSIGNPRIMARYTOKEN_NAME        As String = "SeAssignPrimaryTokenPrivilege"
Private Const SE_LOCK_MEMORY_NAME               As String = "SeLockMemoryPrivilege"
Private Const SE_INCREASE_QUOTA_NAME            As String = "SeIncreaseQuotaPrivilege"
Private Const SE_UNSOLICITED_INPUT_NAME         As String = "SeUnsolicitedInputPrivilege"
Private Const SE_MACHINE_ACCOUNT_NAME           As String = "SeMachineAccountPrivilege"
Private Const SE_TCB_NAME                       As String = "SeTcbPrivilege"
Private Const SE_SECURITY_NAME                  As String = "SeSecurityPrivilege"
Private Const SE_TAKE_OWNERSHIP_NAME            As String = "SeTakeOwnershipPrivilege"
Private Const SE_LOAD_DRIVER_NAME               As String = "SeLoadDriverPrivilege"
Private Const SE_SYSTEM_PROFILE_NAME            As String = "SeSystemProfilePrivilege"
Private Const SE_SYSTEMTIME_NAME                As String = "SeSystemtimePrivilege"
Private Const SE_PROF_SINGLE_PROCESS_NAME       As String = "SeProfileSingleProcessPrivilege"
Private Const SE_INC_BASE_PRIORITY_NAME         As String = "SeIncreaseBasePriorityPrivilege"
Private Const SE_CREATE_PAGEFILE_NAME           As String = "SeCreatePagefilePrivilege"
Private Const SE_CREATE_PERMANENT_NAME          As String = "SeCreatePermanentPrivilege"
Private Const SE_BACKUP_NAME                    As String = "SeBackupPrivilege"
Private Const SE_RESTORE_NAME                   As String = "SeRestorePrivilege"
Private Const SE_SHUTDOWN_NAME                  As String = "SeShutdownPrivilege"
Private Const SE_DEBUG_NAME                     As String = "SeDebugPrivilege"
Private Const SE_AUDIT_NAME                     As String = "SeAuditPrivilege"
Private Const SE_SYSTEM_ENVIRONMENT_NAME        As String = "SeSystemEnvironmentPrivilege"
Private Const SE_CHANGE_NOTIFY_NAME             As String = "SeChangeNotifyPrivilege"
Private Const SE_REMOTE_SHUTDOWN_NAME           As String = "SeRemoteShutdownPrivilege"
Private Const SE_UNDOCK_NAME                    As String = "SeUndockPrivilege"
Private Const SE_SYNC_AGENT_NAME                As String = "SeSyncAgentPrivilege"
Private Const SE_ENABLE_DELEGATION_NAME         As String = "SeEnableDelegationPrivilege"
Private Const SE_MANAGE_VOLUME_NAME             As String = "SeManageVolumePrivilege"
Private Const SE_IMPERSONATE_NAME               As String = "SeImpersonatePrivilege"
Private Const SE_CREATE_GLOBAL_NAME             As String = "SeCreateGlobalPrivilege"
Private Const SE_TRUSTED_CREDMAN_ACCESS_NAME    As String = "SeTrustedCredManAccessPrivilege"
Private Const SE_RELABEL_NAME                   As String = "SeRelabelPrivilege"
Private Const SE_INC_WORKING_SET_NAME           As String = "SeIncreaseWorkingSetPrivilege"
Private Const SE_TIME_ZONE_NAME                 As String = "SeTimeZonePrivilege"
Private Const SE_CREATE_SYMBOLIC_LINK_NAME      As String = "SeCreateSymbolicLinkPrivilege"


Private Const READ_CONTROL As Long = &H20000
Private Const MAXIMUM_ALLOWED = &H2000000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL As Long = &HFFFF

Private Const TOKEN_ASSIGN_PRIMARY As Long = &H1
Private Const TOKEN_DUPLICATE As Long = &H2
Private Const TOKEN_IMPERSONATE As Long = &H4
Private Const TOKEN_QUERY As Long = &H8
Private Const TOKEN_QUERY_SOURCE As Long = &H10
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_ADJUST_GROUPS As Long = &H40
Private Const TOKEN_ADJUST_DEFAULT As Long = &H80
Private Const TOKEN_ADJUST_SESSIONID As Long = &H100
Private Const TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)

Private Const THREAD_DIRECT_IMPERSONATION = (&H200)
Private Const ERROR_SERVICE_ALREADY_RUNNING = &H420
Private Const S_OK = 0&
Private Const STATUS_SUCCESS = 0
Private Const ERROR_FILE_NOT_FOUND = &H80070002
Private Const LOGON_WITH_PROFILE = &H1
Private Const CREATE_UNICODE_ENVIRONMENT = &H400
Private Const SYNCHRONIZE As Long = &H100000
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const ERROR_MR_MID_NOT_FOUND = 317&

Private Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF&)
Private Const PROCESS_CREATE_THREAD = &H2   ' Enables using the process handle in the CreateRemoteThread function to create a thread in the process.
Private Const PROCESS_DUP_HANDLE = &H40   ' Enables using the process handle as either the source or target process in the DuplicateHandle function to duplicate a handle
Private Const PROCESS_QUERY_INFORMATION = &H400 ' Enables using the process handle in the GetExitCodeProcess and GetPriorityClass functions to read information from the process object.
Private Const PROCESS_SET_INFORMATION = &H200 ' Enables using the process handle in the SetPriorityClass function to set the priority class of the process.
Private Const PROCESS_TERMINATE = &H1 ' Enables using the process handle in the TerminateProcess function to terminate the process.
Private Const PROCESS_VM_OPERATION = &H8 ' Enables using the process handle in the VirtualProtectEx and WriteProcessMemory functions to modify the virtual memory of the process.
Private Const PROCESS_VM_READ = &H10     ' Enables using the process handle in the ReadProcessMemory function to read from the virtual memory of the process.
Private Const PROCESS_VM_WRITE = &H20 ' Enables using the process handle in the WriteProcessMemory function to write to the virtual memory of the process.
    
Private Enum SE_PRIVILEGE_ATTRIBUTES
'The attributes of a privilege can be a combination of the following values.
    SE_PRIVILEGE_ENABLED = &H2&                 'The privilege is enabled.
    SE_PRIVILEGE_ENABLED_BY_DEFAULT = &H1&      'The privilege is enabled by default.
    SE_PRIVILEGE_REMOVED = &H4&                 'Used to remove a privilege. For details, see AdjustTokenPrivileges.
    SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000   'The privilege was used to gain access to an object or service.
                                                ' This flag is used to identify the relevant privileges in a set passed by a client application
                                                ' that may contain unnecessary privileges.
                                                'PrivilegeCheck sets the Attributes member of each LUID_AND_ATTRIBUTES structure to
                                                ' SE_PRIVILEGE_USED_FOR_ACCESS if the corresponding privilege is enabled.
End Enum

Private Type LUID
    lowPart As Long
    highPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid       As LUID
    Attributes  As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount      As Long
    Privileges(0 To 1)  As LUID_AND_ATTRIBUTES
End Type

Private Enum TOKEN_TYPE
    TokenPrimary = 1
    TokenImpersonation
End Enum

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type
Public Enum ShowWindowTypes
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
End Enum
Private Enum STARTUP_FLAGS
    STARTF_USESHOWWINDOW = &H1
    STARTF_USESIZE = &H2
    STARTF_USEPOSITION = &H4
    STARTF_USECOUNTCHARS = &H8
    STARTF_USEFILLATTRIBUTE = &H10
    STARTF_RUNFULLSCREEN = &H20            ' ignored For non-x86 platforms
    STARTF_FORCEONFEEDBACK = &H40
    STARTF_FORCEOFFFEEDBACK = &H80
    STARTF_USESTDHANDLES = &H100
    STARTF_USEHOTKEY = &H200
    STARTF_TITLEISLINKNAME = &H800
    STARTF_TITLEISAPPID = &H1000
    STARTF_PREVENTPINNING = &H2000
    STARTF_UNTRUSTEDSOURCE = &H8000
End Enum
Private Type STARTUPINFOW
   cbSize As Long
   lpReserved As Long
   lpDesktop As Long
   lpTitle As Long
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As STARTUP_FLAGS
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type SERVICE_STATUS_PROCESS
    dwServiceType As ServiceType
    dwCurrentState As ServiceState
    dwControlsAccepted As ServiceControlAccepted
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
    dwProcessId As Long
    dwServiceFlags As Long
End Type

Private Enum ServiceState
    SERVICE_STOPPED = &H1
    SERVICE_START_PENDING = &H2
    SERVICE_STOP_PENDING = &H3
    SERVICE_RUNNING = &H4
    SERVICE_CONTINUE_PENDING = &H5
    SERVICE_PAUSE_PENDING = &H6
    SERVICE_PAUSED = &H7
    SERVICE_NO_CHANGE = &HFFFFFFFF
End Enum
Private Enum ServiceType
    SERVICE_KERNEL_DRIVER = &H1
    SERVICE_FILE_SYSTEM_DRIVER = &H2
    SERVICE_WIN32_OWN_PROCESS = &H10
    SERVICE_WIN32_SHARE_PROCESS = &H20
    SERVICE_INTERACTIVE_PROCESS = &H100
    SERVICETYPE_NO_CHANGE = SERVICE_NO_CHANGE
End Enum
Private Enum ServiceControlAccepted
    SERVICE_ACCEPT_STOP = &H1
    SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
    SERVICE_ACCEPT_SHUTDOWN = &H4
    SERVICE_ACCEPT_PARAMCHANGE = &H8
    SERVICE_ACCEPT_NETBINDCHANGE = &H10
    SERVICE_ACCEPT_HARDWAREPROFILECHANGE = &H20
    SERVICE_ACCEPT_POWEREVENT = &H40
    SERVICE_ACCEPT_SESSIONCHANGE = &H80
    SERVICE_ACCEPT_PRESHUTDOWN = &H100
End Enum
Private Enum ServiceControlManagerType
    SC_MANAGER_CONNECT = &H1
    SC_MANAGER_CREATE_SERVICE = &H2
    SC_MANAGER_ENUMERATE_SERVICE = &H4
    SC_MANAGER_LOCK = &H8
    SC_MANAGER_QUERY_LOCK_STATUS = &H10
    SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
    SC_MANAGER_ALL_ACCESS = &HF003F
End Enum
Private Enum ACCESS_TYPE
    SERVICE_QUERY_CONFIG = &H1
    SERVICE_CHANGE_CONFIG = &H2
    SERVICE_QUERY_STATUS = &H4
    SERVICE_ENUMERATE_DEPENDENTS = &H8
    SERVICE_START = &H10
    SERVICE_STOP = &H20
    SERVICE_PAUSE_CONTINUE = &H40
    SERVICE_INTERROGATE = &H80
    SERVICE_USER_DEFINED_CONTROL = &H100
    SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED + SERVICE_QUERY_CONFIG + SERVICE_CHANGE_CONFIG + SERVICE_QUERY_STATUS + SERVICE_ENUMERATE_DEPENDENTS + SERVICE_START + SERVICE_STOP + SERVICE_PAUSE_CONTINUE + SERVICE_INTERROGATE + SERVICE_USER_DEFINED_CONTROL)
End Enum
Private Const SC_STATUS_PROCESS_INFO = 0&
Private Const SERVICE_RUNS_IN_SYSTEM_PROCESS = &H1

Private Type ENUM_SERVICE_STATUS_PROCESS
    lpServiceName As Long
    lpDisplayName As Long
    ServiceStatus As SERVICE_STATUS_PROCESS
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type SECURITY_QUALITY_OF_SERVICE
    Length As Long
    ImpersonationLevel As SECURITY_IMPERSONATION_LEVEL
    ContextTrackingMode As Byte
    EffectiveOnly As Byte
End Type

Private Enum SECURITY_IMPERSONATION_LEVEL
    SecurityAnonymous = 0
    SecurityIdentification = 1
    SecurityImpersonation = 2
    SecurityDelegation = 3
End Enum

Private Enum TH32CS_Flags
    TH32CS_SNAPHEAPLIST = &H1
    TH32CS_SNAPPROCESS = &H2
    TH32CS_SNAPTHREAD = &H4
    TH32CS_SNAPMODULE = &H8
    TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
    TH32CS_INHERIT = &H80000000
End Enum

Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
                        
Private Declare Function CreateProcessWithTokenW Lib "advapi32" (ByVal hToken As Long, ByVal dwLogonFlags As Long, _
    ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFOW, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As Any, ByRef ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValueW Lib "advapi32.dll" (ByVal StrPtrSystemName As Long, ByVal StrPtrName As Long, lpLuid As LUID) As Long
Private Declare Function LookupPrivilegeNameW Lib "advapi32.dll" (ByVal StrPtrSystemName As Long, lpLuid As LUID, ByVal StrPtrName As Long, cbName As Long) As Long
Private Declare Function OpenSCManagerW Lib "advapi32" (ByVal lpMachineName As Long, ByVal lpDatabaseName As Long, ByVal dwDesiredAccess As ServiceControlManagerType) As Long
Private Declare Function OpenServiceW Lib "advapi32" (ByVal hSCManager As Long, ByVal lpServiceName As Long, ByVal dwDesiredAccess As ACCESS_TYPE) As Long
Private Declare Function StartServiceW Lib "advapi32" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As TH32CS_Flags, ByVal th32ProcessID As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function NtImpersonateThread Lib "ntdll" (ByVal hThread As Long, ByVal hThreadToImpersonate As Long, SecurityQualityOfService As SECURITY_QUALITY_OF_SERVICE) As Long
Private Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal hThread As Long, ByVal dwDesiredAccess As Long, ByVal bOpenAsSelf As Long, phToken As Long) As Boolean
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function DuplicateTokenEx Lib "advapi32.dll" (ByVal hExistingToken As Long, ByVal dwDesiredAccess As Long, ByVal lpTokenAttributes As Long, ByVal ImpersonationLevel As SECURITY_IMPERSONATION_LEVEL, ByVal TokenType As TOKEN_TYPE, phNewToken As Long) As Boolean
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function Thread32First Lib "kernel32" (ByVal hSnapshot As Long, lpTE As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "kernel32" (ByVal hSnapshot As Long, lpTE As THREADENTRY32) As Long
Private Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ImpersonateLoggedOnUser Lib "advapi32" (ByVal hToken As Long) As Long
Private Declare Function QueryServiceStatusEx Lib "advapi32.dll" (ByVal hService As Long, ByVal InfoLevel As Long, lpBuffer As SERVICE_STATUS_PROCESS, ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare Function FormatMessageW Lib "kernel32" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal StrPtr As Long, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Function PathGetArgsW Lib "shlwapi" (ByVal pszPath As Long) As Long
Private Declare Sub PathRemoveArgsW Lib "shlwapi" (ByVal pszPath As Long)
Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

'***********************************************************
'IMPORTANT - THIS IS FOR OUR DEMO PROJECT
'If using this module separately, replace with Debug.Print
'or with output to your own logger.
Private Sub PostLog(smsg As String)
Form1.AppendLog smsg
End Sub
'***********************************************************


Public Function LaunchAsTI(sCommandLine As String) As Long
On Error GoTo e0
If bInit = False Then
    PostLog "Enabling privileges..."
    AdjustPrivileges
    PostLog "Impersonating system..."
    If ImpersonateSystem() = False Then
        PostLog "Failed to impersonate system."
        Exit Function
    End If
    bInit = True
End If

If hTiToken = 0& Then
    StartAndAcquireToken
End If

If hTiToken = 0 Then
    PostLog "Token hijack failed :("
    Exit Function
End If

PostLog "Duplicating stolen TI token..."
Dim lRet As Long
Dim lastErr As Long
Dim satr As SECURITY_ATTRIBUTES
Dim hStolenToken As Long
satr.nLength = Len(satr)
lRet = 0&
lRet = DuplicateTokenEx(hTiToken, MAXIMUM_ALLOWED, VarPtr(satr), SecurityImpersonation, TokenImpersonation, hStolenToken)
lastErr = Err.LastDllError
If lRet Then
    lRet = 0&: lastErr = 0&
    Dim tStartInfo As STARTUPINFOW
    Dim tProcInfo As PROCESS_INFORMATION
    
    sDesktop = "WinSta0\Default"
    tStartInfo.cbSize = Len(tStartInfo)
    tStartInfo.lpDesktop = StrPtr(sDesktop)
     
    PostLog "Token duplicated. Creating process..."
    Dim sArg As String
    Dim sTx As String
    Dim lpArg As Long
    
    sTx = sCommandLine
    
    lpArg = PathGetArgsW(StrPtr(sTx))
    sArg = LPWSTRtoStr(lpArg)
    If Len(sArg) > 0& Then
        sTx = Left$(sTx, Len(sTx) - Len(sArg))
    End If
    sTx = Trim$(sTx)
    If Left$(sTx, 1) = Chr$(34) Then
        sTx = Mid$(sTx, 2)
        sTx = Left$(sTx, Len(sTx) - 1)
    End If
    If sArg = "" Then
        LaunchAsTI = CreateProcessWithTokenW(hStolenToken, LOGON_WITH_PROFILE, 0&, StrPtr(sCommandLine), CREATE_UNICODE_ENVIRONMENT, 0&, 0&, tStartInfo, tProcInfo)
    Else
        PostLog "Command line args detected, parsed as:"
        PostLog "  App=" & sTx
        PostLog "  Arg=" & sArg
        LaunchAsTI = CreateProcessWithTokenW(hStolenToken, LOGON_WITH_PROFILE, StrPtr(sTx), StrPtr(sCommandLine), CREATE_UNICODE_ENVIRONMENT, 0&, 0&, tStartInfo, tProcInfo)
    End If
    lastErr = Err.LastDllError
    If LaunchAsTI = 0& Then
        PostLog "LaunchAsTI::CreateProcessWithTokenW failed, lastErr=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
    End If
Else
    PostLog "LaunchAsTI::Failed to duplicate TI token, lastErr=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
End If
Exit Function
e0:
    PostLog "VB Error 0x" & Hex$(Err.Number) & ", " & Err.Description
End Function

Private Sub AdjustPrivileges()
Dim hToken As Long
Dim lRet As Long
lRet = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken)
If lRet Then
    PostLog "AdjustPrivileges::Got process token."
    
    If SetPrivilege(hToken, SE_DEBUG_NAME, True) Then
        PostLog "AdjustPrivileges::Enabled debug privilege."
    Else
        PostLog "AdjustPrivileges::Failed to enable debug privilege."
    End If
    If SetPrivilege(hToken, SE_IMPERSONATE_NAME, True) Then
        PostLog "AdjustPrivileges::Enabled impersonate privilege."
    Else
        PostLog "AdjustPrivileges::Failed to enable impersonate privilege."
    End If
    
    CloseHandle hToken
Else
    PostLog "AdjustPrivileges::Failed to open process token."
End If
End Sub

Private Function ImpersonateSystem() As Boolean
Dim lRet As Long
Dim lastErr As Long
Dim hDupToken As Long
Dim hSysTkn As Long

Dim hWinLogon As Long
Dim pidWinLogon As Long
pidWinLogon = FindProcessByName("winlogon.exe")
If pidWinLogon Then
    PostLog "Got winlogon pid, opening process..."
    hWinLogon = OpenProcess(PROCESS_DUP_HANDLE Or PROCESS_QUERY_INFORMATION, 0&, pidWinLogon)
    lastErr = Err.LastDllError
    If hWinLogon Then
        lastErr = 0&
        PostLog "Got winlogon process handle, opening token..."
        lRet = OpenProcessToken(hWinLogon, TOKEN_QUERY Or TOKEN_DUPLICATE, hSysTkn)
        lastErr = Err.LastDllError
        If lRet Then
            lRet = 0&: lastErr = 0&
            lRet = ImpersonateLoggedOnUser(hSysTkn)
            lastErr = Err.LastDllError
            If lRet Then
                PostLog "Successfully impersonated system!"
                ImpersonateSystem = True
            Else
                PostLog "Failed to impersonate system. lastErr=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
            End If
            CloseHandle hDupToken
            CloseHandle hSysTkn
        Else
            PostLog "Failed to open winlogon process token. lastErr=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
        End If
        CloseHandle hWinLogon
    Else
        PostLog "Failed to open winlogon process, lastErr=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
    End If
Else
    PostLog "Failed to find winlogon processid"
End If
           
End Function

Private Function FindProcessByName(sName As String) As Long
Dim hSnapshot As Long
Dim tProcess As PROCESSENTRY32
Dim hr As Long
hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
If hSnapshot Then
    tProcess.dwSize = Len(tProcess)
    hr = Process32First(hSnapshot, tProcess)
    If hr > 0& Then
        Do While hr > 0&
            If LCase$(sName) = LCase$(Left$(tProcess.szExeFile, IIf(InStr(1, tProcess.szExeFile, Chr$(0)) > 0, InStr(1, tProcess.szExeFile, Chr$(0)) - 1, 0))) Then
                FindProcessByName = tProcess.th32ProcessID
                CloseHandle hSnapshot
                Exit Function
            End If
            hr = Process32Next(hSnapshot, tProcess)
        Loop
    Else
        PostLog "FindProcessByName->Process32First failed."
    End If
    CloseHandle hSnapshot
Else
    PostLog "FindProcessByName->Failed to create snapshot."
End If
End Function

Private Function GetFirstThreadId(pid As Long) As Long
Dim te32 As THREADENTRY32
Dim hSnapshot As Long
Dim hr As Long
hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0&)
If hSnapshot Then
    te32.dwSize = Len(te32)
    hr = Thread32First(hSnapshot, te32)
    Do
        If te32.th32OwnerProcessID = pid Then
            GetFirstThreadId = te32.th32ThreadID
            Exit Function
        End If
    Loop While Thread32Next(hSnapshot, te32)
End If

End Function

Private Function StartAndAcquireToken() As Long
'Start TrustedInstaller and yoink it's token
Dim hSCM As Long
Dim hSvc As Long
Dim hToken As Long
Dim lPid As Long
Dim lTiPid As Long
Dim hThread As Long
Dim hTiTid As Long
Dim hr As Long
Dim lastErr As Long
Dim status As Long
Dim lRet As Long

hSCM = OpenSCManagerW(0&, 0&, SC_MANAGER_ALL_ACCESS)
lastErr = Err.LastDllError
If hSCM = 0& Then
    PostLog "Failed to open SCManager, error=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
    Exit Function
End If
PostLog "Service manager opened. Opening TrustedInstaller service..."
lastErr = 0&
hSvc = OpenServiceW(hSCM, StrPtr("TrustedInstaller"), SERVICE_START Or SERVICE_QUERY_STATUS)
lastErr = Err.LastDllError
If hSvc Then
    PostLog "Attempting to start TrustedInstaller service..."
    Dim tStatus As SERVICE_STATUS_PROCESS
    Dim dwBytes As Long
    Do While QueryServiceStatusEx(hSvc, SC_STATUS_PROCESS_INFO, tStatus, Len(tStatus), dwBytes)
        If tStatus.dwCurrentState = SERVICE_STOPPED Then
            lastErr = 0&
            PostLog "Service currently stopped, starting..."
            hr = StartServiceW(hSvc, 0&, 0&)
            lastErr = Err.LastDllError
            If (hr = 0&) Then
                If lastErr <> ERROR_SERVICE_ALREADY_RUNNING Then
                    PostLog "Error starting TrustedInstaller service, error=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
                    CloseHandle hSvc
                    CloseHandle hSCM
                    Exit Function
                End If
            End If
        ElseIf (tStatus.dwCurrentState = SERVICE_START_PENDING) Or (tStatus.dwCurrentState = SERVICE_STOP_PENDING) Then
            PostLog "Service start pending, waiting " & tStatus.dwWaitHint
            Sleep tStatus.dwWaitHint
        ElseIf tStatus.dwCurrentState = SERVICE_RUNNING Then
            PostLog "Service running, pid=" & tStatus.dwProcessId
            lTiPid = tStatus.dwProcessId
            Exit Do
        End If
    Loop
    
    If lTiPid > 0& Then
        hTiTid = GetFirstThreadId(lTiPid)
        PostLog "First thread id for pid=" & hTiTid
        If hTiTid Then
            lastErr = 0&
            hThread = OpenThread(THREAD_DIRECT_IMPERSONATION, 0&, hTiTid)
            lastErr = Err.LastDllError
            If hThread Then
                Dim sqos As SECURITY_QUALITY_OF_SERVICE
                sqos.Length = Len(sqos)
                sqos.ImpersonationLevel = SecurityImpersonation
                status = NtImpersonateThread(GetCurrentThread(), hThread, sqos)
                If status = STATUS_SUCCESS Then
                    PostLog "NtImpersonateThread STATUS_SUCCESS. Opening current token..."
                    lastErr = 0&: lRet = 0&
                    lRet = OpenThreadToken(GetCurrentThread(), TOKEN_ALL_ACCESS, 0&, hTiToken)
                    lastErr = Err.LastDllError
                    If lRet Then
                        PostLog "OpenThreadToken success, return=0x" & lRet
                        lastErr = 0&: lRet = 0&
                    Else
                        PostLog "Failed to open own token after NtIT, lastErr=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
                    End If
                Else
                    PostLog "NtImpersonateThread failed, NTSTATUS=" & GetNtStatusName(status) & "(0x" & Hex$(status) & ")"
                End If
            Else
                PostLog "Failed to open TrustedInstaller thread, lastErr=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
            End If
        Else
            PostLog "Failed to get TrustedInstaller thread id: 0x" & Hex$(hTiTid)
        End If
    Else
        PostLog "Failed to find TrustedInstaller process, code 0x" & lTiPid
    End If
Else
    PostLog "Failed to open TrustedInstaller service handle, error=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
End If
CloseHandle hSvc
CloseHandle hSCM
End Function
Public Sub ReleaseToken()
CloseHandle hTiToken
hTiToken = 0&
End Sub

Public Function SetPrivilege(hToken As Long, ByVal strPrivilege As String, ByVal booEnable As Boolean) As Boolean
Dim tLUID As LUID
Dim tTP As TOKEN_PRIVILEGES
Dim tTP_Prev As TOKEN_PRIVILEGES
Dim lngReturnLength As Long
Dim lRet As Long
Dim lastErr As Long
SetPrivilege = False

If LookupPrivilegeValueW(0&, StrPtr(strPrivilege), tLUID) = 0 Then
  lastErr = Err.LastDllError
  PostLog "SetPrivilege::LookupPrivilegeValue failed. LastDllError=" & GetErrorName(lastErr) & " (0x" & Hex$(lastErr) & ")"
  Exit Function
End If

With tTP
  .PrivilegeCount = 1
  .Privileges(0).pLuid = tLUID
  If booEnable Then
    .Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
  Else
    .Privileges(0).Attributes = 0
  End If
End With

lRet = AdjustTokenPrivileges(hToken, False, tTP, Len(tTP_Prev), tTP_Prev, lngReturnLength)
lastErr = Err.LastDllError
If lastErr = 0 Then
  SetPrivilege = True
Else
  PostLog "SetPrivilege::Error code=" & GetErrorName(lastErr) & " (0x" & Hex$(Err.LastDllError) & "), return=0x" & Hex$(lRet)
End If
End Function

Private Function LPWSTRtoStr(lptr As Long, Optional ByVal fFree As Boolean = True) As String
SysReAllocString VarPtr(LPWSTRtoStr), lptr
If fFree Then
    Call CoTaskMemFree(lptr)
End If
End Function

Private Function GetErrorName(LastDllError As Long) As String
Dim sErr As String
Dim lRet As Long
sErr = Space$(1024)
lRet = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
            ByVal 0&, LastDllError, 0&, ByVal StrPtr(sErr), Len(sErr), ByVal 0&)
If lRet Then
    GetErrorName = Left$(sErr, lRet)
    If InStr(GetErrorName, vbCrLf) > 0 Then
        GetErrorName = Left$(GetErrorName, InStr(GetErrorName, vbCrLf) - 1)
    End If
End If

End Function
Private Function GetNtStatusName(nt As Long) As String
Dim sErr As String
Dim lRet As Long

If hNtDll = 0& Then
    hNtDll = LoadLibraryW(StrPtr("ntdll.dll"))
End If
If hNtDll Then
sErr = Space$(1024)
    lRet = FormatMessageW(FORMAT_MESSAGE_FROM_HMODULE Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                ByVal hNtDll, nt, 0&, ByVal StrPtr(sErr), Len(sErr), ByVal 0&)
    If lRet Then
        GetNtStatusName = Left$(sErr, lRet)
        If InStr(GetNtStatusName, vbCrLf) > 0 Then
            GetNtStatusName = Left$(GetNtStatusName, InStr(GetNtStatusName, vbCrLf) - 1)
        End If
    End If
End If
End Function


