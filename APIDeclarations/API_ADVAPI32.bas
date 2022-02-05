Attribute VB_Name = "API_ADVAPI32"

Type SID_IDENTIFIER_AUTHORITY
    Value(0 To 5) As Byte
End Type

Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type

Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

#If VBA7 And Win64 Then
    
    Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Integer
        Owner As LongPtr
        Group As LongPtr
        Sacl As ACL
        Dacl As ACL
    End Type

    Declare PtrSafe Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
    Declare PtrSafe Function AccessCheck Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As LongPtr, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, PrivilegeSet As PRIVILEGE_SET, PrivilegeSetLength As Long, GrantedAccess As Long, ByVal Status As LongPtr) As Long
    Declare PtrSafe Function AccessCheckAndAuditAlarm Lib "advapi32.dll" Alias "AccessCheckAndAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, SecurityDescriptor As SECURITY_DESCRIPTOR, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, ByVal ObjectCreation As Long, GrantedAccess As Long, ByVal AccessStatus As LongPtr, ByVal pfGenerateOnClose As LongPtr) As Long
    Declare PtrSafe Function AddAccessAllowedAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
    Declare PtrSafe Function AddAccessDeniedAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
    Declare PtrSafe Function AddAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, pAceList As Any, ByVal nAceListLength As Long) As Long
    Declare PtrSafe Function AddAuditAccessAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwAccessMask As Long, pSid As Any, ByVal bAuditSuccess As Long, ByVal bAuditFailure As Long) As Long
    Declare PtrSafe Function AdjustTokenGroups Lib "advapi32.dll" (ByVal TokenHandle As LongPtr, ByVal ResetToDefault As Long, NewState As TOKEN_GROUPS, ByVal BufferLength As Long, PreviousState As TOKEN_GROUPS, ReturnLength As Long) As Long
    Declare PtrSafe Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As LongPtr, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
    Declare PtrSafe Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As LongPtr) As Long
    Declare PtrSafe Function AllocateLocallyUniqueId Lib "advapi32.dll" (Luid As LARGE_INTEGER) As Long
    Declare PtrSafe Function AreAllAccessesGranted Lib "advapi32.dll" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
    Declare PtrSafe Function AreAnyAccessesGranted Lib "advapi32.dll" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
    Declare PtrSafe Function BackupEventLog Lib "advapi32.dll" Alias "BackupEventLogA" (ByVal hEventLog As LongPtr, ByVal lpBackupFileName As String) As Long
    Declare PtrSafe Function ClearEventLog Lib "advapi32.dll" Alias "ClearEventLogA" (ByVal hEventLog As LongPtr, ByVal lpBackupFileName As String) As Long
    Declare PtrSafe Function CloseEventLog Lib "advapi32.dll" (ByVal hEventLog As LongPtr) As Long
    Declare PtrSafe Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As LongPtr) As Long
    Declare PtrSafe Function ControlService Lib "advapi32.dll" (ByVal hService As LongPtr, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
    Declare PtrSafe Function CopySid Lib "advapi32.dll" (ByVal nDestinationSidLength As Long, pDestinationSid As Any, pSourceSid As Any) As Long
    Declare PtrSafe Function CreatePrivateObjectSecurity Lib "advapi32.dll" (ParentDescriptor As SECURITY_DESCRIPTOR, CreatorDescriptor As SECURITY_DESCRIPTOR, NewDescriptor As SECURITY_DESCRIPTOR, ByVal IsDirectoryObject As Long, ByVal Token As LongPtr, GenericMapping As GENERIC_MAPPING) As Long
    Declare PtrSafe Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As LongPtr, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As LongPtr
    Declare PtrSafe Function DeleteAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceIndex As Long) As Long
    Declare PtrSafe Function DeleteService Lib "advapi32.dll" (ByVal hService As LongPtr) As Long
    Declare PtrSafe Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As LongPtr) As Long
    Declare PtrSafe Function DestroyPrivateObjectSecurity Lib "advapi32.dll" (ObjectDescriptor As SECURITY_DESCRIPTOR) As Long
    Declare PtrSafe Function DuplicateToken Lib "advapi32.dll" (ByVal ExistingTokenHandle As LongPtr, ImpersonationLevel As Integer, DuplicateTokenHandle As LongPtr) As Long
    Declare PtrSafe Function EnumDependentServices Lib "advapi32.dll" Alias "EnumDependentServicesA" (ByVal hService As LongPtr, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long) As Long
    Declare PtrSafe Function EnumServicesStatus Lib "advapi32.dll" Alias "EnumServicesStatusA" (ByVal hSCManager As LongPtr, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
    Declare PtrSafe Function EqualPrefixSid Lib "advapi32.dll" (pSid1 As Any, pSid2 As Any) As Long
    Declare PtrSafe Function EqualSid Lib "advapi32.dll" (pSid1 As Any, pSid2 As Any) As Long
    Declare PtrSafe Function FindFirstFreeAce Lib "advapi32.dll" (pAcl As ACL, pAce As LongPtr) As Long
    Declare PtrSafe Sub FreeSid Lib "advapi32.dll" (pSid As Any)
    Declare PtrSafe Function GetAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceIndex As Long, pAce As Any) As Long
    Declare PtrSafe Function GetAclInformation Lib "advapi32.dll" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
    Declare PtrSafe Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
    Declare PtrSafe Function GetKernelObjectSecurity Lib "advapi32.dll" (ByVal Handle As LongPtr, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
    Declare PtrSafe Function GetLengthSid Lib "advapi32.dll" (pSid As Any) As Long
    Declare PtrSafe Function GetOldestEventLogRecord Lib "advapi32.dll" (ByVal hEventLog As LongPtr, OldestRecord As Long) As Long
    Declare PtrSafe Function GetPrivateObjectSecurity Lib "advapi32.dll" (ObjectDescriptor As SECURITY_DESCRIPTOR, ByVal SecurityInformation As Long, ResultantDescriptor As SECURITY_DESCRIPTOR, ByVal DescriptorLength As Long, ReturnLength As Long) As Long
    Declare PtrSafe Function GetSecurityDescriptorControl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pControl As Integer, lpdwRevision As Long) As Long
    Declare PtrSafe Function GetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, lpbDaclPresent As Long, pDacl As ACL, lpbDaclDefaulted As Long) As Long
    Declare PtrSafe Function GetSecurityDescriptorGroup Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal lpbGroupDefaulted As LongPtr) As Long
    Declare PtrSafe Function GetSecurityDescriptorLength Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
    Declare PtrSafe Function GetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal lpbOwnerDefaulted As LongPtr) As Long
    Declare PtrSafe Function GetSecurityDescriptorSacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal lpbSaclPresent As LongPtr, pSacl As ACL, ByVal lpbSaclDefaulted As LongPtr) As Long
    Declare PtrSafe Function GetServiceDisplayName Lib "advapi32.dll" Alias "GetServiceDisplayNameA" (ByVal hSCManager As LongPtr, ByVal lpServiceName As String, ByVal lpDisplayName As String, lpcchBuffer As Long) As Long
    Declare PtrSafe Function GetServiceKeyName Lib "advapi32.dll" Alias "GetServiceKeyNameA" (ByVal hSCManager As LongPtr, ByVal lpDisplayName As String, ByVal lpServiceName As String, lpcchBuffer As Long) As Long
    Declare PtrSafe Function GetSidIdentifierAuthority Lib "advapi32.dll" (pSid As Any) As SID_IDENTIFIER_AUTHORITY
    Declare PtrSafe Function GetSidLengthRequired Lib "advapi32.dll" (ByVal nSubAuthorityCount As Byte) As Long
    Declare PtrSafe Function GetSidSubAuthority Lib "advapi32.dll" (pSid As Any, ByVal nSubAuthority As Long) As LongPtr
    Declare PtrSafe Function GetSidSubAuthorityCount Lib "advapi32.dll" (pSid As Any) As LongPtr
    Declare PtrSafe Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As LongPtr, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
    Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare PtrSafe Function ImpersonateNamedPipeClient Lib "advapi32.dll" (ByVal hNamedPipe As LongPtr) As Long
    Declare PtrSafe Function ImpersonateSelf Lib "advapi32.dll" (ImpersonationLevel As Integer) As Long
    Declare PtrSafe Function InitializeAcl Lib "advapi32.dll" (pAcl As ACL, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
    Declare PtrSafe Function InitializeSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
    Declare PtrSafe Function InitializeSid Lib "advapi32.dll" (Sid As Any, pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte) As Long
    Declare PtrSafe Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
    Declare PtrSafe Function LockServiceDatabase Lib "advapi32.dll" (ByVal hSCManager As LongPtr) As LongPtr
    Declare PtrSafe Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, ByVal Sid As LongPtr, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
    Declare PtrSafe Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidW" (ByVal lpSystemName As Any, Sid As Any, Name As Any, cbName As Long, ReferencedDomainName As Any, cbReferencedDomainName As Long, peUse As Integer) As Long
    Declare PtrSafe Function LookupPrivilegeDisplayName Lib "advapi32.dll" Alias "LookupPrivilegeDisplayNameA" (ByVal lpSystemName As String, ByVal lpName As String, ByVal lpDisplayName As String, cbDisplayName As Long, lpLanguageID As Long) As Long
    Declare PtrSafe Function LookupPrivilegeName Lib "advapi32.dll" Alias "LookupPrivilegeNameA" (ByVal lpSystemName As String, lpLuid As LARGE_INTEGER, ByVal lpName As String, cbName As Long) As Long
    Declare PtrSafe Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
    Declare PtrSafe Function NotifyBootConfigStatus Lib "advapi32.dll" (ByVal BootAcceptable As Long) As Long
    Declare PtrSafe Function NotifyChangeEventLog Lib "advapi32" (ByVal hEventLog As LongPtr, ByVal hEvent As LongPtr) As Long
    Declare PtrSafe Function ObjectCloseAuditAlarm Lib "advapi32.dll" Alias "ObjectCloseAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal GenerateOnClose As Long) As Long
    Declare PtrSafe Function ObjectPrivilegeAuditAlarm Lib "advapi32.dll" Alias "ObjectPrivilegeAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ClientToken As LongPtr, ByVal DesiredAccess As Long, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
    Declare PtrSafe Function OpenBackupEventLog Lib "advapi32.dll" Alias "OpenBackupEventLogA" (ByVal lpUNCServerName As String, ByVal lpFileName As String) As LongPtr
    Declare PtrSafe Function OpenEventLog Lib "advapi32.dll" Alias "OpenEventLogA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As LongPtr
    Declare PtrSafe Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As LongPtr, ByVal DesiredAccess As Long, TokenHandle As LongPtr) As Long
    Declare PtrSafe Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As LongPtr
    Declare PtrSafe Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As LongPtr, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As LongPtr
    Declare PtrSafe Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As LongPtr, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As LongPtr) As Long
    Declare PtrSafe Function PrivilegeCheck Lib "advapi32.dll" (ByVal ClientToken As LongPtr, RequiredPrivileges As PRIVILEGE_SET, ByVal pfResult As LongPtr) As Long
    Declare PtrSafe Function PrivilegedServiceAuditAlarm Lib "advapi32.dll" Alias "PrivilegedServiceAuditAlarmA" (ByVal SubsystemName As String, ByVal ServiceName As String, ByVal ClientToken As LongPtr, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
    Declare PtrSafe Function ReadEventLog Lib "advapi32.dll" Alias "ReadEventLogA" (ByVal hEventLog As LongPtr, ByVal dwReadFlags As Long, ByVal dwRecordOffset As Long, lpBuffer As EVENTLOGRECORD, ByVal nNumberOfBytesToRead As Long, pnBytesRead As Long, pnMinNumberOfBytesNeeded As Long) As Long
    Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
    Declare Ptrsafe Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As LongPtr, phkResult As LongPtr) As Long
    Declare PtrSafe Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, phkResult As LongPtr) As Long
    Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As LongPtr, lpdwDisposition As Long) As Long
    Declare PtrSafe Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String) As Long
    Declare PtrSafe Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As LongPtr, ByVal lpValueName As String) As Long
    Declare PtrSafe Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
    Declare PtrSafe Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As LongPtr, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
    Declare PtrSafe Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As LongPtr, lpType As Long, lpData As Byte, lpcbData As Long) As Long
    Declare PtrSafe Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
    Declare PtrSafe Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
    Declare PtrSafe Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As LongPtr
    Declare PtrSafe Function RegisterServiceCtrlHandler Lib "advapi32.dll" Alias "RegisterServiceCtrlHandlerA" (ByVal lpServiceName As String, ByVal lpHandlerProc As LongPtr) As LongPtr
    Declare PtrSafe Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal lpFile As String) As Long
    Declare PtrSafe Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As LongPtr, ByVal fAsynchronus As Long) As Long
    Declare PtrSafe Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, phkResult As LongPtr) As Long
    Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As LongPtr) As Long
    Declare PtrSafe Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As LongPtr, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As LongPtr, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
    Declare PtrSafe Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
    Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    Declare PtrSafe Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
    Declare PtrSafe Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As LongPtr, ByVal lpFile As String, ByVal dwFlags As Long) As Long
    Declare PtrSafe Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As LongPtr, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
    Declare PtrSafe Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
    Declare PtrSafe Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
    Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    Declare PtrSafe Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As LongPtr, ByVal lpSubKey As String) As Long
    Declare PtrSafe Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As LongPtr, ByVal wType As Long, ByVal wCategory As Long, ByVal dwEventID As Long, lpUserSid As Any, ByVal wNumStrings As Long, ByVal dwDataSize As Long, ByVal lpStrings As LongPtr, lpRawData As Any) As Long
    Declare PtrSafe Function RevertToSelf Lib "advapi32.dll" () As Long
    Declare PtrSafe Function SetAclInformation Lib "advapi32.dll" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
    Declare PtrSafe Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
    Declare PtrSafe Function SetPrivateObjectSecurity Lib "advapi32.dll" (ByVal SecurityInformation As Long, ModificationDescriptor As SECURITY_DESCRIPTOR, ObjectsSecurityDescriptor As SECURITY_DESCRIPTOR, GenericMapping As GENERIC_MAPPING, ByVal Token As LongPtr) As Long
    Declare PtrSafe Function SetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, pDacl As ACL, ByVal bDaclDefaulted As Long) As Long
    Declare PtrSafe Function SetSecurityDescriptorGroup Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal bGroupDefaulted As Long) As Long
    Declare PtrSafe Function SetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal bOwnerDefaulted As Long) As Long
    Declare PtrSafe Function SetSecurityDescriptorSacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bSaclPresent As Long, pSacl As ACL, ByVal bSaclDefaulted As Long) As Long
    Declare PtrSafe Function SetServiceBits Lib "advapi32" (ByVal hServiceStatus As LongPtr, ByVal dwServiceBits As Long, ByVal bSetBitsOn As Long, ByVal bUpdateImmediately As Long) As Long
    Declare PtrSafe Function SetServiceObjectSecurity Lib "advapi32.dll" (ByVal hService As LongPtr, ByVal dwSecurityInformation As Long, lpSecurityDescriptor As Any) As Long
    Declare PtrSafe Function SetServiceStatus Lib "advapi32.dll" (ByVal hServiceStatus As LongPtr, lpServiceStatus As SERVICE_STATUS) As Long
    Declare PtrSafe Function SetThreadToken Lib "advapi32" (Thread As LongPtr, ByVal Token As LongPtr) As Long
    Declare PtrSafe Function SetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As LongPtr, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long) As Long
    Declare PtrSafe Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As LongPtr, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As LongPtr) As Long
    Declare PtrSafe Function StartServiceCtrlDispatcher Lib "advapi32.dll" Alias "StartServiceCtrlDispatcherA" (lpServiceStartTable As SERVICE_TABLE_ENTRY) As Long
    Declare PtrSafe Function UnlockServiceDatabase Lib "advapi32.dll" (ScLock As Any) As Long
    
#Else
    
    Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
    Declare Function AccessCheck Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As Long, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, PrivilegeSet As PRIVILEGE_SET, PrivilegeSetLength As Long, GrantedAccess As Long, ByVal Status As Long) As Long
    Declare Function AccessCheckAndAuditAlarm Lib "advapi32.dll" Alias "AccessCheckAndAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, SecurityDescriptor As SECURITY_DESCRIPTOR, ByVal DesiredAccess As Long, GenericMapping As GENERIC_MAPPING, ByVal ObjectCreation As Long, GrantedAccess As Long, ByVal AccessStatus As Long, ByVal pfGenerateOnClose As Long) As Long
    Declare Function AddAccessAllowedAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
    Declare Function AddAccessDeniedAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Any) As Long
    Declare Function AddAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, pAceList As Any, ByVal nAceListLength As Long) As Long
    Declare Function AddAuditAccessAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceRevision As Long, ByVal dwAccessMask As Long, pSid As Any, ByVal bAuditSuccess As Long, ByVal bAuditFailure As Long) As Long
    Declare Function AdjustTokenGroups Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal ResetToDefault As Long, NewState As TOKEN_GROUPS, ByVal BufferLength As Long, PreviousState As TOKEN_GROUPS, ReturnLength As Long) As Long
    Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
    Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
    Declare Function AllocateLocallyUniqueId Lib "advapi32.dll" (Luid As LARGE_INTEGER) As Long
    Declare Function AreAllAccessesGranted Lib "advapi32.dll" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
    Declare Function AreAnyAccessesGranted Lib "advapi32.dll" (ByVal GrantedAccess As Long, ByVal DesiredAccess As Long) As Long
    Declare Function BackupEventLog Lib "advapi32.dll" Alias "BackupEventLogA" (ByVal hEventLog As Long, ByVal lpBackupFileName As String) As Long
    Declare Function ClearEventLog Lib "advapi32.dll" Alias "ClearEventLogA" (ByVal hEventLog As Long, ByVal lpBackupFileName As String) As Long
    Declare Function CloseEventLog Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
    Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
    Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
    Declare Function CopySid Lib "advapi32.dll" (ByVal nDestinationSidLength As Long, pDestinationSid As Any, pSourceSid As Any) As Long
    Declare Function CreatePrivateObjectSecurity Lib "advapi32.dll" (ParentDescriptor As SECURITY_DESCRIPTOR, CreatorDescriptor As SECURITY_DESCRIPTOR, NewDescriptor As SECURITY_DESCRIPTOR, ByVal IsDirectoryObject As Long, ByVal Token As Long, GenericMapping As GENERIC_MAPPING) As Long
    Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
    Declare Function DeleteAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceIndex As Long) As Long
    Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
    Declare Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
    Declare Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
    Declare Function DestroyPrivateObjectSecurity Lib "advapi32.dll" (ObjectDescriptor As SECURITY_DESCRIPTOR) As Long
    Declare Function DuplicateToken Lib "advapi32.dll" (ByVal ExistingTokenHandle As Long, ImpersonationLevel As Integer, DuplicateTokenHandle As Long) As Long
    Declare Function EnumDependentServices Lib "advapi32.dll" Alias "EnumDependentServicesA" (ByVal hService As Long, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long) As Long
    Declare Function EnumServicesStatus Lib "advapi32.dll" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
    Declare Function EqualPrefixSid Lib "advapi32.dll" (pSid1 As Any, pSid2 As Any) As Long
    Declare Function EqualSid Lib "advapi32.dll" (pSid1 As Any, pSid2 As Any) As Long
    Declare Function FindFirstFreeAce Lib "advapi32.dll" (pAcl As ACL, pAce As Long) As Long
    Declare Sub FreeSid Lib "advapi32.dll" (pSid As Any)
    Declare Function GetAce Lib "advapi32.dll" (pAcl As ACL, ByVal dwAceIndex As Long, pAce As Any) As Long
    Declare Function GetAclInformation Lib "advapi32.dll" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
    Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
    Declare Function GetKernelObjectSecurity Lib "advapi32.dll" (ByVal Handle As Long, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
    Declare Function GetLengthSid Lib "advapi32.dll" (pSid As Any) As Long
    declare Function GetOldestEventLogRecord Lib "advapi32.dll" (ByVal hEventLog As Long, OldestRecord As Long) As Long
    Declare Function GetPrivateObjectSecurity Lib "advapi32.dll" (ObjectDescriptor As SECURITY_DESCRIPTOR, ByVal SecurityInformation As Long, ResultantDescriptor As SECURITY_DESCRIPTOR, ByVal DescriptorLength As Long, ReturnLength As Long) As Long
    Declare Function GetSecurityDescriptorControl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pControl As Integer, lpdwRevision As Long) As Long
    Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, lpbDaclPresent As Long, pDacl As ACL, lpbDaclDefaulted As Long) As Long
    Declare Function GetSecurityDescriptorGroup Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal lpbGroupDefaulted As Long) As Long
    Declare Function GetSecurityDescriptorLength Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
    Declare Function GetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal lpbOwnerDefaulted As Long) As Long
    Declare Function GetSecurityDescriptorSacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal lpbSaclPresent As Long, pSacl As ACL, ByVal lpbSaclDefaulted As Long) As Long
    Declare Function GetServiceDisplayName Lib "advapi32.dll" Alias "GetServiceDisplayNameA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, lpcchBuffer As Long) As Long
    Declare Function GetServiceKeyName Lib "advapi32.dll" Alias "GetServiceKeyNameA" (ByVal hSCManager As Long, ByVal lpDisplayName As String, ByVal lpServiceName As String, lpcchBuffer As Long) As Long
    Declare Function GetSidIdentifierAuthority Lib "advapi32.dll" (pSid As Any) As SID_IDENTIFIER_AUTHORITY
    Declare Function GetSidLengthRequired Lib "advapi32.dll" (ByVal nSubAuthorityCount As Byte) As Long
    Declare Function GetSidSubAuthority Lib "advapi32.dll" (pSid As Any, ByVal nSubAuthority As Long) As Long
    Declare Function GetSidSubAuthorityCount Lib "advapi32.dll" (pSid As Any) As Byte
    Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
    Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare Function ImpersonateNamedPipeClient Lib "advapi32.dll" (ByVal hNamedPipe As Long) As Long
    Declare Function ImpersonateSelf Lib "advapi32.dll" (ImpersonationLevel As Integer) As Long
    Declare Function InitializeAcl Lib "advapi32.dll" (pAcl As ACL, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
    Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
    Declare Function InitializeSid Lib "advapi32.dll" (Sid As Any, pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte) As Long
    Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
    Declare Function LockServiceDatabase Lib "advapi32.dll" (ByVal hSCManager As Long) As Long
    Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, Sid As Long, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
    Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, Sid As Any, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
    Declare Function LookupPrivilegeDisplayName Lib "advapi32.dll" Alias "LookupPrivilegeDisplayNameA" (ByVal lpSystemName As String, ByVal lpName As String, ByVal lpDisplayName As String, cbDisplayName As Long, lpLanguageID As Long) As Long
    Declare Function LookupPrivilegeName Lib "advapi32.dll" Alias "LookupPrivilegeNameA" (ByVal lpSystemName As String, lpLuid As LARGE_INTEGER, ByVal lpName As String, cbName As Long) As Long
    Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
    Declare Function NotifyBootConfigStatus Lib "advapi32.dll" (ByVal BootAcceptable As Long) As Long
    Declare Function NotifyChangeEventLog Lib "advapi32" (ByVal hEventLog As Long, ByVal hEvent As Long) As Boolean
    Declare Function ObjectCloseAuditAlarm Lib "advapi32.dll" Alias "ObjectCloseAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal GenerateOnClose As Long) As Long
    Declare Function ObjectPrivilegeAuditAlarm Lib "advapi32.dll" Alias "ObjectPrivilegeAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ClientToken As Long, ByVal DesiredAccess As Long, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
    Declare Function OpenBackupEventLog Lib "advapi32.dll" Alias "OpenBackupEventLogA" (ByVal lpUNCServerName As String, ByVal lpFileName As String) As Long
    Declare Function OpenEventLog Lib "advapi32.dll" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
    Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
    Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
    Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
    Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long
    Declare Function PrivilegeCheck Lib "advapi32.dll" (ByVal ClientToken As Long, RequiredPrivileges As PRIVILEGE_SET, ByVal pfResult As Long) As Long
    Declare Function PrivilegedServiceAuditAlarm Lib "advapi32.dll" Alias "PrivilegedServiceAuditAlarmA" (ByVal SubsystemName As String, ByVal ServiceName As String, ByVal ClientToken As Long, Privileges As PRIVILEGE_SET, ByVal AccessGranted As Long) As Long
    Declare Function ReadEventLog Lib "advapi32.dll" Alias "ReadEventLogA" (ByVal hEventLog As Long, ByVal dwReadFlags As Long, ByVal dwRecordOffset As Long, lpBuffer As EVENTLOGRECORD, ByVal nNumberOfBytesToRead As Long, pnBytesRead As Long, pnMinNumberOfBytesNeeded As Long) As Long
    Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
    Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
    Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
    Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
    Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
    Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
    Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
    Declare Function RegisterServiceCtrlHandler Lib "advapi32.dll" Alias "RegisterServiceCtrlHandlerA" (ByVal lpServiceName As String, ByVal lpHandlerProc As Long) As Long
    Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
    Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
    Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
    Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
    Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
    Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
    Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
    Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
    Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
    Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    Declare Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Long, ByVal wCategory As Long, ByVal dwEventID As Long, lpUserSid As Any, ByVal wNumStrings As Long, ByVal dwDataSize As Long, ByVal lpStrings As Long, lpRawData As Any) As Long
    Declare Function RevertToSelf Lib "advapi32.dll" () As Long
    Declare Function SetAclInformation Lib "advapi32.dll" (pAcl As ACL, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Integer) As Long
    Declare Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
    Declare Function SetPrivateObjectSecurity Lib "advapi32.dll" (ByVal SecurityInformation As Long, ModificationDescriptor As SECURITY_DESCRIPTOR, ObjectsSecurityDescriptor As SECURITY_DESCRIPTOR, GenericMapping As GENERIC_MAPPING, ByVal Token As Long) As Long
    Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, pDacl As ACL, ByVal bDaclDefaulted As Long) As Long
    Declare Function SetSecurityDescriptorGroup Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pGroup As Any, ByVal bGroupDefaulted As Long) As Long
    Declare Function SetSecurityDescriptorOwner Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, pOwner As Any, ByVal bOwnerDefaulted As Long) As Long
    Declare Function SetSecurityDescriptorSacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bSaclPresent As Long, pSacl As ACL, ByVal bSaclDefaulted As Long) As Long
    Declare Function SetServiceBits Lib "advapi32" (ByVal hServiceStatus As Long, ByVal dwServiceBits As Long, ByVal bSetBitsOn As Boolean, ByVal bUpdateImmediately As Boolean) As Boolean
    Declare Function SetServiceObjectSecurity Lib "advapi32.dll" (ByVal hService As Long, ByVal dwSecurityInformation As Long, lpSecurityDescriptor As Any) As Long
    Declare Function SetServiceStatus Lib "advapi32.dll" (ByVal hServiceStatus As Long, lpServiceStatus As SERVICE_STATUS) As Long
    Declare Function SetThreadToken Lib "advapi32" (Thread As Long, ByVal Token As Long) As Boolean
    Declare Function SetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long) As Long
    Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
    Declare Function StartServiceCtrlDispatcher Lib "advapi32.dll" Alias "StartServiceCtrlDispatcherA" (lpServiceStartTable As SERVICE_TABLE_ENTRY) As Long
    Declare Function UnlockServiceDatabase Lib "advapi32.dll" (ScLock As Any) As Long

#End If

