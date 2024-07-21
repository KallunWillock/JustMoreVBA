Attribute VB_Name = "API_KERNEL32"
'@Lang VBA

#If VBA7 And Win64 Then

    Declare PtrSafe Function AllocConsole Lib "kernel32.dll" Alias "AllocConsole" () As Long
    Declare PtrSafe Function BackupRead Lib "kernel32.dll" Alias "BackupRead" (ByVal hFile As LongPtr, lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As Any) As Long
    Declare PtrSafe Function BackupSeek Lib "kernel32.dll" Alias "BackupSeek" (ByVal hFile As LongPtr, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, lpdwLowByteSeeked As Long, lpdwHighByteSeeked As Long, lpContext As LongPtr) As Long
    Declare PtrSafe Function BackupWrite Lib "kernel32.dll" Alias "BackupWrite" (ByVal hFile As LongPtr, lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As LongPtr) As Long
    Declare PtrSafe Function Beep Lib "kernel32.dll" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
    Declare PtrSafe Function BeginUpdateResource Lib "kernel32.dll" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As LongPtr
    Declare PtrSafe Function BuildCommDCB Lib "kernel32.dll" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
    Declare PtrSafe Function BuildCommDCBAndTimeouts Lib "kernel32.dll" Alias "BuildCommDCBAndTimeoutsA" (ByVal lpDef As String, lpDCB As DCB, lpCommTimeouts As COMMTIMEOUTS) As Long
    Declare PtrSafe Function ClearCommBreak Lib "kernel32.dll" Alias "ClearCommBreak" (ByVal nCid As LongPtr) As Long
    Declare PtrSafe Function ClearCommError Lib "kernel32.dll" Alias "ClearCommError" (ByVal hFile As LongPtr, lpErrors As Long, lpStat As COMSTAT) As Long
    Declare PtrSafe Function CloseHandle Lib "kernel32.dll" Alias "CloseHandle" (ByVal hObject As LongPtr) As Long
    Declare PtrSafe Function CommConfigDialog Lib "kernel32.dll" Alias "CommConfigDialogA" (ByVal lpszName As String, ByVal hWnd As LongPtr, lpCC As COMMCONFIG) As Long
    Declare PtrSafe Function CompareFileTime Lib "kernel32.dll" Alias "CompareFileTime" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
    Declare PtrSafe Function CompareString Lib "kernel32.dll" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
    Declare PtrSafe Function ConnectNamedPipe Lib "kernel32.dll" Alias "ConnectNamedPipe" (ByVal hNamedPipe As LongPtr, lpOverlapped As OVERLAPPED) As Long
    Declare PtrSafe Function ContinueDebugEvent Lib "kernel32.dll" Alias "ContinueDebugEvent" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long
    Declare PtrSafe Function ConvertDefaultLocale Lib "kernel32.dll" Alias "ConvertDefaultLocale" (ByVal Locale As Long) As Long
    Declare PtrSafe Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
    Declare PtrSafe Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
    Declare PtrSafe Function CreateDirectoryEx Lib "kernel32.dll" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
    Declare PtrSafe Function CreateEvent Lib "kernel32.dll" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As LongPtr
    Declare PtrSafe Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
    Declare PtrSafe Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hFile As LongPtr, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As LongPtr
    Declare PtrSafe Function CreateIoCompletionPort Lib "kernel32.dll" Alias "CreateIoCompletionPort" (ByVal FileHandle As LongPtr, ByVal ExistingCompletionPort As LongPtr, ByVal CompletionKey As LongPtr, ByVal NumberOfConcurrentThreads As Long) As LongPtr
    Declare PtrSafe Function CreateMailslot Lib "kernel32.dll" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As LongPtr
    Declare PtrSafe Function CreateMutex Lib "kernel32.dll" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As LongPtr
    Declare PtrSafe Function CreateNamedPipe Lib "kernel32.dll" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As LongPtr
    Declare PtrSafe Function CreatePipe Lib "kernel32.dll" Alias "CreatePipe" (phReadPipe As LongPtr, phWritePipe As LongPtr, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
    Declare PtrSafe Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As LongPtr
    Declare PtrSafe Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserA" (ByVal hToken As LongPtr, ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Declare PtrSafe Function CreateRemoteThread Lib "kernel32.dll" Alias "CreateRemoteThread" (ByVal hProcess As LongPtr, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As LongPtr, lpStartAddress As LongPtr, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As LongPtr
    Declare PtrSafe Function CreateSemaphore Lib "kernel32.dll" Alias "CreateSemaphoreA" (lpSemaphoreAttributes As SECURITY_ATTRIBUTES, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As LongPtr
    Declare PtrSafe Function CreateTapePartition Lib "kernel32.dll" Alias "CreateTapePartition" (ByVal hDevice As LongPtr, ByVal dwPartitionMethod As Long, ByVal dwCount As Long, ByVal dwSize As Long) As Long
    Declare PtrSafe Function CreateThread Lib "kernel32.dll" Alias "CreateThread" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As LongPtr, lpStartAddress As LongPtr, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As LongPtr
    Declare PtrSafe Function DefineDosDevice Lib "kernel32.dll" Alias "DefineDosDeviceA" (ByVal dwFlags As Long, ByVal lpDeviceName As String, ByVal lpTargetPath As String) As Long
    Declare PtrSafe Sub DeleteCriticalSection Lib "kernel32.dll" Alias "DeleteCriticalSection" (lpCriticalSection As CRITICAL_SECTION)
    Declare PtrSafe Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
    Declare PtrSafe Function DeviceIoControl Lib "kernel32.dll" Alias "DeviceIoControl" (ByVal hDevice As LongPtr, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
    Declare PtrSafe Function DisableThreadLibraryCalls Lib "kernel32.dll" Alias "DisableThreadLibraryCalls" (ByVal hLibModule As LongPtr) As Long
    Declare PtrSafe Function DisconnectNamedPipe Lib "kernel32.dll" Alias "DisconnectNamedPipe" (ByVal hNamedPipe As LongPtr) As Long
    Declare PtrSafe Function DosDateTimeToFileTime Lib "kernel32.dll" Alias "DosDateTimeToFileTime" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
    Declare PtrSafe Function DuplicateHandle Lib "kernel32.dll" Alias "DuplicateHandle" (ByVal hSourceProcessHandle As LongPtr, ByVal hSourceHandle As LongPtr, ByVal hTargetProcessHandle As LongPtr, lpTargetHandle As LongPtr, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
    Declare PtrSafe Function EndUpdateResource Lib "kernel32.dll" Alias "EndUpdateResourceA" (ByVal hUpdate As LongPtr, ByVal fDiscard As Long) As Long
    Declare PtrSafe Sub EnterCriticalSection Lib "kernel32.dll" Alias "EnterCriticalSection" (lpCriticalSection As CRITICAL_SECTION)
    Declare PtrSafe Function EnumCalendarInfo Lib "kernel32.dll" Alias "EnumCalendarInfoA" (ByVal lpCalInfoEnumProc As LongPtr, ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long) As Long
    Declare PtrSafe Function EnumDateFormats Lib "kernel32.dll" Alias "EnumDateFormats" (ByVal lpDateFmtEnumProc As LongPtr, ByVal Locale As Long, ByVal dwFlags As Long) As Long
    Declare PtrSafe Function EnumResourceLanguages Lib "kernel32.dll" Alias "EnumResourceLanguagesA" (ByVal hModule As LongPtr, ByVal lpType As String, ByVal lpName As String, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare PtrSafe Function EnumResourceNames Lib "kernel32.dll" Alias "EnumResourceNamesA" (ByVal hModule As LongPtr, ByVal lpType As String, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare PtrSafe Function EnumResourceTypes Lib "kernel32.dll" Alias "EnumResourceTypesA" (ByVal hModule As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Declare PtrSafe Function EnumSystemCodePages Lib "kernel32.dll" Alias "EnumSystemCodePages" (ByVal lpCodePageEnumProc As LongPtr, ByVal dwFlags As Long) As Long
    Declare PtrSafe Function EnumSystemLocales Lib "kernel32.dll" Alias "EnumSystemLocales" (ByVal lpLocaleEnumProc As LongPtr, ByVal dwFlags As Long) As Long
    Declare PtrSafe Function EnumTimeFormats Lib "kernel32.dll" Alias "EnumTimeFormats" (ByVal lpTimeFmtEnumProc As LongPtr, ByVal Locale As Long, ByVal dwFlags As Long) As Long
    Declare PtrSafe Function EraseTape Lib "kernel32.dll" Alias "EraseTape" (ByVal hDevice As LongPtr, ByVal dwEraseType As Long, ByVal bimmediate As Long) As Long
    Declare PtrSafe Function EscapeCommFunction Lib "kernel32.dll" Alias "EscapeCommFunction" (ByVal nCid As LongPtr, ByVal nFunc As Long) As Long
    Declare PtrSafe Sub ExitProcess Lib "kernel32.dll" Alias "ExitProcess" (ByVal uExitCode As Long)
    Declare PtrSafe Sub ExitThread Lib "kernel32.dll" Alias "ExitThread" (ByVal dwExitCode As Long)
    Declare PtrSafe Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
    Declare PtrSafe Sub FatalAppExit Lib "kernel32.dll" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
    Declare PtrSafe Sub FatalExit Lib "kernel32.dll" Alias "FatalExit" (ByVal code As Long)
    Declare PtrSafe Function FileTimeToDosDateTime Lib "kernel32.dll" Alias "FileTimeToDosDateTime" (lpFileTime As FILETIME, ByVal lpFatDate As LongPtr, ByVal lpFatTime As LongPtr) As Long
    Declare PtrSafe Function FileTimeToLocalFileTime Lib "kernel32.dll" Alias "FileTimeToLocalFileTime" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
    Declare PtrSafe Function FileTimeToSystemTime Lib "kernel32.dll" Alias "FileTimeToSystemTime" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
    Declare PtrSafe Function FillConsoleOutputAttribute Lib "kernel32.dll" Alias "FillConsoleOutputAttribute" (ByVal hConsoleOutput As LongPtr, ByVal wAttribute As Long, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
    Declare PtrSafe Function FillConsoleOutputCharacter Lib "kernel32.dll" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As LongPtr, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
    Declare PtrSafe Function FindClose Lib "kernel32.dll" Alias "FindClose" (ByVal hFindFile As LongPtr) As Long
    Declare PtrSafe Function FindCloseChangeNotification Lib "kernel32.dll" Alias "FindCloseChangeNotification" (ByVal hChangeHandle As LongPtr) As Long
    Declare PtrSafe Function FindFirstChangeNotification Lib "kernel32.dll" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As LongPtr
    Declare PtrSafe Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As LongPtr
    Declare PtrSafe Function FindNextChangeNotification Lib "kernel32.dll" Alias "FindNextChangeNotification" (ByVal hChangeHandle As LongPtr) As Long
    Declare PtrSafe Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As LongPtr, lpFindFileData As WIN32_FIND_DATA) As Long
    Declare PtrSafe Function FindResource Lib "kernel32.dll" Alias "FindResourceA" (ByVal hInstance As LongPtr, ByVal lpName As String, ByVal lpType As String) As LongPtr
    Declare PtrSafe Function FindResourceEx Lib "kernel32.dll" Alias "FindResourceExA" (ByVal hModule As LongPtr, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long) As LongPtr
    Declare PtrSafe Function FreeConsole Lib "kernel32.dll" Alias "FreeConsole" () As Long
    Declare PtrSafe Function FreeEnvironmentStrings Lib "kernel32.dll" Alias "FreeEnvironmentStringsA" (ByVal lpsz As String) As Long
    Declare PtrSafe Function FreeLibrary Lib "kernel32.dll" Alias "FreeLibrary" (ByVal hLibModule As LongPtr) As Long
    Declare PtrSafe Sub FreeLibraryAndExitThread Lib "kernel32.dll" Alias "FreeLibraryAndExitThread" (ByVal hLibModule As LongPtr, ByVal dwExitCode As Long)
    Declare PtrSafe Function FreeResource Lib "kernel32.dll" Alias "FreeResource" (ByVal hResData As LongPtr) As Long
    Declare PtrSafe Function GenerateConsoleCtrlEvent Lib "kernel32.dll" Alias "GenerateConsoleCtrlEvent" (ByVal dwCtrlEvent As Long, ByVal dwProcessGroupId As Long) As Long
    Declare PtrSafe Function GetACP Lib "kernel32.dll" Alias "GetACP" () As Long
    Declare PtrSafe Function GetBinaryType Lib "kernel32.dll" Alias "GetBinaryTypeA" (ByVal lpApplicationName As String, lpBinaryType As Long) As Long
    Declare PtrSafe Function GetCommandLine Lib "kernel32.dll" Alias "GetCommandLineA" () As String
    Declare PtrSafe Function GetCommConfig Lib "kernel32.dll" Alias "GetCommConfig" (ByVal hCommDev As LongPtr, lpCC As COMMCONFIG, lpdwSize As Long) As Long
    Declare PtrSafe Function GetCommMask Lib "kernel32.dll" Alias "GetCommMask" (ByVal hFile As LongPtr, lpEvtMask As Long) As Long
    Declare PtrSafe Function GetCommModemStatus Lib "kernel32.dll" Alias "GetCommModemStatus" (ByVal hFile As LongPtr, lpModemStat As Long) As Long
    Declare PtrSafe Function GetCommProperties Lib "kernel32.dll" Alias "GetCommProperties" (ByVal hFile As LongPtr, lpCommProp As COMMPROP) As Long
    Declare PtrSafe Function GetCommState Lib "kernel32.dll" Alias "GetCommState" (ByVal nCid As LongPtr, lpDCB As DCB) As Long
    Declare PtrSafe Function GetCommTimeouts Lib "kernel32.dll" Alias "GetCommTimeouts" (ByVal hFile As LongPtr, lpCommTimeouts As COMMTIMEOUTS) As Long
    Declare PtrSafe Function GetCompressedFileSize Lib "kernel32.dll" Alias "GetCompressedFileSizeA" (ByVal lpFileName As String, lpFileSizeHigh As Long) As Long
    Declare PtrSafe Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare PtrSafe Function GetConsoleCP Lib "kernel32.dll" Alias "GetConsoleCP" () As Long
    Declare PtrSafe Function GetConsoleCursorInfo Lib "kernel32.dll" Alias "GetConsoleCursorInfo" (ByVal hConsoleOutput As LongPtr, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
    Declare PtrSafe Function GetConsoleMode Lib "kernel32.dll" Alias "GetConsoleMode" (ByVal hConsoleHandle As LongPtr, lpMode As Long) As Long
    Declare PtrSafe Function GetConsoleOutputCP Lib "kernel32.dll" Alias "GetConsoleOutputCP" () As Long
    Declare PtrSafe Function GetConsoleScreenBufferInfo Lib "kernel32.dll" Alias "GetConsoleScreenBufferInfo" (ByVal hConsoleOutput As LongPtr, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
    Declare PtrSafe Function GetDateFormat Lib "kernel32.dll" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
    Declare PtrSafe Function GetDefaultCommConfig Lib "kernel32.dll" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long
    Declare PtrSafe Function GetDiskFreeSpace Lib "kernel32.dll" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
    Declare PtrSafe Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As LongPtr
    Declare PtrSafe Function GetEnvironmentStrings Lib "kernel32.dll" Alias "GetEnvironmentStringsA" () As String
    Declare PtrSafe Function GetEnvironmentVariable Lib "kernel32.dll" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare PtrSafe Function GetExitCodeProcess Lib "kernel32.dll" Alias "GetExitCodeProcess" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
    Declare PtrSafe Function GetExitCodeThread Lib "kernel32.dll" Alias "GetExitCodeThread" (ByVal hThread As LongPtr, lpExitCode As Long) As Long
    Declare PtrSafe Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
    Declare PtrSafe Function GetFileInformationByHandle Lib "kernel32.dll" Alias "GetFileInformationByHandle" (ByVal hFile As LongPtr, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
    Declare PtrSafe Function GetFileSize Lib "kernel32.dll" Alias "GetFileSize" (ByVal hFile As LongPtr, lpFileSizeHigh As Long) As Long
    Declare PtrSafe Function GetFileTime Lib "kernel32.dll" Alias "GetFileTime" (ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
    Declare PtrSafe Function GetFileType Lib "kernel32.dll" Alias "GetFileType" (ByVal hFile As LongPtr) As Long
    Declare PtrSafe Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
    Declare PtrSafe Function GetHandleInformation Lib "kernel32.dll" Alias "GetHandleInformation" (ByVal hObject As LongPtr, lpdwFlags As Long) As Long
    Declare PtrSafe Function GetLargestConsoleWindowSize Lib "kernel32.dll" Alias "GetLargestConsoleWindowSize" (ByVal hConsoleOutput As LongPtr) As COORD
    Declare PtrSafe Function GetLastError Lib "kernel32.dll" Alias "GetLastError" () As Long
    Declare PtrSafe Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Declare PtrSafe Sub GetLocalTime Lib "kernel32.dll" Alias "GetLocalTime" (lpSystemTime As SYSTEMTIME)
    Declare PtrSafe Function GetLogicalDrives Lib "kernel32.dll" Alias "GetLogicalDrives" () As Long
    Declare PtrSafe Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Declare PtrSafe Function GetMailslotInfo Lib "kernel32.dll" Alias "GetMailslotInfo" (ByVal hMailslot As LongPtr, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
    Declare PtrSafe Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameA" (ByVal hModule As LongPtr, ByVal lpFileName As String, ByVal nSize As Long) As Long
    Declare PtrSafe Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Declare PtrSafe Function GetOEMCP Lib "kernel32.dll" Alias "GetOEMCP" () As Long
    Declare PtrSafe Function GetOverlappedResult Lib "kernel32.dll" Alias "GetOverlappedResult" (ByVal hFile As LongPtr, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
    Declare PtrSafe Function GetPriorityClass Lib "kernel32.dll" Alias "GetPriorityClass" (ByVal hProcess As LongPtr) As Long
    Declare PtrSafe Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
    Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32.dll" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare PtrSafe Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare PtrSafe Function GetProcAddress Lib "kernel32.dll" Alias "GetProcAddress" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Declare PtrSafe Function GetProcessAffinityMask Lib "kernel32.dll" Alias "GetProcessAffinityMask" (ByVal hProcess As LongPtr, lpProcessAffinityMask As LongPtr, SystemAffinityMask As LongPtr) As Long
    Declare PtrSafe Function GetProcessHeap Lib "kernel32.dll" Alias "GetProcessHeap" () As LongPtr
    Declare PtrSafe Function GetProcessHeaps Lib "kernel32.dll" Alias "GetProcessHeaps" (ByVal NumberOfHeaps As Long, ProcessHeaps As LongPtr) As Long
    Declare PtrSafe Function GetProcessShutdownParameters Lib "kernel32.dll" Alias "GetProcessShutdownParameters" (lpdwLevel As Long, lpdwFlags As Long) As Long
    Declare PtrSafe Function GetProcessTimes Lib "kernel32.dll" Alias "GetProcessTimes" (ByVal hProcess As LongPtr, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
    Declare PtrSafe Function GetProcessWorkingSetSize Lib "kernel32.dll" Alias "GetProcessWorkingSetSize" (ByVal hProcess As LongPtr, lpMinimumWorkingSetSize As LongPtr, lpMaximumWorkingSetSize As LongPtr) As Long
    Declare PtrSafe Function GetProfileInt Lib "kernel32.dll" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
    Declare PtrSafe Function GetProfileSection Lib "kernel32.dll" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
    Declare PtrSafe Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
    Declare PtrSafe Function GetQueuedCompletionStatus Lib "kernel32.dll" Alias "GetQueuedCompletionStatus" (ByVal CompletionPort As LongPtr, lpNumberOfBytesTransferred As Long, lpCompletionKey As LongPtr, lpOverlapped As LongPtr, ByVal dwMilliseconds As Long) As Long
    Declare PtrSafe Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
    Declare PtrSafe Sub GetStartupInfo Lib "kernel32.dll" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
    Declare PtrSafe Function GetStdHandle Lib "kernel32.dll" Alias "GetStdHandle" (ByVal nStdHandle As Long) As LongPtr
    Declare PtrSafe Function GetStringTypeEx Lib "kernel32.dll" Alias "GetStringTypeExA" (ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Long
    Declare PtrSafe Function GetSystemPowerStatus Lib "kernel32.dll" Alias "GetSystemPowerStatus" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
    Declare PtrSafe Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTime" (lpSystemTime As SYSTEMTIME)
    Declare PtrSafe Function GetSystemTimeAdjustment Lib "kernel32.dll" Alias "GetSystemTimeAdjustment" (lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Long) As Long
    Declare PtrSafe Function GetTapeParameters Lib "kernel32.dll" Alias "GetTapeParameters" (ByVal hDevice As LongPtr, ByVal dwOperation As Long, lpdwSize As Long, lpTapeInformation As Any) As Long
    Declare PtrSafe Function GetTapePosition Lib "kernel32.dll" Alias "GetTapePosition" (ByVal hDevice As LongPtr, ByVal dwPositionType As Long, lpdwPartition As Long, lpdwOffsetLow As Long, lpdwOffsetHigh As Long) As Long
    Declare PtrSafe Function GetTapeStatus Lib "kernel32.dll" Alias "GetTapeStatus" (ByVal hDevice As LongPtr) As Long
    Declare PtrSafe Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
    Declare PtrSafe Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As LongPtr, ByVal lpBuffer As String) As Long
    Declare PtrSafe Function GetThreadContext Lib "kernel32.dll" Alias "GetThreadContext" (ByVal hThread As LongPtr, lpContext As CONTEXT) As Long
    Declare PtrSafe Function GetThreadLocale Lib "kernel32.dll" Alias "GetThreadLocale" () As Long
    Declare PtrSafe Function GetThreadPriority Lib "kernel32.dll" Alias "GetThreadPriority" (ByVal hThread As LongPtr) As Long
    Declare PtrSafe Function GetThreadSelectorEntry Lib "kernel32.dll" Alias "GetThreadSelectorEntry" (ByVal hThread As LongPtr, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long
    Declare PtrSafe Function GetThreadTimes Lib "kernel32.dll" Alias "GetThreadTimes" (ByVal hThread As LongPtr, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
    Declare PtrSafe Function GetTickCount Lib "kernel32.dll" Alias "GetTickCount" () As Long
    Declare PtrSafe Function GetTimeFormat Lib "kernel32.dll" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
    Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32.dll" Alias "GetTimeZoneInformation" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    Declare PtrSafe Function GetUserDefaultLangID Lib "kernel32.dll" Alias "GetUserDefaultLangID" () As Integer
    Declare PtrSafe Function GetUserDefaultLCID Lib "kernel32.dll" Alias "GetUserDefaultLCID" () As Long
    Declare PtrSafe Function GetVersion Lib "kernel32.dll" Alias "GetVersion" () As Long
    Declare PtrSafe Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    Declare PtrSafe Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
    Declare PtrSafe Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare PtrSafe Function GlobalAddAtom Lib "kernel32.dll" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
    Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalCompact Lib "kernel32.dll" Alias "GlobalCompact" (ByVal dwMinFree As Long) As LongPtr
    Declare PtrSafe Function GlobalDeleteAtom Lib "kernel32.dll" Alias "GlobalDeleteAtom" (ByVal nAtom As Integer) As Integer
    Declare PtrSafe Function GlobalFindAtom Lib "kernel32.dll" Alias "GlobalFindAtomA" (ByVal lpString As String) As Integer
    Declare PtrSafe Sub GlobalFix Lib "kernel32.dll" Alias "GlobalFix" (ByVal hMem As LongPtr)
    Declare PtrSafe Function GlobalFlags Lib "kernel32.dll" Alias "GlobalFlags" (ByVal hMem As LongPtr) As Long
    Declare PtrSafe Function GlobalFree Lib "kernel32.dll" Alias "GlobalFree" (ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalGetAtomName Lib "kernel32.dll" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare PtrSafe Function HeapAlloc Lib "kernel32.dll" Alias "HeapAlloc" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Declare PtrSafe Function HeapCompact Lib "kernel32.dll" Alias "HeapCompact" (ByVal hHeap As LongPtr, ByVal dwFlags As Long) As LongPtr
    Declare PtrSafe Function HeapCreate Lib "kernel32.dll" Alias "HeapCreate" (ByVal flOptions As Long, ByVal dwInitialSize As LongPtr, ByVal dwMaximumSize As LongPtr) As LongPtr
    Declare PtrSafe Function HeapDestroy Lib "kernel32.dll" Alias "HeapDestroy" (ByVal hHeap As LongPtr) As Long
    Declare PtrSafe Function HeapFree Lib "kernel32.dll" Alias "HeapFree" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any) As Long
    Declare PtrSafe Function HeapLock Lib "kernel32.dll" Alias "HeapLock" (ByVal hHeap As LongPtr) As Long
    Declare PtrSafe Function HeapReAlloc Lib "kernel32.dll" Alias "HeapReAlloc" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As LongPtr) As LongPtr
    Declare PtrSafe Function HeapSize Lib "kernel32.dll" Alias "HeapSize" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any) As LongPtr
    Declare PtrSafe Function HeapUnlock Lib "kernel32.dll" Alias "HeapUnlock" (ByVal hHeap As LongPtr) As Long
    Declare PtrSafe Function HeapValidate Lib "kernel32.dll" Alias "HeapValidate" (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any) As Long
    Declare PtrSafe Function hread Lib "kernel32.dll" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
    Declare PtrSafe Function hwrite Lib "kernel32.dll" Alias "_hwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal lBytes As Long) As Long
    Declare PtrSafe Function ImpersonateLoggedOnUser Lib "kernel32.dll" Alias "ImpersonateLoggedOnUser" (ByVal hToken As LongPtr) As Long
    Declare PtrSafe Function InitAtomTable Lib "kernel32.dll" Alias "InitAtomTable" (ByVal nSize As Long) As Long
    Declare PtrSafe Sub InitializeCriticalSection Lib "kernel32.dll" Alias "InitializeCriticalSection" (lpCriticalSection As CRITICAL_SECTION)
    Declare PtrSafe Function InterlockedDecrement Lib "kernel32.dll" Alias "InterlockedDecrement" (lpAddend As Long) As Long
    Declare PtrSafe Function InterlockedExchange Lib "kernel32.dll" Alias "InterlockedExchange" (Target As Long, ByVal Value As Long) As Long
    Declare PtrSafe Function InterlockedIncrement Lib "kernel32.dll" Alias "InterlockedIncrement" (lpAddend As Long) As Long
    Declare PtrSafe Function IsBadCodePtr Lib "kernel32.dll" Alias "IsBadCodePtr" (ByVal lpfn As LongPtr) As Long
    Declare PtrSafe Function IsBadHugeReadPtr Lib "kernel32.dll" Alias "IsBadHugeReadPtr" (lp As Any, ByVal ucb As LongPtr) As Long
    Declare PtrSafe Function IsBadHugeWritePtr Lib "kernel32.dll" Alias "IsBadHugeWritePtr" (lp As Any, ByVal ucb As LongPtr) As Long
    Declare PtrSafe Function IsBadReadPtr Lib "kernel32.dll" Alias "IsBadReadPtr" (lp As Any, ByVal ucb As LongPtr) As Long
    Declare PtrSafe Function IsBadStringPtr Lib "kernel32.dll" Alias "IsBadStringPtrA" (ByVal lpsz As String, ByVal ucchMax As LongPtr) As Long
    Declare PtrSafe Function IsBadWritePtr Lib "kernel32.dll" Alias "IsBadWritePtr" (lp As Any, ByVal ucb As LongPtr) As Long
    Declare PtrSafe Function IsDBCSLeadByte Lib "kernel32.dll" Alias "IsDBCSLeadByte" (ByVal bTestChar As Byte) As Long
    Declare PtrSafe Function lclose Lib "kernel32.dll" Alias "_lclose" (ByVal hFile As Long) As Long
    Declare PtrSafe Function LCMapString Lib "kernel32.dll" Alias "LCMapStringA" (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
    Declare PtrSafe Function lcreat Lib "kernel32.dll" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
    Declare PtrSafe Sub LeaveCriticalSection Lib "kernel32.dll" Alias "LeaveCriticalSection" (lpCriticalSection As CRITICAL_SECTION)
    Declare PtrSafe Function llseek Lib "kernel32.dll" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
    Declare Ptrsafe Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Declare PtrSafe Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As LongPtr, ByVal dwFlags As Long) As LongPtr
    Declare PtrSafe Function LoadModule Lib "kernel32.dll" Alias "LoadModule" (ByVal lpModuleName As String, lpParameterBlock As Any) As Long
    Declare PtrSafe Function LoadResource Lib "kernel32.dll" Alias "LoadResource" (ByVal hInstance As LongPtr, ByVal hResInfo As LongPtr) As LongPtr
    Declare PtrSafe Function LocalAlloc Lib "kernel32.dll" Alias "LocalAlloc" (ByVal wFlags As Long, ByVal wBytes As LongPtr) As LongPtr
    Declare PtrSafe Function LocalCompact Lib "kernel32.dll" Alias "LocalCompact" (ByVal uMinFree As Long) As LongPtr
    Declare PtrSafe Function LocalFileTimeToFileTime Lib "kernel32.dll" Alias "LocalFileTimeToFileTime" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
    Declare PtrSafe Function LocalFlags Lib "kernel32.dll" Alias "LocalFlags" (ByVal hMem As LongPtr) As Long
    Declare PtrSafe Function LocalFree Lib "kernel32.dll" Alias "LocalFree" (ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function LocalHandle Lib "kernel32.dll" Alias "LocalHandle" (wMem As Any) As LongPtr
    Declare PtrSafe Function LocalLock Lib "kernel32.dll" Alias "LocalLock" (ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function LocalReAlloc Lib "kernel32.dll" Alias "LocalReAlloc" (ByVal hMem As LongPtr, ByVal wBytes As LongPtr, ByVal wFlags As Long) As LongPtr
    Declare PtrSafe Function LocalShrink Lib "kernel32.dll" Alias "LocalShrink" (ByVal hMem As LongPtr, ByVal cbNewSize As Long) As LongPtr
    Declare PtrSafe Function LocalSize Lib "kernel32.dll" Alias "LocalSize" (ByVal hMem As LongPtr) As Long
    Declare PtrSafe Function LocalUnlock Lib "kernel32.dll" Alias "LocalUnlock" (ByVal hMem As LongPtr) As Long
    Declare PtrSafe Function LockFile Lib "kernel32.dll" Alias "LockFile" (ByVal hFile As LongPtr, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
    Declare PtrSafe Function LockFileEx Lib "kernel32.dll" Alias "LockFileEx" (ByVal hFile As LongPtr, ByVal dwFlags As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long, lpOverlapped As OVERLAPPED) As Long
    Declare PtrSafe Function LockResource Lib "kernel32.dll" Alias "LockResource" (ByVal hResData As LongPtr) As LongPtr
    Declare PtrSafe Function LogonUser Lib "kernel32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As LongPtr) As Long
    Declare PtrSafe Function lopen Lib "kernel32.dll" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
    Declare PtrSafe Function lread Lib "kernel32.dll" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
    Declare PtrSafe Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
    Declare PtrSafe Function lstrcmp Lib "kernel32.dll" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Declare PtrSafe Function lstrcmpi Lib "kernel32.dll" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
    Declare PtrSafe Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As LongPtr
    Declare PtrSafe Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
    Declare PtrSafe Function lwrite Lib "kernel32.dll" Alias "_lwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long) As Long
    Declare PtrSafe Function MapViewOfFile Lib "kernel32.dll" Alias "MapViewOfFile" (ByVal hFileMappingObject As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As LongPtr) As LongPtr
    Declare PtrSafe Function MapViewOfFileEx Lib "kernel32.dll" Alias "MapViewOfFileEx" (ByVal hFileMappingObject As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As LongPtr, lpBaseAddress As Any) As LongPtr
    Declare PtrSafe Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
    Declare PtrSafe Function MoveFileEx Lib "kernel32.dll" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
    Declare PtrSafe Function MulDiv Lib "kernel32.dll" Alias "MulDiv" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Declare PtrSafe Function MultiByteToWideChar Lib "kernel32.dll" Alias "MultiByteToWideChar" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
    Declare PtrSafe Function ObjectOpenAuditAlarm Lib "kernel32.dll" Alias "ObjectOpenAuditAlarmA" (ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As LongPtr, ByVal DesiredAccess As Long, ByVal GrantedAccess As Long, Privileges As PRIVILEGE_SET, ByVal ObjectCreation As Long, ByVal AccessGranted As Long, ByVal GenerateOnClose As LongPtr) As Long
    Declare PtrSafe Function OpenEvent Lib "kernel32.dll" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As LongPtr
    Declare PtrSafe Function OpenFile Lib "kernel32.dll" Alias "OpenFile" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
    Declare PtrSafe Function OpenFileMapping Lib "kernel32.dll" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As LongPtr
    Declare PtrSafe Function OpenMutex Lib "kernel32.dll" Alias "OpenMutexA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As LongPtr
    Declare PtrSafe Function OpenProcess Lib "kernel32.dll" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
    Declare PtrSafe Function OpenSemaphore Lib "kernel32.dll" Alias "OpenSemaphoreA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As LongPtr
    Declare PtrSafe Sub OutputDebugString Lib "kernel32.dll" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
    Declare PtrSafe Function PeekNamedPipe Lib "kernel32.dll" Alias "PeekNamedPipe" (ByVal hNamedPipe As LongPtr, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
    Declare PtrSafe Function PrepareTape Lib "kernel32.dll" Alias "PrepareTape" (ByVal hDevice As LongPtr, ByVal dwOperation As Long, ByVal bimmediate As Long) As Long
    Declare PtrSafe Function PulseEvent Lib "kernel32.dll" Alias "PulseEvent" (ByVal hEvent As LongPtr) As Long
    Declare PtrSafe Function ReadConsoleOutputAttribute Lib "kernel32.dll" Alias "ReadConsoleOutputAttribute" (ByVal hConsoleOutput As LongPtr, lpAttribute As Long, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfAttrsRead As Long) As Long
    Declare PtrSafe Function ReadConsoleOutputCharacter Lib "kernel32.dll" Alias "ReadConsoleOutputCharacterA" (ByVal hConsoleOutput As LongPtr, ByVal lpCharacter As String, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfCharsRead As Long) As Long
    Declare PtrSafe Function ReadFile Lib "kernel32.dll" Alias "ReadFile" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
    Declare PtrSafe Function ReadFileEx Lib "kernel32.dll" Alias "ReadFileEx" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As LongPtr) As Long
    Declare PtrSafe Function ReadProcessMemory Lib "kernel32.dll" Alias "ReadProcessMemory" (ByVal hProcess As LongPtr, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As LongPtr, lpNumberOfBytesWritten As LongPtr) As Long
    Declare PtrSafe Function ReleaseMutex Lib "kernel32.dll" Alias "ReleaseMutex" (ByVal hMutex As LongPtr) As Long
    Declare PtrSafe Function ReleaseSemaphore Lib "kernel32.dll" Alias "ReleaseSemaphore" (ByVal hSemaphore As LongPtr, ByVal lReleaseCount As Long, lpPreviousCount As Long) As Long
    Declare PtrSafe Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
    Declare PtrSafe Function ResetEvent Lib "kernel32.dll" Alias "ResetEvent" (ByVal hEvent As LongPtr) As Long
    Declare PtrSafe Function ResumeThread Lib "kernel32.dll" Alias "ResumeThread" (ByVal hThread As LongPtr) As Long
    Declare PtrSafe Function SetCommBreak Lib "kernel32.dll" Alias "SetCommBreak" (ByVal nCid As LongPtr) As Long
    Declare PtrSafe Function SetCommConfig Lib "kernel32.dll" Alias "SetCommConfig" (ByVal hCommDev As LongPtr, lpCC As COMMCONFIG, ByVal dwSize As Long) As Long
    Declare PtrSafe Function SetCommMask Lib "kernel32.dll" Alias "SetCommMask" (ByVal hFile As LongPtr, ByVal dwEvtMask As Long) As Long
    Declare PtrSafe Function SetCommState Lib "kernel32.dll" Alias "SetCommState" (ByVal hCommDev As LongPtr, lpDCB As DCB) As Long
    Declare PtrSafe Function SetCommTimeouts Lib "kernel32.dll" Alias "SetCommTimeouts" (ByVal hFile As LongPtr, lpCommTimeouts As COMMTIMEOUTS) As Long
    Declare PtrSafe Function SetComputerName Lib "kernel32.dll" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
    Declare PtrSafe Function SetConsoleActiveScreenBuffer Lib "kernel32.dll" Alias "SetConsoleActiveScreenBuffer" (ByVal hConsoleOutput As LongPtr) As Long
    Declare PtrSafe Function SetConsoleCP Lib "kernel32.dll" Alias "SetConsoleCP" (ByVal wCodePageID As Long) As Long
    Declare PtrSafe Function SetConsoleCtrlHandler Lib "kernel32.dll" Alias "SetConsoleCtrlHandler" (ByVal HandlerRoutine As LongPtr, ByVal Add As Long) As Long
    Declare PtrSafe Function SetConsoleCursorInfo Lib "kernel32.dll" Alias "SetConsoleCursorInfo" (ByVal hConsoleOutput As LongPtr, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
    Declare PtrSafe Function SetConsoleCursorPosition Lib "kernel32.dll" Alias "SetConsoleCursorPosition" (ByVal hConsoleOutput As LongPtr, dwCursorPosition As COORD) As Long
    Declare PtrSafe Function SetConsoleMode Lib "kernel32.dll" Alias "SetConsoleMode" (ByVal hConsoleHandle As LongPtr, ByVal dwMode As Long) As Long
    Declare PtrSafe Function SetConsoleOutputCP Lib "kernel32.dll" Alias "SetConsoleOutputCP" (ByVal wCodePageID As Long) As Long
    Declare PtrSafe Function SetConsoleScreenBufferSize Lib "kernel32.dll" Alias "SetConsoleScreenBufferSize" (ByVal hConsoleOutput As LongPtr, dwSize As COORD) As Long
    Declare PtrSafe Function SetConsoleTextAttribute Lib "kernel32.dll" Alias "SetConsoleTextAttribute" (ByVal hConsoleOutput As LongPtr, ByVal wAttributes As Long) As Long
    Declare PtrSafe Function SetConsoleTitle Lib "kernel32.dll" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
    Declare PtrSafe Function SetConsoleWindowInfo Lib "kernel32.dll" Alias "SetConsoleWindowInfo" (ByVal hConsoleOutput As LongPtr, ByVal bAbsolute As Long, lpConsoleWindow As SMALL_RECT) As Long
    Declare PtrSafe Function SetCurrentDirectory Lib "kernel32.dll" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
    Declare PtrSafe Function SetDefaultCommConfig Lib "kernel32.dll" Alias "SetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, ByVal dwSize As Long) As Long
    Declare PtrSafe Function SetEndOfFile Lib "kernel32.dll" Alias "SetEndOfFile" (ByVal hFile As LongPtr) As Long
    Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32.dll" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
    Declare PtrSafe Function SetErrorMode Lib "kernel32.dll" Alias "SetErrorMode" (ByVal wMode As Long) As Long
    Declare PtrSafe Function SetEvent Lib "kernel32.dll" Alias "SetEvent" (ByVal hEvent As LongPtr) As Long
    Declare PtrSafe Sub SetFileApisToANSI Lib "kernel32.dll" Alias "SetFileApisToANSI" ()
    Declare PtrSafe Sub SetFileApisToOEM Lib "kernel32.dll" Alias "SetFileApisToOEM" ()
    Declare PtrSafe Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
    Declare PtrSafe Function SetFilePointer Lib "kernel32.dll" Alias "SetFilePointer" (ByVal hFile As LongPtr, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
    Declare PtrSafe Function SetFileTime Lib "kernel32.dll" Alias "SetFileTime" (ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
    Declare PtrSafe Function SetLocaleInfo Lib "kernel32.dll" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
    Declare PtrSafe Function SetLocalTime Lib "kernel32.dll" Alias "SetLocalTime" (lpSystemTime As SYSTEMTIME) As Long
    Declare PtrSafe Function SetMailslotInfo Lib "kernel32.dll" Alias "SetMailslotInfo" (ByVal hMailslot As LongPtr, ByVal lReadTimeout As Long) As Long
    Declare PtrSafe Function SetNamedPipeHandleState Lib "kernel32.dll" Alias "SetNamedPipeHandleState" (ByVal hNamedPipe As LongPtr, lpMode As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long) As Long
    Declare PtrSafe Function SetPriorityClass Lib "kernel32.dll" Alias "SetPriorityClass" (ByVal hProcess As LongPtr, ByVal dwPriorityClass As Long) As Long
    Declare PtrSafe Function SetProcessShutdownParameters Lib "kernel32.dll" Alias "SetProcessShutdownParameters" (ByVal dwLevel As Long, ByVal dwFlags As Long) As Long
    Declare PtrSafe Function SetProcessWorkingSetSize Lib "kernel32.dll" Alias "SetProcessWorkingSetSize" (ByVal hProcess As LongPtr, ByVal dwMinimumWorkingSetSize As LongPtr, ByVal dwMaximumWorkingSetSize As LongPtr) As Long
    Declare PtrSafe Function SetStdHandle Lib "kernel32.dll" Alias "SetStdHandle" (ByVal nStdHandle As Long, ByVal nHandle As LongPtr) As Long
    Declare PtrSafe Function SetSystemPowerState Lib "kernel32.dll" Alias "SetSystemPowerState" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
    Declare PtrSafe Function SetSystemTime Lib "kernel32.dll" Alias "SetSystemTime" (lpSystemTime As SYSTEMTIME) As Long
    Declare PtrSafe Function SetSystemTimeAdjustment Lib "kernel32.dll" Alias "SetSystemTimeAdjustment" (ByVal dwTimeAdjustment As Long, ByVal bTimeAdjustmentDisabled As Long) As Long
    Declare PtrSafe Function SetTapeParameters Lib "kernel32.dll" Alias "SetTapeParameters" (ByVal hDevice As LongPtr, ByVal dwOperation As Long, lpTapeInformation As Any) As Long
    Declare PtrSafe Function SetTapePosition Lib "kernel32.dll" Alias "SetTapePosition" (ByVal hDevice As LongPtr, ByVal dwPositionMethod As Long, ByVal dwPartition As Long, ByVal dwOffsetLow As Long, ByVal dwOffsetHigh As Long, ByVal bimmediate As Long) As Long
    Declare PtrSafe Function SetThreadAffinityMask Lib "kernel32.dll" Alias "SetThreadAffinityMask" (ByVal hThread As LongPtr, ByVal dwThreadAffinityMask As LongPtr) As LongPtr
    Declare PtrSafe Function SetThreadContext Lib "kernel32.dll" Alias "SetThreadContext" (ByVal hThread As LongPtr, lpContext As CONTEXT) As Long
    Declare PtrSafe Function SetThreadLocale Lib "kernel32.dll" Alias "SetThreadLocale" (ByVal Locale As Long) As Long
    Declare PtrSafe Function SetThreadPriority Lib "kernel32.dll" Alias "SetThreadPriority" (ByVal hThread As LongPtr, ByVal nPriority As Long) As Long
    Declare PtrSafe Function SetTimeZoneInformation Lib "kernel32.dll" Alias "SetTimeZoneInformation" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    Declare PtrSafe Function SetUnhandledExceptionFilter Lib "kernel32.dll" Alias "SetUnhandledExceptionFilter" (ByVal lpTopLevelExceptionFilter As LongPtr) As LongPtr
    Declare PtrSafe Function SetupComm Lib "kernel32.dll" Alias "SetupComm" (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
    Declare PtrSafe Function SetVolumeLabel Lib "kernel32.dll" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
    Declare PtrSafe Function SuspendThread Lib "kernel32.dll" Alias "SuspendThread" (ByVal hThread As LongPtr) As Long
    Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32.dll" Alias "SystemTimeToFileTime" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
    Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32.dll" Alias "SystemTimeToTzSpecificLocalTime" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
    Declare PtrSafe Function TerminateProcess Lib "kernel32.dll" Alias "TerminateProcess" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
    Declare PtrSafe Function TerminateThread Lib "kernel32.dll" Alias "TerminateThread" (ByVal hThread As LongPtr, ByVal dwExitCode As Long) As Long
    Declare PtrSafe Function TlsAlloc Lib "kernel32.dll" Alias "TlsAlloc" () As Long
    Declare PtrSafe Function TlsFree Lib "kernel32.dll" Alias "TlsFree" (ByVal dwTlsIndex As Long) As Long
    Declare PtrSafe Function TlsGetValue Lib "kernel32.dll" Alias "TlsGetValue" (ByVal dwTlsIndex As Long) As LongPtr
    Declare PtrSafe Function TlsSetValue Lib "kernel32.dll" Alias "TlsSetValue" (ByVal dwTlsIndex As Long, lpTlsValue As Any) As Long
    Declare PtrSafe Function TransactNamedPipe Lib "kernel32.dll" Alias "TransactNamedPipe" (ByVal hNamedPipe As LongPtr, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
    Declare PtrSafe Function TransmitCommChar Lib "kernel32.dll" Alias "TransmitCommChar" (ByVal nCid As LongPtr, ByVal cChar As Byte) As Long
    Declare PtrSafe Function UnhandledExceptionFilter Lib "kernel32.dll" Alias "UnhandledExceptionFilter" (ExceptionInfo As EXCEPTION_POINTERS) As Long
    Declare PtrSafe Function UnlockFile Lib "kernel32.dll" Alias "UnlockFile" (ByVal hFile As LongPtr, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
    Declare PtrSafe Function UnlockFileEx Lib "kernel32.dll" Alias "UnlockFileEx" (ByVal hFile As LongPtr, ByVal dwReserved As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long, lpOverlapped As OVERLAPPED) As Long
    Declare PtrSafe Function UnmapViewOfFile Lib "kernel32.dll" Alias "UnmapViewOfFile" (lpBaseAddress As Any) As Long
    Declare PtrSafe Function UpdateResource Lib "kernel32.dll" Alias "UpdateResourceA" (ByVal hUpdate As LongPtr, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
    Declare PtrSafe Function VerLanguageName Lib "kernel32.dll" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
    Declare PtrSafe Function WaitCommEvent Lib "kernel32.dll" Alias "WaitCommEvent" (ByVal hFile As LongPtr, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long
    Declare PtrSafe Function WaitForMultipleObjects Lib "kernel32.dll" Alias "WaitForMultipleObjects" (ByVal nCount As Long, lpHandles As LongPtr, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
    Declare PtrSafe Function WaitForMultipleObjectsEx Lib "kernel32.dll" Alias "WaitForMultipleObjectsEx" (ByVal nCount As Long, lpHandles As LongPtr, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
    Declare PtrSafe Function WaitForSingleObject Lib "kernel32.dll" Alias "WaitForSingleObject" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
    Declare PtrSafe Function WaitForSingleObjectEx Lib "kernel32.dll" Alias "WaitForSingleObjectEx" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
    Declare PtrSafe Function WaitNamedPipe Lib "kernel32.dll" Alias "WaitNamedPipeA" (ByVal lpNamedPipeName As String, ByVal nTimeOut As Long) As Long
    Declare PtrSafe Function WideCharToMultiByte Lib "kernel32.dll" Alias "WideCharToMultiByte" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As LongPtr) As Long
    Declare PtrSafe Function WinExec Lib "kernel32.dll" Alias "WinExec" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
    Declare PtrSafe Function WriteConsole Lib "kernel32.dll" Alias "WriteConsoleA" (ByVal hConsoleOutput As LongPtr, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, ByVal lpReserved As LongPtr) As Long
    Declare PtrSafe Function WriteTapemark Lib "kernel32.dll" Alias "WriteTapemark" (ByVal hDevice As LongPtr, ByVal dwTapemarkType As Long, ByVal dwTapemarkCount As Long, ByVal bimmediate As Long) As Long

#Else

    Declare Function AddAtom Lib "kernel32" Alias "AddAtomA" ( ByVal lpString As String) As Integer
    Declare Function AllocConsole Lib "kernel32" ( ) As Long
    Declare Function BackupRead Lib "kernel32" ( ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As Any) As Long
    Declare Function BackupSeek Lib "kernel32" ( ByVal hFile As Long, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, lpdwLowByteSeeked As Long, lpdwHighByteSeeked As Long, lpContext As Long) As Long
    Declare Function BackupWrite Lib "kernel32" ( ByVal hFile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, lpContext As Long) As Long
    Declare Function Beep Lib "kernel32" ( ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
    Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" ( ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
    Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" ( ByVal lpDef As String, lpDCB As DCB) As Long
    Declare Function BuildCommDCBAndTimeouts Lib "kernel32" Alias "BuildCommDCBAndTimeoutsA" ( ByVal lpDef As String, lpDCB As DCB, lpCommTimeouts As COMMTIMEOUTS) As Long
    Declare Function ClearCommBreak Lib "kernel32" ( ByVal nCid As Long) As Long
    Declare Function ClearCommError Lib "kernel32" ( ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
    Declare Function CloseHandle Lib "kernel32" ( ByVal hObject As Long) As Long
    Declare Function CommConfigDialog Lib "kernel32" Alias "CommConfigDialogA" ( ByVal lpszName As String, ByVal hWnd As Long, lpCC As COMMCONFIG) As Boolean
    Declare Function CompareFileTime Lib "kernel32" ( lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
    Declare Function CompareString Lib "kernel32" Alias "CompareStringA" ( ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
    Declare Function ConnectNamedPipe Lib "kernel32" ( ByVal hNamedPipe As Long, lpOverlapped As OVERLAPPED) As Long
    Declare Function ContinueDebugEvent Lib "kernel32" ( ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long
    Declare Function ConvertDefaultLocale Lib "kernel32" ( ByVal Locale As Long) As Long
    Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" ( ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
    Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" ( ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
    Declare Function CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" ( ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
    Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" ( lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
    Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" ( ByVal hFile As Long, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
    Declare Function CreateIoCompletionPort Lib "kernel32" ( ByVal FileHandle As Long, ByVal ExistingCompletionPort As Long, ByVal CompletionKey As Long, ByVal NumberOfConcurrentThreads As Long) As Long
    Declare Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" ( ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
    Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" ( lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
    Declare Function CreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" ( ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
    Declare Function CreatePipe Lib "kernel32" ( phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
    Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Declare Function CreateProcessAsUser Lib "kernel32" Alias "CreateProcessAsUserA" ( ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As SECURITY_ATTRIBUTES, ByVal lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, ByVal lpStartupInfo As STARTUPINFO, ByVal lpProcessInformation As PROCESS_INFORMATION) As Long
    Declare Function CreateRemoteThread Lib "kernel32" ( ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
    Declare Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" ( lpSemaphoreAttributes As SECURITY_ATTRIBUTES, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As Long
    Declare Function CreateTapePartition Lib "kernel32" ( ByVal hDevice As Long, ByVal dwPartitionMethod As Long, ByVal dwCount As Long, ByVal dwSize As Long) As Long
    Declare Function CreateThread Lib "kernel32" ( lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
    Declare Function DefineDosDevice Lib "kernel32" Alias "DefineDosDeviceA" ( ByVal dwFlags As Long, ByVal lpDeviceName As String, ByVal lpTargetPath As String) As Long
    Declare Function DeleteAtom Lib "kernel32" ( ByVal nAtom As Integer) As Integer
    Declare Sub DeleteCriticalSection Lib "kernel32" ( lpCriticalSection As CRITICAL_SECTION)
    Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" ( ByVal lpFileName As String) As Long
    Declare Function DeviceIoControl Lib "kernel32" ( ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As OVERLAPPED) As Long
    Declare Function DisableThreadLibraryCalls Lib "kernel32" ( ByVal hLibModule As Long) As Boolean
    Declare Function DisconnectNamedPipe Lib "kernel32" ( ByVal hNamedPipe As Long) As Long
    Declare Function DosDateTimeToFileTime Lib "kernel32" ( ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
    Declare Function DuplicateHandle Lib "kernel32" ( ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
    Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" ( ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
    Declare Sub EnterCriticalSection Lib "kernel32" ( lpCriticalSection As CRITICAL_SECTION)
    Declare Function EnumCalendarInfo Lib "kernel32" Alias "EnumCalendarInfoA" ( ByVal lpCalInfoEnumProc As Long, ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long) As Boolean
    Declare Function EnumDateFormats Lib "kernel32" ( ByVal lpDateFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
    Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" ( ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
    Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" ( ByVal hModule As Long, ByVal lpType As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
    Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" ( ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
    Declare Function EnumSystemCodePages Lib "kernel32" ( ByVal lpCodePageEnumProc As Long, ByVal dwFlags As Long) As Long
    Declare Function EnumSystemLocales Lib "kernel32" ( ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Long
    Declare Function EnumTimeFormats Lib "kernel32" ( ByVal lpTimeFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
    Declare Function EraseTape Lib "kernel32" ( ByVal hDevice As Long, ByVal dwEraseType As Long, ByVal bimmediate As Long) As Long
    Declare Function EscapeCommFunction Lib "kernel32" ( ByVal nCid As Long, ByVal nFunc As Long) As Long
    Declare Sub ExitProcess Lib "kernel32" ( ByVal uExitCode As Long)
    Declare Sub ExitThread Lib "kernel32" ( ByVal dwExitCode As Long)
    Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" ( ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
    Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" ( ByVal uAction As Long, ByVal lpMessageText As String)
    Declare Sub FatalExit Lib "kernel32" ( ByVal code As Long)
    Declare Function FileTimeToDosDateTime Lib "kernel32" ( lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
    Declare Function FileTimeToLocalFileTime Lib "kernel32" ( lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
    Declare Function FileTimeToSystemTime Lib "kernel32" ( lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
    Declare Function FillConsoleOutputAttribute Lib "kernel32" ( ByVal hConsoleOutput As Long, ByVal wAttribute As Long, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
    Declare Function FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" ( ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
    Declare Function FindAtom Lib "kernel32" Alias "FindAtomA" ( ByVal lpString As String) As Integer
    Declare Function FindClose Lib "kernel32" ( ByVal hFindFile As Long) As Long
    Declare Function FindCloseChangeNotification Lib "kernel32" ( ByVal hChangeHandle As Long) As Long
    Declare Function FindFirstChangeNotification Lib "kernel32" Alias "FindFirstChangeNotificationA" ( ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
    Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" ( ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
    Declare Function FindNextChangeNotification Lib "kernel32" ( ByVal hChangeHandle As Long) As Long
    Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" ( ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    Declare Function FindResource Lib "kernel32" Alias "FindResourceA" ( ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
    Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" ( ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long) As Long
    Declare Function FreeConsole Lib "kernel32" () As Long
    Declare Function FreeEnvironmentStrings Lib "kernel32" Alias "FreeEnvironmentStringsA" ( ByVal lpsz As String) As Boolean
    Declare Function FreeLibrary Lib "kernel32" ( ByVal hLibModule As Long) As Long
    Declare Sub FreeLibraryAndExitThread Lib "kernel32" ( ByVal hLibModule As Long, ByVal dwExitCode As Long)
    Declare Function FreeResource Lib "kernel32" ( ByVal hResData As Long) As Boolean
    Declare Function GenerateConsoleCtrlEvent Lib "kernel32" ( ByVal dwCtrlEvent As Long, ByVal dwProcessGroupId As Long) As Long
    Declare Function GetACP Lib "kernel32" () As Long
    Declare Function GetAtomName Lib "kernel32" Alias "GetAtomNameA" ( ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeA" ( ByVal lpApplicationName As String, lpBinaryType As Long) As Long
    Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String
    Declare Function GetCommConfig Lib "kernel32" ( ByVal hCommDev As Long, lpCC As COMMCONFIG, lpdwSize As Long) As Boolean
    Declare Function GetCommMask Lib "kernel32" ( ByVal hFile As Long, lpEvtMask As Long) As Long
    Declare Function GetCommModemStatus Lib "kernel32" ( ByVal hFile As Long, lpModemStat As Long) As Long
    Declare Function GetCommProperties Lib "kernel32" ( ByVal hFile As Long, lpCommProp As COMMPROP) As Long
    Declare Function GetCommState Lib "kernel32" ( ByVal nCid As Long, lpDCB As DCB) As Long
    Declare Function GetCommTimeouts Lib "kernel32" ( ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
    Declare Function GetCompressedFileSize Lib "kernel32" Alias "GetCompressedFileSizeA" ( ByVal lpFileName As String, lpFileSizeHigh As Long) As Long
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" ( ByVal lpBuffer As String, nSize As Long) As Long
    Declare Function GetConsoleCP Lib "kernel32" () As Long
    Declare Function GetConsoleCursorInfo Lib "kernel32" ( ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
    Declare Function GetConsoleMode Lib "kernel32" ( ByVal hConsoleHandle As Long, lpMode As Long) As Long
    Declare Function GetConsoleOutputCP Lib "kernel32" () As Long
    Declare Function GetConsoleScreenBufferInfo Lib "kernel32" ( ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
    Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" ( ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
    Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" ( ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Boolean
    Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" ( ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
    Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" ( ByVal nDrive As String) As Long
    Declare Function GetEnvironmentStrings Lib "kernel32" Alias "GetEnvironmentStringsA" ( ) As String
    Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" ( ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare Function GetExitCodeProcess Lib "kernel32" ( ByVal hProcess As Long, lpExitCode As Long) As Long
    Declare Function GetExitCodeThread Lib "kernel32" ( ByVal hThread As Long, lpExitCode As Long) As Long
    Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" ( ByVal lpFileName As String) As Long
    Declare Function GetFileInformationByHandle Lib "kernel32" ( ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
    Declare Function GetFileSize Lib "kernel32" ( ByVal hFile As Long, lpFileSizeHigh As Long) As Long
    Declare Function GetFileTime Lib "kernel32" ( ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
    Declare Function GetFileType Lib "kernel32" ( ByVal hFile As Long) As Long
    Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" ( ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
    Declare Function GetHandleInformation Lib "kernel32" ( ByVal hObject As Long, lpdwFlags As Long) As Boolean
    Declare Function GetLargestConsoleWindowSize Lib "kernel32" ( ByVal hConsoleOutput As Long) As COORD
    Declare Function GetLastError Lib "kernel32" ( ) As Long
    Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" ( ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Declare Sub GetLocalTime Lib "kernel32" ( lpSystemTime As SYSTEMTIME)
    Declare Function GetLogicalDrives Lib "kernel32" ( ) As Long
    Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" ( ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Declare Function GetMailslotInfo Lib "kernel32" ( ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
    Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" ( ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
    Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" ( ByVal lpModuleName As String) As Long
    Declare Function GetOEMCP Lib "kernel32" ( ) As Long
    Declare Function GetOverlappedResult Lib "kernel32" ( ByVal hFile As Long, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
    Declare Function GetPriorityClass Lib "kernel32" ( ByVal hProcess As Long) As Long
    Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" ( ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
    Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" ( ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function GetProcAddress Lib "kernel32" ( ByVal hModule As Long, ByVal lpProcName As String) As Long
    Declare Function GetProcessAffinityMask Lib "kernel32" ( ByVal hProcess As Long, lpProcessAffinityMask As Long, SystemAffinityMask As Long) As Long
    Declare Function GetProcessHeap Lib "kernel32" () As Long
    Declare Function GetProcessHeaps Lib "kernel32" ( ByVal NumberOfHeaps As Long, ProcessHeaps As Long) As Long
    Declare Function GetProcessShutdownParameters Lib "kernel32" ( lpdwLevel As Long, lpdwFlags As Long) As Long
    Declare Function GetProcessTimes Lib "kernel32" ( ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
    Declare Function GetProcessWorkingSetSize Lib "kernel32" ( ByVal hProcess As Long, lpMinimumWorkingSetSize As Long, lpMaximumWorkingSetSize As Long) As Boolean
    Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" ( ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
    Declare Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" ( ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
    Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" ( ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
    Declare Function GetQueuedCompletionStatus Lib "kernel32" ( ByVal CompletionPort As Long, lpNumberOfBytesTransferred As Long, lpCompletionKey As Long, lpOverlapped As Long, ByVal dwMilliseconds As Long) As Boolean
    Declare Function GetShortPathName Lib "kernel32" ( ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
    Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" ( lpStartupInfo As STARTUPINFO)
    Declare Function GetStdHandle Lib "kernel32" ( ByVal nStdHandle As Long) As Long
    Declare Function GetStringTypeA Lib "kernel32" ( ByVal lcid As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Long) As Long
    Declare Function GetStringTypeEx Lib "kernel32" Alias "GetStringTypeExA" ( ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Boolean
    Declare Function GetStringTypeW Lib "kernel32" ( ByVal dwInfoType As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, lpCharType As Integer) As Boolean
    Declare Function GetSystemPowerStatus Lib "kernel32" ( lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
    Declare Sub GetSystemTime Lib "kernel32" ( lpSystemTime As SYSTEMTIME)
    Declare Function GetSystemTimeAdjustment Lib "kernel32" ( lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Boolean) As Boolean
    Declare Function GetTapeParameters Lib "kernel32" ( ByVal hDevice As Long, ByVal dwOperation As Long, lpdwSize As Long, lpTapeInformation As Any) As Long
    Declare Function GetTapePosition Lib "kernel32" ( ByVal hDevice As Long, ByVal dwPositionType As Long, lpdwPartition As Long, lpdwOffsetLow As Long, lpdwOffsetHigh As Long) As Long
    Declare Function GetTapeStatus Lib "kernel32" ( ByVal hDevice As Long) As Long
    Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" ( ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
    Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" ( ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Declare Function GetThreadContext Lib "kernel32" ( ByVal hThread As Long, lpContext As CONTEXT) As Long
    Declare Function GetThreadLocale Lib "kernel32" ( ) As Long
    Declare Function GetThreadPriority Lib "kernel32" ( ByVal hThread As Long) As Long
    Declare Function GetThreadSelectorEntry Lib "kernel32" ( ByVal hThread As Long, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long
    Declare Function GetThreadTimes Lib "kernel32" ( ByVal hThread As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
    Declare Function GetTickCount Lib "kernel32" () As Long
    Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" ( ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
    Declare Function GetTimeZoneInformation Lib "kernel32" ( lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
    Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
    Declare Function GetVersion Lib "kernel32" () As Long
    Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( ByVal lpVersionInformation As OSVERSIONINFO) As Long
    Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" ( ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
    Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" ( ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" ( ByVal lpString As String) As Integer
    Declare Function GlobalAlloc Lib "kernel32" ( ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Declare Function GlobalCompact Lib "kernel32" ( ByVal dwMinFree As Long) As Long
    Declare Function GlobalDeleteAtom Lib "kernel32" ( ByVal nAtom As Integer) As Integer
    Declare Function GlobalFindAtom Lib "kernel32" Alias "GlobalFindAtomA" ( ByVal lpString As String) As Integer
    Declare Sub GlobalFix Lib "kernel32" ( ByVal hMem As Long)
    Declare Function GlobalFlags Lib "kernel32" ( ByVal hMem As Long) As Long
    Declare Function GlobalFree Lib "kernel32" ( ByVal hMem As Long) As Long
    Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" ( ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare Function HeapAlloc Lib "kernel32" ( ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
    Declare Function HeapCompact Lib "kernel32" ( ByVal hHeap As Long, ByVal dwFlags As Long) As Long
    Declare Function HeapCreate Lib "kernel32" ( ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
    Declare Function HeapDestroy Lib "kernel32" ( ByVal hHeap As Long) As Long
    Declare Function HeapFree Lib "kernel32" ( ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
    Declare Function HeapLock Lib "kernel32" ( ByVal hHeap As Long) As Long
    Declare Function HeapReAlloc Lib "kernel32" ( ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
    Declare Function HeapSize Lib "kernel32" ( ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
    Declare Function HeapUnlock Lib "kernel32" ( ByVal hHeap As Long) As Long
    Declare Function HeapValidate Lib "kernel32" ( ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
    Declare Function hread Lib "kernel32" Alias "_hread" ( ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
    Declare Function hwrite Lib "kernel32" Alias "_hwrite" ( ByVal hFile As Long, ByVal lpBuffer As String, ByVal lBytes As Long) As Long
    Declare Function ImpersonateLoggedOnUser Lib "kernel32" ( ByVal hToken As Long) As Long
    Declare Function InitAtomTable Lib "kernel32" ( ByVal nSize As Long) As Long
    Declare Sub InitializeCriticalSection Lib "kernel32" ( lpCriticalSection As CRITICAL_SECTION)
    Declare Function InterlockedDecrement Lib "kernel32" ( lpAddend As Long) As Long
    Declare Function InterlockedExchange Lib "kernel32" ( Target As Long, ByVal Value As Long) As Long
    Declare Function InterlockedIncrement Lib "kernel32" ( lpAddend As Long) As Long
    Declare Function IsBadCodePtr Lib "kernel32" ( ByVal lpfn As Long) As Boolean
    Declare Function IsBadHugeReadPtr Lib "kernel32" ( lp As Any, ByVal ucb As Long) As Long
    Declare Function IsBadHugeWritePtr Lib "kernel32" ( lp As Any, ByVal ucb As Long) As Long
    Declare Function IsBadReadPtr Lib "kernel32" ( lp As Any, ByVal ucb As Long) As Long
    Declare Function IsBadStringPtr Lib "kernel32" Alias "IsBadStringPtrA" ( ByVal lpsz As String, ByVal ucchMax As Long) As Long
    Declare Function IsBadWritePtr Lib "kernel32" ( lp As Any, ByVal ucb As Long) As Long
    Declare Function IsDBCSLeadByte Lib "kernel32" ( ByVal bTestChar As Byte) As Long
    Declare Function lclose Lib "kernel32" Alias "_lclose" ( ByVal hFile As Long) As Long
    Declare Function LCMapString Lib "kernel32" Alias "LCMapStringA" ( ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As String, ByVal cchSrc As Long, ByVal lpDestStr As String, ByVal cchDest As Long) As Long
    Declare Function lcreat Lib "kernel32" Alias "_lcreat" ( ByVal lpPathName As String, ByVal iAttribute As Long) As Long
    Declare Sub LeaveCriticalSection Lib "kernel32" ( lpCriticalSection As CRITICAL_SECTION)
    Declare Function llseek Lib "kernel32" Alias "_llseek" ( ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
    Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( ByVal lpLibFileName As String) As Long
    Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" ( ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
    Declare Function LoadModule Lib "kernel32" ( ByVal lpModuleName As String, lpParameterBlock As Any) As Long
    Declare Function LoadResource Lib "kernel32" ( ByVal hInstance As Long, ByVal hResInfo As Long) As Long
    Declare Function LocalAlloc Lib "kernel32" ( ByVal wFlags As Long, ByVal wBytes As Long) As Long
    Declare Function LocalCompact Lib "kernel32" ( ByVal uMinFree As Long) As Long
    Declare Function LocalFileTimeToFileTime Lib "kernel32" ( lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
    Declare Function LocalFlags Lib "kernel32" ( ByVal hMem As Long) As Long
    Declare Function LocalFree Lib "kernel32" ( ByVal hMem As Long) As Long
    Declare Function LocalHandle Lib "kernel32" ( wMem As Any) As Long
    Declare Function LocalLock Lib "kernel32" ( ByVal hMem As Long) As Long
    Declare Function LocalReAlloc Lib "kernel32" ( ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long
    Declare Function LocalShrink Lib "kernel32" ( ByVal hMem As Long, ByVal cbNewSize As Long) As Long
    Declare Function LocalSize Lib "kernel32" ( ByVal hMem As Long) As Long
    Declare Function LocalUnlock Lib "kernel32" ( ByVal hMem As Long) As Long
    Declare Function LockFile Lib "kernel32" ( ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
    Declare Function LockFileEx Lib "kernel32" ( ByVal hFile As Long, ByVal dwFlags As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long, lpOverlapped As OVERLAPPED) As Long
    Declare Function LockResource Lib "kernel32" ( ByVal hResData As Long) As Long
    Declare Function LogonUser Lib "kernel32" Alias " LogonUserA" ( ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
    Declare Function lopen Lib "kernel32" Alias "_lopen" ( ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
    Declare Function lread Lib "kernel32" Alias "_lread" ( ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
    Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" ( ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpA" ( ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" ( ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" ( ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
    Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" ( ByVal lpString As String) As Long
    Declare Function lwrite Lib "kernel32" Alias "_lwrite" ( ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long) As Long
    Declare Function MapViewOfFile Lib "kernel32" ( ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
    Declare Function MapViewOfFileEx Lib "kernel32" ( ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long, lpBaseAddress As Any) As Long
    Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" ( ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
    Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" ( ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
    Declare Function MulDiv Lib "kernel32" ( ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Declare Function MultiByteToWideChar Lib "kernel32" ( ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
    Declare Function ObjectOpenAuditAlarm Lib "kernel32" Alias "ObjectOpenAuditAlarmA" ( ByVal SubsystemName As String, HandleId As Any, ByVal ObjectTypeName As String, ByVal ObjectName As String, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal ClientToken As Long, ByVal DesiredAccess As Long, ByVal GrantedAccess As Long, Privileges As PRIVILEGE_SET, ByVal ObjectCreation As Long, ByVal AccessGranted As Long, ByVal GenerateOnClose As Long) As Long
    Declare Function OpenEvent Lib "kernel32" Alias "OpenEventA" ( ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
    Declare Function OpenFile Lib "kernel32" ( ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
    Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" ( ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
    Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" ( ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
    Declare Function OpenProcess Lib "kernel32" ( ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Declare Function OpenSemaphore Lib "kernel32" Alias "OpenSemaphoreA" ( ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
    Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" ( ByVal lpOutputString As String)
    Declare Function PeekNamedPipe Lib "kernel32" ( ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
    Declare Function PrepareTape Lib "kernel32" ( ByVal hDevice As Long, ByVal dwOperation As Long, ByVal bimmediate As Long) As Long
    Declare Function PulseEvent Lib "kernel32" ( ByVal hEvent As Long) As Long
    Declare Function ReadConsoleOutputAttribute Lib "kernel32" ( ByVal hConsoleOutput As Long, lpAttribute As Long, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfAttrsRead As Long) As Long
    Declare Function ReadConsoleOutputCharacter Lib "kernel32" Alias "ReadConsoleOutputCharacterA" ( ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwReadCoord As COORD, lpNumberOfCharsRead As Long) As Long
    Declare Function ReadFile Lib "kernel32" ( ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
    Declare Function ReadFileEx Lib "kernel32" ( ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Boolean
    Declare Function ReadProcessMemory Lib "kernel32" ( ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
    Declare Function ReleaseMutex Lib "kernel32" ( ByVal hMutex As Long) As Long
    Declare Function ReleaseSemaphore Lib "kernel32" ( ByVal hSemaphore As Long, ByVal lReleaseCount As Long, lpPreviousCount As Long) As Long
    Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" ( ByVal lpPathName As String) As Long
    Declare Function ResetEvent Lib "kernel32" ( ByVal hEvent As Long) As Long
    Declare Function ResumeThread Lib "kernel32" ( ByVal hThread As Long) As Long
    Declare Function SetCommBreak Lib "kernel32" ( ByVal nCid As Long) As Long
    Declare Function SetCommConfig Lib "kernel32" ( ByVal hCommDev As Long, lpCC As COMMCONFIG, ByVal dwSize As Long) As Boolean
    Declare Function SetCommMask Lib "kernel32" ( ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
    Declare Function SetCommState Lib "kernel32" ( ByVal hCommDev As Long, lpDCB As DCB) As Long
    Declare Function SetCommTimeouts Lib "kernel32" ( ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
    Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" ( ByVal lpComputerName As String) As Long
    Declare Function SetConsoleActiveScreenBuffer Lib "kernel32" ( ByVal hConsoleOutput As Long) As Long
    Declare Function SetConsoleCP Lib "kernel32" ( ByVal wCodePageID As Long) As Long
    Declare Function SetConsoleCtrlHandler Lib "kernel32" ( ByVal HandlerRoutine As Long, ByVal Add As Long) As Long
    Declare Function SetConsoleCursorInfo Lib "kernel32" ( ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
    Declare Function SetConsoleCursorPosition Lib "kernel32" ( ByVal hConsoleOutput As Long, dwCursorPosition As COORD) As Long
    Declare Function SetConsoleMode Lib "kernel32" ( ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long
    Declare Function SetConsoleOutputCP Lib "kernel32" ( ByVal wCodePageID As Long) As Long
    Declare Function SetConsoleScreenBufferSize Lib "kernel32" ( ByVal hConsoleOutput As Long, dwSize As COORD) As Long
    Declare Function SetConsoleTextAttribute Lib "kernel32" ( ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
    Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" ( ByVal lpConsoleTitle As String) As Long
    Declare Function SetConsoleWindowInfo Lib "kernel32" ( ByVal hConsoleOutput As Long, ByVal bAbsolute As Long, lpConsoleWindow As SMALL_RECT) As Long
    Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" ( ByVal lpPathName As String) As Long
    Declare Function SetDefaultCommConfig Lib "kernel32" Alias "SetDefaultCommConfigA" ( ByVal lpszName As String, lpCC As COMMCONFIG, ByVal dwSize As Long) As Boolean
    Declare Function SetEndOfFile Lib "kernel32" ( ByVal hFile As Long) As Long
    Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" ( ByVal lpName As String, ByVal lpValue As String) As Long
    Declare Function SetErrorMode Lib "kernel32" ( ByVal wMode As Long) As Long
    Declare Function SetEvent Lib "kernel32" ( ByVal hEvent As Long) As Long
    Declare Sub SetFileApisToANSI Lib "kernel32" ( )
    Declare Sub SetFileApisToOEM Lib "kernel32" ( )
    Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" ( ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
    Declare Function SetFilePointer Lib "kernel32" ( ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
    Declare Function SetFileTime Lib "kernel32" ( ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
    Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" ( ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
    Declare Function SetLocalTime Lib "kernel32" ( lpSystemTime As SYSTEMTIME) As Long
    Declare Function SetMailslotInfo Lib "kernel32" ( ByVal hMailslot As Long, ByVal lReadTimeout As Long) As Long
    Declare Function SetNamedPipeHandleState Lib "kernel32" ( ByVal hNamedPipe As Long, lpMode As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long) As Long
    Declare Function SetPriorityClass Lib "kernel32" ( ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
    Declare Function SetProcessShutdownParameters Lib "kernel32" ( ByVal dwLevel As Long, ByVal dwFlags As Long) As Long
    Declare Function SetProcessWorkingSetSize Lib "kernel32" ( ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Boolean
    Declare Function SetStdHandle Lib "kernel32" ( ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
    Declare Function SetSystemPowerState Lib "kernel32" ( ByVal fSuspend As Long, ByVal fForce As Long) As Long
    Declare Function SetSystemTime Lib "kernel32" ( lpSystemTime As SYSTEMTIME) As Long
    Declare Function SetSystemTimeAdjustment Lib "kernel32" ( ByVal dwTimeAdjustment As Long, ByVal bTimeAdjustmentDisabled As Boolean) As Boolean
    Declare Function SetTapeParameters Lib "kernel32" ( ByVal hDevice As Long, ByVal dwOperation As Long, lpTapeInformation As Any) As Long
    Declare Function SetTapePosition Lib "kernel32" ( ByVal hDevice As Long, ByVal dwPositionMethod As Long, ByVal dwPartition As Long, ByVal dwOffsetLow As Long, ByVal dwOffsetHigh As Long, ByVal bimmediate As Long) As Long
    Declare Function SetThreadAffinityMask Lib "kernel32" ( ByVal hThread As Long, ByVal dwThreadAffinityMask As Long) As Long
    Declare Function SetThreadContext Lib "kernel32" ( ByVal hThread As Long, lpContext As CONTEXT) As Long
    Declare Function SetThreadLocale Lib "kernel32" ( ByVal Locale As Long) As Long
    Declare Function SetThreadPriority Lib "kernel32" ( ByVal hThread As Long, ByVal nPriority As Long) As Long
    Declare Function SetTimeZoneInformation Lib "kernel32" ( lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    Declare Function SetUnhandledExceptionFilter Lib "kernel32" ( ByVal lpTopLevelExceptionFilter As Long) As Long
    Declare Function SetupComm Lib "kernel32" ( ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
    Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" ( ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
    Declare Function SuspendThread Lib "kernel32" ( ByVal hThread As Long) As Long
    Declare Function SystemTimeToFileTime Lib "kernel32" ( lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
    Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" ( lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Boolean
    Declare Function TerminateProcess Lib "kernel32" ( ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    Declare Function TerminateThread Lib "kernel32" ( ByVal hThread As Long, ByVal dwExitCode As Long) As Long
    Declare Function TlsAlloc Lib "kernel32" () As Long
    Declare Function TlsFree Lib "kernel32" ( ByVal dwTlsIndex As Long) As Long
    Declare Function TlsGetValue Lib "kernel32" ( ByVal dwTlsIndex As Long) As Long
    Declare Function TlsSetValue Lib "kernel32" ( ByVal dwTlsIndex As Long, lpTlsValue As Any) As Long
    Declare Function TransactNamedPipe Lib "kernel32" ( ByVal hNamedPipe As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
    Declare Function TransmitCommChar Lib "kernel32" ( ByVal nCid As Long, ByVal cChar As Byte) As Long
    Declare Function UnhandledExceptionFilter Lib "kernel32" ( ExceptionInfo As EXCEPTION_POINTERS) As Long
    Declare Function UnlockFile Lib "kernel32" ( ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
    Declare Function UnlockFileEx Lib "kernel32" ( ByVal hFile As Long, ByVal dwReserved As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long, lpOverlapped As OVERLAPPED) As Long
    Declare Function UnmapViewOfFile Lib "kernel32" ( lpBaseAddress As Any) As Long
    Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" ( ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
    Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" ( ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
    Declare Function WaitCommEvent Lib "kernel32" ( ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long
    Declare Function WaitForMultipleObjects Lib "kernel32" ( ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
    Declare Function WaitForMultipleObjectsEx Lib "kernel32" ( ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
    Declare Function WaitForSingleObject Lib "kernel32" ( ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Declare Function WaitForSingleObjectEx Lib "kernel32" ( ByVal hHandle As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
    Declare Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" ( ByVal lpNamedPipeName As String, ByVal nTimeOut As Long) As Long
    Declare Function WideCharToMultiByte Lib "kernel32" ( ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
    Declare Function WinExec Lib "kernel32" ( ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
    Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" ( ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
    Declare Function WriteTapemark Lib "kernel32" ( ByVal hDevice As Long, ByVal dwTapemarkType As Long, ByVal dwTapemarkCount As Long, ByVal bimmediate As Long) As Long

#End If