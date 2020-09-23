Attribute VB_Name = "ModlLoadProcess"
' #####################################################################################
' #####################################################################################
' Description:      Wraps the Windows API Functions necessary to retrieve the Process
'                     information from windows.
'
'   Needs:          PSAPI.dll to be installed and registered in the System Directory.
'
'
' #####################################################################################
' #####################################################################################
Option Explicit

'***************************************************************************************
'   API Declares
'***************************************************************************************
'Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, _
'                                                        lppe As PROCESSENTRY32) As Long
'Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, _
'                                                        lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, _
                                                        ByVal bInheritHandle As Long, _
                                                        ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, _
                                                        ByVal cb As Long, _
                                                        ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, _
                                                        ByVal hModule As Long, _
                                                        ByVal ModuleName As String, _
                                                        ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, _
                                                        ByRef lphModule As Long, _
                                                        ByVal cb As Long, _
                                                        ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, _
                                                        ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
                                                        ByVal uExitCode As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'***************************************************************************************
'   Types Used to Retrieve Information From Windows
'***************************************************************************************
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

'***************************************************************************************
'   Constants to set buffer sizes, rights, and determine OS Version
'***************************************************************************************
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const MAX_PATH = 260

'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000

Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const hNull = 0

Public Const TH32CS_SNAPheaplist = &H1
Public Const TH32CS_SNAPthread = &H4
Public Const TH32CS_SNAPmodule = &H8
Public Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule

'define PROCESSENTRY32 structure
Public Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long
 th32DefaultHeapID As Long
 th32ModuleID As Long
 cntThreads As Long
 th32ParentProcessID As Long
 pcPriClassBase As Long
 dwFlags As Long
 szExeFile As String * MAX_PATH
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'Used to Get the Error Message
Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0
Const SUBLANG_DEFAULT = &H1

'Used to determine what OS Version
Public Const WINNT As Integer = 2
Public Const WIN98 As Integer = 1

Public NoNewCon As Boolean
Public ForceRefresh As Boolean

Public ProgEntries As Long


'Private Type ID_
'Filename As String
'ProcessNumber As Long
'End Type
'Public Id(0 To 2000) As ID_
Public Sub LoadNTProcess()
  Dim cb As Long
  Dim cbNeeded As Long
  Dim NumElements As Long
  Dim ProcessIDs() As Long
  Dim cbNeeded2 As Long
  Dim NumElements2 As Long
  Dim Modules(1 To 200) As Long
  Dim lRet As Long
  Dim ModuleName As String
  Dim nSize As Long
  Dim hProcess As Long
  Dim i As Long
  Dim Y As Long
  Dim q As Long
  Dim Huh As Boolean
  Dim NewCnt As Long
  Dim OldCnt As Long
  Dim MessageAnswer

    'Get the array containing the process id's for each process object
    cb = 8
    cbNeeded = 96
    Do While cb <= cbNeeded
        cb = cb * 2
        ReDim ProcessIDs(cb / 4) As Long
        lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
    DoEvents
    Loop
         
    NumElements = cbNeeded / 4
    For i = 1 To NumElements
      'Get a handle to the Process
         hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
      'Got a Process handle
         If hProcess <> 0 Then
           'Get an array of the module handles for the specified process
             lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
             
           'If the Module Array is retrieved, Get the ModuleFileName
            If lRet <> 0 Then
                ModuleName = Space(MAX_PATH)
                nSize = 500
                lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
                                     
                                     
                If CBool(InStr(1, (Left(ModuleName, lRet)), "", vbTextCompare)) Then
                'Go through all connections and find it

                For Y = 0 To GetEntryCount
                
                If Connection(Y).State <> 2 Then
                    If Connection(Y).ProcessID = ProcessIDs(i) Then
                    Connection(Y).FileName = Left(ModuleName, lRet)
                    End If
                End If
                
                Next Y
                
                End If
            End If
        End If
               
        'Close the handle to the process
        lRet = CloseHandle(hProcess)
    Next
    
   
End Sub
Public Function GetProcessName(iID As Long) As String
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim theloop As Long

snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0) 'get snapshot handle
proc.dwSize = Len(proc)
theloop = ProcessFirst(snap, proc)       'first process and return value

i = 0

While theloop <> 0      'next process
If iID = proc.th32ProcessID Then
GetProcessName = proc.szExeFile  'add process ID to listview
End If
i = i + 1
theloop = ProcessNext(snap, proc)
Wend
CloseHandle snap       'close snapshot handle
End Function

'***************************************************************************************
'   Public Functions
'***************************************************************************************
'   Function Name:  getVersion
'
'   Description:    Gets the OS Version (NT or 98).  See Constants above for value
'
'   Inputs:         NONE
'   Returns:        An integer value corresponding to the OS Version
'
'***************************************************************************************
Public Function getVersion() As Integer
  Dim udtOSInfo As OSVERSIONINFO
  Dim intRetVal As Integer
         
  'Initialize the type's buffer sizes
    With udtOSInfo
        .dwOSVersionInfoSize = 148
        .szCSDVersion = Space$(128)
    End With
    
  'Make an API Call to Retrieve the OSVersion info
    intRetVal = GetVersionExA(udtOSInfo)
  
  'Set the return value
    getVersion = udtOSInfo.dwPlatformId
End Function

'***************************************************************************************
'   Sub Name:       KillProcessByID
'
'   Description:    Given a ProcessID, this function will get its Windows Handle and
'                       Terminate the process.
'
'   Inputs:         p_lngProcessId -->  The processid of the process to terminate.
'   Returns:        NONE
'
'***************************************************************************************
Public Function KillProcessById(p_lngProcessId As Long) As Boolean
  Dim lnghProcess As Long
  Dim lngReturn As Long
    
    lnghProcess = OpenProcess(1&, -1&, p_lngProcessId)
    lngReturn = TerminateProcess(lnghProcess, 0&)
    
    If lngReturn = 0 Then
        RetrieveError
        KillProcessById = False
        Exit Function
    End If
    
    KillProcessById = True
End Function

'***************************************************************************************
'   Sub Name:       RetrieveError
'
'   Description:    Called when the process can't terminate.  Used to retrieve the error
'                     generated during the terminate attempt.
'
'   Inputs:         NONE
'   Returns:        NONE
'
'***************************************************************************************
Private Sub RetrieveError()
  Dim strBuffer As String
    
    'Create a string buffer
    strBuffer = Space(200)
    
    'Format the message string
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, strBuffer, 200, ByVal 0&
    'Show the message
    'MsgBox strBuffer
End Sub
