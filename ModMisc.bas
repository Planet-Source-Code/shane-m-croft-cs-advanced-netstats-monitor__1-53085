Attribute VB_Name = "ModMisc"
Option Explicit

'Icon Sizes in pixels
Public Const LARGE_ICON As Integer = 32
Public Const SMALL_ICON As Integer = 16

Public Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Public Const SHGFI_LARGEICON = &H0       'Large icon
Public Const SHGFI_SMALLICON = &H1       'Small icon
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400

Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Public Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal X&, ByVal Y&, ByVal Flags&) As Long

'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Public ShInfo As SHFILEINFO

'***********************************************************************************************

Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hwnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Private Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hwnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2

'***********************************************************************************************
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function InternetGetConnectedState _
    Lib "wininet.dll" (ByRef lpdwFlags As Long, _
    ByVal dwReserved As Long) As Long
    'Local system uses a modem to connect to
    '     the Internet.
    Public Const INTERNET_CONNECTION_MODEM As Long = &H1
    'Local system uses a LAN to connect to t
    '     he Internet.
    Public Const INTERNET_CONNECTION_LAN As Long = &H2
    'Local system uses a proxy server to con
    '     nect to the Internet.
    Public Const INTERNET_CONNECTION_PROXY As Long = &H4
    'No longer used.
    Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
    Public Const INTERNET_RAS_INSTALLED As Long = &H10
    Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20
    Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
    'InternetGetConnectedState wrapper funct
    '     ions
'************************************************************************************************
Public ShowTrafficInBytes As Boolean

'***************************************List View************************************************
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1
'************************************************************************************************

Public Const INVALID_HANDLE_VALUE As Long = -1

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type
   
Public Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
'***************************************File Information******************************************

Public Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersion As Long     'e.g. 0x00000042 = "0.42"
   dwFileVersionMS As Long    'e.g. 0x00030075 = "3.75"
   dwFileVersionLS As Long    'e.g. 0x00000031 = "0.31"
   dwProductVersionMS As Long 'e.g. 0x00030010 = "3.10"
   dwProductVersionLS As Long 'e.g. 0x00000031 = "0.31"
   dwFileFlagsMask As Long    'e.g. 0x3F for version "0.42"
   dwFileFlags As Long        'e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long           'e.g. VOS_DOS_WINDOWS16
   dwFileType As Long         'e.g. VFT_DRIVER
   dwFileSubtype As Long      'e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long       'e.g. 0
   dwFileDateLS As Long       'e.g. 0
End Type

Public Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Public Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Public Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lplpBuffer As Any, nVerSize As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)


Public Const MAXDWORD As Long = &HFFFFFFFF
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Public Type FILE_PARAMS  'my custom type for passing info
   bRecurse As Boolean   'var not used in this demo
   bList As Boolean
   bFound As Boolean     'var not used in this demo
   sFileRoot As String
   sFileNameExt As String
   sResult As String     'var not used in this demo
   nFileCount As Long    'var not used in this demo
   nFileSize As Double   'var not used in this demo
End Type

Public Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Public Declare Function lstrlenA Lib "kernel32" _
  (ByVal Ptr As Any) As Long
  '*************************************Load Explorer*********************************************

Public Const CREATE_NEW_CONSOLE As Long = &H10
Public Const NORMAL_PRIORITY_CLASS As Long = &H20
Public Const INFINITE As Long = -1
Public Const STARTF_USESHOWWINDOW As Long = &H1
Public Const SW_SHOWNORMAL As Long = 1

Public Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Public Const ERROR_FILE_NOT_FOUND As Long = 2
Public Const ERROR_PATH_NOT_FOUND As Long = 3
Public Const ERROR_FILE_SUCCESS As Long = 32 'my constant
Public Const ERROR_BAD_FORMAT As Long = 11

Public Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Public Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadID As Long
End Type

Public Declare Function CreateProcess Lib "kernel32" _
   Alias "CreateProcessA" _
  (ByVal lpAppName As String, _
   ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, _
   ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, _
   ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long
     
'Public Declare Function CloseHandle Lib "kernel32" _
'  (ByVal hObject As Long) As Long

Public Declare Function FindExecutable Lib "shell32" _
   Alias "FindExecutableA" _
  (ByVal lpFile As String, _
   ByVal lpDirectory As String, _
   ByVal sResult As String) As Long

Public Declare Function GetTempPath Lib "kernel32" _
   Alias "GetTempPathA" _
  (ByVal nSize As Long, _
   ByVal lpBuffer As String) As Long
'--end block--'
'*************************************************************************************************
Private Declare Function PathStripPath Lib "shlwapi" _
   Alias "PathStripPathA" _
  (ByVal pPath As String) As Long
       
Public Enum BYTEVALUES
    KiloByte = 1024
    MegaByte = 1048576
    GigaByte = 107374182
End Enum
'************************************************************************************************

'************************************************************************************************
Function GiveByteValues(Bytes As Double) As String


    If Bytes < BYTEVALUES.KiloByte Then
        GiveByteValues = Bytes & " Bytes"
    ElseIf Bytes >= BYTEVALUES.GigaByte Then
        GiveByteValues = CutDecimal(Bytes / BYTEVALUES.GigaByte, 2) & " GB" '" Gigabytes"
    ElseIf Bytes >= BYTEVALUES.MegaByte Then
        GiveByteValues = CutDecimal(Bytes / BYTEVALUES.MegaByte, 2) & " MB" '" Megabytes"
    ElseIf Bytes >= BYTEVALUES.KiloByte Then
        GiveByteValues = CutDecimal(Bytes / BYTEVALUES.KiloByte, 2) & " KB" '" Kilobytes"
    End If
End Function

Public Function CutDecimal(Number As String, ByPlace As Byte) As String
    Dim Dec As Byte
    Dec = InStr(1, Number, ".", vbBinaryCompare) ' find the Decimal


    If Dec = 0 Then
        CutDecimal = Number 'if there is no decimal Then dont do anything
        Exit Function
    End If
    CutDecimal = Mid(Number, 1, Dec + ByPlace) 'How many places you want after the decimal point
End Function
'****************************************Net Detect**********************************************

    'Text1 = IsNetConnectViaLAN()
    'Text2 = IsNetConnectViaModem()
    'Text3 = IsNetConnectViaProxy()
    'Text4 = IsNetConnectOnline()
    'Text5 = IsNetRASInstalled()
    'Text6 = GetNetConnectString()





Public Function IsNetConnectViaLAN() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwFlags, 0&)
    'return True if the flags indicate a LAN
    '     connection
    IsNetConnectViaLAN = dwFlags And INTERNET_CONNECTION_LAN
End Function


Public Function IsNetConnectViaModem() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwFlags, 0&)
    'return True if the flags indicate a mod
    '     em connection
    IsNetConnectViaModem = dwFlags And INTERNET_CONNECTION_MODEM
End Function


Public Function IsNetConnectViaProxy() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwFlags, 0&)
    'return True if the flags indicate a pro
    '     xy connection
    IsNetConnectViaProxy = dwFlags And INTERNET_CONNECTION_PROXY
End Function


Public Function IsNetConnectOnline() As Boolean
    'no flags needed here - the API returns
    '     True
    'if there is a connection of any type
    IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
End Function


Public Function IsNetRASInstalled() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwFlags, 0&)
    'return True if the falgs include RAS in
    '     stalled
    IsNetRASInstalled = dwFlags And INTERNET_RAS_INSTALLED
End Function

'************************************************************************************************

Public Function HiWord(dw As Long) As Long
  
   If dw And &H80000000 Then
         HiWord = (dw \ 65535) - 1
   Else: HiWord = dw \ 65535
   End If
    
End Function
  

Public Function LoWord(dw As Long) As Long
  
   If dw And &H8000& Then
         LoWord = &H8000& Or (dw And &H7FFF&)
   Else: LoWord = dw And &HFFFF&
   End If
    
End Function


Public Function GetFileDescription(sSourceFile As String) As String

   Dim FI As VS_FIXEDFILEINFO
   Dim sBuffer() As Byte
   Dim nBufferSize As Long
   Dim lpBuffer As Long
   Dim nVerSize As Long
   Dim nUnused As Long
   Dim tmpVer As String
   Dim sBlock As String
   
   If sSourceFile > "" Then

     'set file that has the encryption level
     'info and call to get required size
      nBufferSize = GetFileVersionInfoSize(sSourceFile, nUnused)
      
      ReDim sBuffer(nBufferSize)
      
      If nBufferSize > 0 Then
      
        'get the version info
         Call GetFileVersionInfo(sSourceFile, 0&, nBufferSize, sBuffer(0))
         Call VerQueryValue(sBuffer(0), "\", lpBuffer, nVerSize)
         Call CopyMemory(FI, ByVal lpBuffer, Len(FI))
   
         If VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lpBuffer, nVerSize) Then
            
            If nVerSize Then
               tmpVer = GetPointerToString(lpBuffer, nVerSize)
               tmpVer = Right("0" & Hex(Asc(Mid(tmpVer, 2, 1))), 2) & _
                        Right("0" & Hex(Asc(Mid(tmpVer, 1, 1))), 2) & _
                        Right("0" & Hex(Asc(Mid(tmpVer, 4, 1))), 2) & _
                        Right("0" & Hex(Asc(Mid(tmpVer, 3, 1))), 2)
               sBlock = "\StringFileInfo\" & tmpVer & "\FileDescription"
               
              'Get predefined version resources
               If VerQueryValue(sBuffer(0), sBlock, lpBuffer, nVerSize) Then
               
                  If nVerSize Then
                  
                    'get the file description
                     GetFileDescription = GetStrFromPtrA(lpBuffer)

                  End If  'If nVerSize
               End If  'If VerQueryValue
            End If  'If nVerSize
         End If  'If VerQueryValue
      End If  'If nBufferSize
   End If  'If sSourcefile

End Function


Private Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function


Private Function GetPointerToString(lpString As Long, nbytes As Long) As String

   Dim Buffer As String
   
   If nbytes Then
      Buffer = Space$(nbytes)
      CopyMemory ByVal Buffer, ByVal lpString, nbytes
      GetPointerToString = Buffer
   End If
   
End Function


Public Function GetFileVersion(sDriverFile As String) As String
   
   Dim FI As VS_FIXEDFILEINFO
   Dim sBuffer() As Byte
   Dim nBufferSize As Long
   Dim lpBuffer As Long
   Dim nVerSize As Long
   Dim nUnused As Long
   Dim tmpVer As String
   
  'GetFileVersionInfoSize determines whether the operating
  'system can obtain version information about a specified
  'file. If version information is available, it returns
  'the size in bytes of that information. As with other
  'file installation functions, GetFileVersionInfoSize
  'works only with Win32 file images.
  '
  'A empty variable must be passed as the second
  'parameter, which the call returns 0 in.
   nBufferSize = GetFileVersionInfoSize(sDriverFile, nUnused)
   
   If nBufferSize > 0 Then
   
     'create a buffer to receive file-version
     '(FI) information.
      ReDim sBuffer(nBufferSize)
      Call GetFileVersionInfo(sDriverFile, 0&, nBufferSize, sBuffer(0))
      
     'VerQueryValue function returns selected version info
     'from the specified version-information resource. Grab
     'the file info and copy it into the  VS_FIXEDFILEINFO structure.
      Call VerQueryValue(sBuffer(0), "\", lpBuffer, nVerSize)
      Call CopyMemory(FI, ByVal lpBuffer, Len(FI))
     
     'extract the file version from the FI structure
      tmpVer = Format$(HiWord(FI.dwFileVersionMS)) & "." & _
               Format$(LoWord(FI.dwFileVersionMS), "00") & "."
         
      If FI.dwFileVersionLS > 0 Then
         tmpVer = tmpVer & Format$(HiWord(FI.dwFileVersionLS), "00") & "." & _
                           Format$(LoWord(FI.dwFileVersionLS), "00")
      Else
         tmpVer = tmpVer & Format$(FI.dwFileVersionLS, "0000")
      End If
      
      End If
   
   GetFileVersion = tmpVer
   
End Function
'--end block--'

'************************************************************************************************

Public Function GetFilePath(ByVal sFilename As String, Optional ByVal bAddBackslash As Boolean) As String
    'Returns Path Without FileTitle
    Dim lPos As Long
    lPos = InStrRev(sFilename, "\")


    If lPos > 0 Then
        GetFilePath = Left$(sFilename, lPos - 1) _
        & IIf(bAddBackslash, "\", "")
    Else
        GetFilePath = ""
    End If
    
End Function

Public Function GetAppPath() As String
If Right(App.Path, 1) <> "\" Then GetAppPath = App.Path & "\" Else GetAppPath = App.Path
End Function
'***************************************Build File from resource*****************************************
Public Function BuildFileFromResource(destFILE As String, resID As Long, Optional resTITLE As String = "CUSTOM") As String
    On Error GoTo ErrorBuildFileFromResource
    Dim resBYTE() As Byte
    resBYTE = LoadResData(resID, resTITLE)
    Open destFILE For Binary Access Write As #1
    Put #1, , resBYTE
    Close #1
    BuildFileFromResource = destFILE
    Exit Function
ErrorBuildFileFromResource:
    BuildFileFromResource = ""
    MsgBox Err & ":Error in BuildFileFromResource.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function
Public Function GetFileCompany(sSourceFile As String) As String

   Dim FI As VS_FIXEDFILEINFO
   Dim sBuffer() As Byte
   Dim nBufferSize As Long
   Dim lpBuffer As Long
   Dim nVerSize As Long
   Dim nUnused As Long
   Dim tmpVer As String
   Dim sBlock As String
   
   If sSourceFile > "" Then

     'set file that has the encryption level
     'info and call to get required size
      nBufferSize = GetFileVersionInfoSize(sSourceFile, nUnused)
      
      ReDim sBuffer(nBufferSize)
      
      If nBufferSize > 0 Then
      
        'get the version info
         Call GetFileVersionInfo(sSourceFile, 0&, nBufferSize, sBuffer(0))
         Call VerQueryValue(sBuffer(0), "\", lpBuffer, nVerSize)
         Call CopyMemory(FI, ByVal lpBuffer, Len(FI))
   
         If VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lpBuffer, nVerSize) Then
            
            If nVerSize Then
               tmpVer = GetPointerToString(lpBuffer, nVerSize)
               tmpVer = Right("0" & Hex(Asc(Mid(tmpVer, 2, 1))), 2) & _
                        Right("0" & Hex(Asc(Mid(tmpVer, 1, 1))), 2) & _
                        Right("0" & Hex(Asc(Mid(tmpVer, 4, 1))), 2) & _
                        Right("0" & Hex(Asc(Mid(tmpVer, 3, 1))), 2)
               sBlock = "\StringFileInfo\" & tmpVer & "\CompanyName"
               
              'Get predefined version resources
               If VerQueryValue(sBuffer(0), sBlock, lpBuffer, nVerSize) Then
               
                  If nVerSize Then
                  
                    'get the file description
                     GetFileCompany = GetStrFromPtrA(lpBuffer)

                  End If  'If nVerSize
               End If  'If VerQueryValue
            End If  'If nVerSize
         End If  'If VerQueryValue
      End If  'If nBufferSize
   End If  'If sSourcefile

End Function

