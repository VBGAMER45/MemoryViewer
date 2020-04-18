Attribute VB_Name = "modGlobals"


Public Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type


Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long

Public Const TH32CS_SNAPPROCESS As Long = 2&

Type PROCESSENTRY32
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




Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type



Public Type SYSTEM_INFO ' 36 Bytes
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type


Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function VirtualQueryEx& Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long)
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)




Public Const PROCESS_VM_READ = (&H10)
Public Const PROCESS_VM_WRITE = (&H20)
Public Const PROCESS_VM_OPERATION = (&H8)
Public Const PROCESS_QUERY_INFORMATION = (&H400)
Public Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION

Public Const MEM_PRIVATE& = &H20000
Public Const MEM_COMMIT& = &H1000


Global PIDs(1000) As Long
Global BoxXpos As Long
Global BoxYpos As Long
Global BoxAsciiXpos As Long
Global BoxAsciiYpos As Long
Global HexArray(1000) As Byte
Global OffsetArray(1000) As Double
Global DrawGridLines As Boolean


'Store the Version
Global Version As String
Global mBaseAddress As Long
Global mMaxAddress As Long
Global mCurrentAddress As Long
Global myHandle As Long
Global ProcessNumber As Long
Global ProccessName As String


'Color Globals
Global cOffSetColor As Long
Global cBackroundColor As Long
Global cHexColor As Long
Global cGridColor As Long
Global ProgramLoaded As Boolean
Global MyPid As Long
Global si As SYSTEM_INFO
Public Declare Function GetModuleHandleA Lib "kernel32" (ByVal ModName As Any) As Long
Public Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" _
         (ByVal hProcess As Long, ByVal hModule As Long, _
            ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Const MAX_PATHLEN = &H104
Private Type PROCESSENTRY31
      dwSize As Long
      cntUsage  As Long
      th32ProcessID As Long
      th32DefaultHeapID As Long
      th32ModuleID As Long
      cntThreads As Long
      th32ParentProcessID As Long
      pcPriClassBase As Long
      dwFlags As Long
      exeFilename(1 To MAX_PATHLEN) As Byte
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
   (ByVal dwFlags As Long, ByVal dprocess As Long) As Long
Private Declare Function Process32First Lib "kernel32" _
   (ByVal hSnapshot As Long, pProcessEntry As PROCESSENTRY31) As Long

Private Declare Function Process32Next Lib "kernel32" _
   (ByVal hSnapshot As Long, pProcessEntry As PROCESSENTRY31) As Long
Private Declare Function EnumProcessModules Lib "psapi" _
   (ByVal ProcessId As Long, hModule As Long, ByVal cbSize As Long, _
    cbReturned As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi" _
   Alias "GetModuleFileNameExA" (ByVal hProcess As Long, _
   ByVal hModule As Long, ByVal lpFileName As String, _
   ByVal nSize As Long) As Long
Private Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long


Private Const MEM_DECOMMIT  As Long = &H4000

Public Function Hex2ASCII(sText As String) As String
    On Error Resume Next
    Dim sBuff() As String, a As Long
    sBuff() = Split(sText, Space$(1))
    For a = 0 To UBound(sBuff)
        Hex2ASCII = Hex2ASCII & Chr$("&h" & sBuff(a))
        DoEvents
    Next a
End Function

Public Function Hex2Dec(sText As String) As Long
    On Error GoTo err
    Dim H As String
    H = sText
    Dim Tmp$
    Dim lo1 As Integer, lo2 As Integer
    Dim hi1 As Long, hi2 As Long
    Const Hx = "&H"
    Const BigShift = 65536
    Const LilShift = 256, Two = 2
    Tmp = H
    If UCase(Left$(H, 2)) = "&H" Then Tmp = Mid$(H, 3)
    Tmp = Right$("0000000" & Tmp, 8)
        If IsNumeric(Hx & Tmp) Then
            lo1 = CInt(Hx & Right$(Tmp, Two))
            hi1 = CLng(Hx & Mid$(Tmp, 5, Two))
            lo2 = CInt(Hx & Mid$(Tmp, 3, Two))
            hi2 = CLng(Hx & Left$(Tmp, Two))
            Hex2Dec = CCur(hi2 * LilShift + lo2) * BigShift + (hi1 * LilShift) + lo1
        End If
Exit Function
err:
End Function
Public Function BigDecToHex(ByVal DecNum) As String
    ' This function is 100% accurate untill
    '     15,000,000,000,000,000 (1.5E+16)
    Dim NextHexDigit As Double
    Dim HexNum As String
    HexNum = ""


    While DecNum <> 0
    NextHexDigit = DecNum - (Int(DecNum / 16) * 16)


    If NextHexDigit < 10 Then
        HexNum = Chr(Asc(NextHexDigit)) & HexNum
    Else
        HexNum = Chr(Asc("A") + NextHexDigit - 10) & HexNum
    End If
    DecNum = Int(DecNum / 16)
Wend
'HexNum = "1"
If Len(HexNum) = 1 Then HexNum = ("0" + HexNum)

If HexNum = "" Then HexNum = "00"
BigDecToHex = HexNum
End Function



Function ExePathFromProcessId(ByVal ProcessId As Long) As String
   Dim hProcess As Long
   '
   If Not IsWindowsNT Then   ' Win2000 is included
      Dim Process As PROCESSENTRY31, hSnap As Long, f As Long, Exename$
      hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
      If hSnap = -1 Then Exit Function
      Process.dwSize = Len(Process)
      f = Process32First(hSnap, Process)
      Do While f <> 0
         If Process.th32ProcessID = ProcessId Then
            GoSub GetExeName
            ExePathFromProcessId = Exename ' strZtoStr(process.exeFilename)
            Call CloseHandle(hSnap)
            Exit Function
            End If
         Process.dwSize = Len(Process)
         f = Process32Next(hSnap, Process)
      Loop
   Else
      Dim s As String, c As Long, hModule As Long
      Const cMaxPath = 1023
      s = String$(cMaxPath, 0)
      hProcess = ProcessHandleFromProcessId(ProcessId)
      hModule = BaseModuleHandleFromProcessId(ProcessId)
      c = GetModuleFileNameEx(hProcess, hModule, s, cMaxPath)
      If c Then ExePathFromProcessId = Left$(s, c)
      End If
   Exit Function

GetExeName:
   Dim i&, cb As Byte
   Exename = ""
   Do While i < MAX_PATHLEN
      i = i + 1
      cb = Process.exeFilename(i)
      If cb = 0 Then Return
      Exename = Exename & Chr$(cb)
      Loop
   Return
End Function

Private Function IsWindowsNT() As Boolean
   Dim verinfo As OSVERSIONINFO
   verinfo.dwOSVersionInfoSize = Len(verinfo)
   If (GetVersionEx(verinfo)) = 0 Then Exit Function
   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
End Function

Function ProcessHandleFromProcessId(ProcessId As Long)
   ProcessHandleFromProcessId = _
      OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessId)
End Function

Function BaseModuleHandleFromProcessId(ProcessId As Long)
   Dim hModule As Long, hProcess As Long, bReturned As Long
   hProcess = ProcessHandleFromProcessId(ProcessId)
   Call EnumProcessModules(hProcess, hModule, 4&, bReturned)
   BaseModuleHandleFromProcessId = hModule
End Function

Public Function ReadLongFromMemory(MemoryAddress As Long, SelectedProcessHandle As Long) As Long
Dim Buffer As Long



Call ReadProcessMemory(SelectedProcessHandle, MemoryAddress, Buffer, 4, 0&)
ReadLongFromMemory = Buffer

End Function


Public Function ReadLongArrayFromMemory(MemoryAddress As Long, NumberOfLongs As Long, SelectedProcessHandle As Long) As Long()

Dim Buffer() As Long
Dim ret As Long
ReDim Buffer(NumberOfLongs)
Dim Old As Long

    ret = VirtualProtectEx(SelectedProcessHandle, MemoryAddress, NumberOfLongs * 4, &H40, Old)

    ret = ReadProcessMemory(SelectedProcessHandle, MemoryAddress, Buffer(0), NumberOfLongs * 4, 0&)
    
  '  If ret = 0 Then DoError "Failed to read Longs from memory " & GetLastError: TRACE "ReadLongArrayFromMemory Suffers from error " & GetLastError()
    ReadLongArrayFromMemory = Buffer

End Function
Public Function ReadIntFromMemory(MemoryAddress As Long, SelectedProcessHandle As Long) As Integer
Dim Buffer As String * 2


Call ReadProcessMemory(SelectedProcessHandle, MemoryAddress, Buffer, 2, 0&)
ReadIntFromMemory = Buffer

End Function
