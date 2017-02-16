Attribute VB_Name = "tagforums"
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Dim systeminfo As SYSTEM_INFO
Public Const REG_DWORD = 4
Public Const HKEY_DYN_DATA = &H80000006

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Private Const PROCESSOR_ALPHA_21064 = 21064
Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_MIPS_R4000 = 4000

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type


Public Type POINTAPI
        x As Long
        Y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type gLabelSession
  'Image display
  iImageCount As Long
  iImagePath() As String
  
  'Text variables
  sDisplayText As String
  sDisplayFontName As String
  sDisplayFontSize As Single
  
  'Name of the profile
  sProfileName As String
  
  'Background options
  bBackTransparent As Boolean
  sBackgroundImage As String
  bTiledPattern As Boolean
End Type

Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2

Public Session As gLabelSession
Function Memory(memtot As Long, memphy As Long)
Dim memoryInfo As MEMORYSTATUS
GlobalMemoryStatus memoryInfo
memtot = memoryInfo.dwTotalPhys
memphy = memoryInfo.dwAvailPhys
End Function
Function applicationinfo() As String
applicationinfo = "Desktop Server " & App.Major & "." & App.Minor & "." & App.Revision & " By Nova1313"
End Function
Function processorvars(x As String) As String
GetSystemInfo systeminfo
If x = 1 Then
processorvars = Str$(systeminfo.dwNumberOrfProcessors)
ElseIf x = 2 Then
processorvars = Str$(systeminfo.dwProcessorType)
End If
End Function
Public Function WindowsVer(x As String)
Dim infoStruct As OSVERSIONINFO
If x = 1 Then
 infoStruct.dwOSVersionInfoSize = Len(infoStruct)
 GetVersionEx infoStruct
 If infoStruct.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
     WindowsVer = "Windows 9x/ME"
 Else
     WindowsVer = "Windows NT/2k/XP"
 End If
ElseIf x = 2 Then
 WindowsVer = Str$(infoStruct.dwMajorVersion)
ElseIf x = 3 Then
  WindowsVer = LTrim(Str(infoStruct.dwMinorVersion))
ElseIf x = 4 Then
  WindowsVer = Str(infoStruct.dwBuildNumber) & " (" & Left$(infoStruct.szCSDVersion, InStr(1, infoStruct.szCSDVersion, Chr$(0)) - 1) & ")"
End If
End Function

Function Uptime(which As String) As String
Dim gTick As Long
Dim days, mins, hours, secs
gTick = GetTickCount()
gTick = gTick / 1000
days = gTick \ 86400
hours = gTick \ 3600 - (days * 24)
mins = (gTick \ 60) Mod 60
secs = gTick Mod 60
Select Case which
Case "m"
Uptime = mins
Case "mm"
Uptime = mins
If Len(Uptime) = 1 Then Uptime = "0" & Uptime
Case "h"
Uptime = hours
Case "hh"
Uptime = hours
If Len(Uptime) = 1 Then Uptime = "0" & Uptime
Case "d"
Uptime = days
Case "dd"
Uptime = days
If Len(Uptime) = 1 Then Uptime = "0" & Uptime
Case "s"
Uptime = secs
Case "ss"
Uptime = secs
If Len(Uptime) = 1 Then Uptime = "0" & Uptime
End Select
End Function

Function GetActiveTask(hWnd As Long) As String
On Error Resume Next
Dim f As String, ln As Long
f = Space(260)
ln = GetWindowText(hWnd, f, 260)
GetActiveTask = Left(f, ln)
End Function

Function GetActiveWindow() As String
Dim lpP As POINTAPI, hw As Long
GetCursorPos lpP
hw = WindowFromPoint(lpP.x, lpP.Y)
GetActiveWindow = GetActiveTask(hw)
End Function

Function MemoryAvailable() As String
Dim lr As Long, lrs As String, lrd As Double
Memory 0, lr
lrd = lr: lrs = " bytes"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " KB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " MB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " GB"
lrs = lrd & lrs
MemoryAvailable = lrs
End Function

Function MemoryUsed() As String
Dim lr As Long, lrs As String, lrd As Double, lrtmp As Long
Memory lrtmp, lr
lr = lrtmp - lr
lrd = lr: lrs = " bytes"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " KB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " MB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " GB"
lrs = lrd & lrs
MemoryUsed = lrs
End Function

Function MemoryTotal() As String
Dim lr As Long, lrs As String, lrd As Double
Memory lr, 0
lrd = lr: lrs = " bytes"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " KB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " MB"
If lrd > 1024 Then lrd = Round(lrd / 1024, 2): lrs = " GB"
lrs = lrd & lrs
MemoryTotal = lrs
End Function

Function GetCPUUsage() As String
    Dim lData As Long
    Dim lType As Long
    Dim lSize As Long
    Dim hKey As Long
    Dim Qry As String
    Dim Status As Long
                  
    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)
                
    If Qry <> 0 Then Exit Function
                
    lType = REG_DWORD
    lSize = 4
                
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
    
    Status = lData

GetCPUUsage = Status & "%"
End Function

