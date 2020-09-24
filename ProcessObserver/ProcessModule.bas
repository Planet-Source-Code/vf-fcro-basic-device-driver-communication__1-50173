Attribute VB_Name = "ProcessObserver"
Public hService As Long
Public hComm As Long
Public CEvent As Long

Public Type PROCESSCREATION
ParentProcess As Long
processId As Long
bCreate As Byte
End Type


Public Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long
    PebBaseAddress As Long
    AffinityMask As Long
    BasePriority As Long
    UniqueProcessId As Long
    InheritedFromUniqueProcessId As Long
End Type

Declare Function OpenEvent Lib "kernel32" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function NtQueryInformationProcess Lib "ntdll" (ByVal hProcess As Long, ByVal ProcessInfoClass As Long, ByRef ProcInfoOut As Any, ByVal BufferLength As Long, ByRef ProcLength As Long) As Long
Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const SYNCHRONIZE = &H100000
Public SCMANAGER As New ServiceManager 'Device Driver Manager!
Public PROCESSES As New Collection

Public WorkerId As Long

Public Sub AddProcess(ByVal processId As Long, ByVal CText As String)
On Error GoTo Dalje
Dim TPid As String
TPid = "X" & processId
PROCESSES.Add CText, TPid
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Sub RemoveProcess(ByVal processId As Long)
On Error GoTo Dalje
Dim TPid As String
TPid = "X" & processId
PROCESSES.Remove TPid
Exit Sub
Dalje:
On Error GoTo 0
End Sub
Public Function GetProcess(ByVal processId As Long) As String
On Error GoTo Dalje
Dim TPid As String
TPid = "X" & processId
GetProcess = PROCESSES.Item(TPid)
Exit Function
Dalje:
On Error GoTo 0
End Function

Public Function GetPeb(ByVal ProcessH As Long) As Long
Dim PBI As PROCESS_BASIC_INFORMATION
Dim Bfrlen As Long
Call NtQueryInformationProcess(ProcessH, 0, PBI, Len(PBI), Bfrlen)
GetPeb = PBI.PebBaseAddress
End Function
Public Function GetStartParams(ByVal ProcessH As Long, ByVal PEB As Long) As String
Dim SAdr As Long
Dim StartParamP As Long
SAdr = LongFromPTR(ProcessH, PEB + 16)
StartParamP = LongFromPTR(ProcessH, SAdr + 68)
GetStartParams = StringFromPTR(ProcessH, StartParamP, , 1)
End Function
Private Function LongFromPTR(ByVal ProcessHandle As Long, ByVal Address As Long) As Long
ReadProcessMemory ProcessHandle, ByVal Address, LongFromPTR, 4, ByVal 0&
End Function
Private Function StringFromPTR(ByVal ProcessHandle As Long, ByVal Address As Long, Optional ByVal Mlen As Long = 255, Optional ByVal Unicode As Long) As String
If Mlen <= 0 Then Exit Function
Dim StringsD() As Byte
Dim Iret As Long
Dim LLen As Long
Do
ReDim StringsD(Mlen - 1)
Iret = ReadProcessMemory(ProcessHandle, ByVal Address, StringsD(0), Mlen, ByVal 0&)
Mlen = Mlen / 2
If Mlen = 0 Then Exit Function
Loop While Iret = 0

If Unicode = 1 Then
LLen = lstrlenW(StringsD(0))
Else
LLen = lstrlen(StringsD(0))
End If

StringFromPTR = Space(LLen)

If Unicode = 1 Then
CopyMemory ByVal StrPtr(StringFromPTR), StringsD(0), LLen * 2
Else
CopyMemory ByVal StringFromPTR, StringsD(0), LLen
End If

End Function
Public Sub Notify()
Dim PS As PROCESSCREATION
Dim PEB As Long
Dim CH As Long
Dim S As String
Dim IsValid As Boolean


Again:
PS.bCreate = 0

WaitNotify PS


If PS.bCreate <> 0 Then
MULTITHREADER.CallThread 0, 1, PS.processId, Empty, UsingCopy, UsingAPC, IsValid
Else
MULTITHREADER.CallThread 0, 0, PS.processId, Empty, UsingCopy, UsingAPC, IsValid
End If


GoTo Again
End Sub
Public Sub InitDeviceDriver()
Dim CSTATUS As SERVICE_STATUS
SCMANAGER.InitializeManager
hService = SCMANAGER.InstallService("ProcObsrv", SERVICE_KERNEL_DRIVER, App.Path & "\ProcObsrv.sys")
Call SCMANAGER.StartService(hService, 0, 0)
hComm = SCMANAGER.OpenCommunicator("\\.\ProcObsrv", 0)
CEvent = OpenEvent(SYNCHRONIZE, 0, "ProcObsrvProcessEvent")
End Sub
Public Sub FreeDeviceDriver()
Set PROCESSES = Nothing
Dim CSTATUS As SERVICE_STATUS
CloseHandle CEvent
SCMANAGER.CloseCommunicator hComm
SCMANAGER.Control hService, SERVICE_CONTROL_STOP, CSTATUS
SCMANAGER.UninstallService hService
SCMANAGER.CloseManager
End Sub
Public Sub StartWatch()
Dim Stop_Start As Byte
Stop_Start = 1
Call SCMANAGER.CallDevice(hComm, &H22E000, VarPtr(Stop_Start), 1, 0, 0, 0&, 0)
End Sub
Public Sub StopWatch()
Dim Stop_Start As Byte
Stop_Start = 0
Call SCMANAGER.CallDevice(hComm, &H22E000, VarPtr(Stop_Start), 1, 0, 0, 0&, 0)
End Sub
Public Sub WaitNotify(PCREATE As PROCESSCREATION)
WaitForSingleObject CEvent, -1
Call SCMANAGER.CallDevice(hComm, &H22E004, 0, 0, VarPtr(PCREATE), 12, 0, 0)
End Sub
