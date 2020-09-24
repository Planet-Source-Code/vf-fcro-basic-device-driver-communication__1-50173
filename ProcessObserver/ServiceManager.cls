VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServiceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private SCHandle As Long
'Return TRUE if succeeds
Public Function InitializeManager() As Boolean
SCHandle = OpenSCManager(vbNullString, vbNullString, &HF003F)
InitializeManager = CBool(SCHandle)
End Function

'Return TRUE if succeeds
Public Function CloseManager() As Boolean
If SCHandle <> 0 Then CloseManager = CBool(CloseServiceHandle(SCHandle))
End Function

'Return Service Handle
Public Function InstallService(ByVal AliasName As String, ByVal ServiceType As Service_Type, ByVal PathToDriver As String) As Long
InstallService = CreateService(SCHandle, AliasName, AliasName, &HF01FF, ServiceType, SERVICE_DEMAND_START, SERVICE_ERROR_NORMAL, PathToDriver, vbNullString, ByVal 0&, vbNullString, vbNullString, vbNullString)
If InstallService = 0 Then
InstallService = OpenService(SCHandle, AliasName, &HF01FF)
End If
End Function
Friend Function QueryStatus(ByVal hService As Long, SStatus As SERVICE_STATUS) As Boolean
QueryStatus = CBool(QueryServiceStatus(hService, SStatus))
End Function

'Return TRUE if succeeds
Friend Function Control(ByVal hService As Long, ByVal ControlType As Service_Control, SStatus As SERVICE_STATUS) As Boolean
Control = CBool(ControlService(hService, ControlType, SStatus))
End Function

'Return TRUE if succeeds
Public Function StartService(ByVal hService As Long, ByVal ArgCount As Long, ByVal Arguments As Long) As Boolean
StartService = CBool(Start_Service(hService, ArgCount, Arguments))
End Function

'Return TRUE if succeeds
Public Function UninstallService(ByVal hService As Long) As Boolean
UninstallService = CBool(DeleteService(hService))
If UninstallService Then CloseServiceHandle hService
End Function

'Return Service Communicator
'Flags=dwFlagsAndAttributes for File
Public Function OpenCommunicator(ByVal SymbolicLink As String, ByVal Flags As Long) As Long
OpenCommunicator = CreateFile(SymbolicLink, &HC0000000, 0, ByVal 0&, 3, Flags, 0)
End Function
'Return TRUE if succeeds
Public Function CloseCommunicator(ByVal CHandle As Long) As Long
CloseCommunicator = CBool(CloseHandle(CHandle))
End Function

'Communication with Device Driver!
Public Function CallDevice(ByVal CHandle As Long, ByVal CtrlCode As Long, ByVal InputBuffer As Long, ByVal InputBufferLength As Long, ByVal OutputBuffer As Long, ByVal OutputBufferLength As Long, ByRef ReturnedBytes As Long, ByVal AsyncOp As Long) As Long
CallDevice = DeviceIoControl(CHandle, CtrlCode, ByVal InputBuffer, InputBufferLength, ByVal OutputBuffer, OutputBufferLength, ReturnedBytes, ByVal AsyncOp)
End Function

