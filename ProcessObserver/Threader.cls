VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Threader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum ThreadPriority
    THREAD_PRIORITY_IDLE = -15
    THREAD_PRIORITY_LOWEST = -2
    THREAD_PRIORITY_BELOW_NORMAL = -1
    THREAD_PRIORITY_NORMAL = 0
    THREAD_PRIORITY_ABOVE_NORMAL = 1
    THREAD_PRIORITY_HIGHEST = 2
    THREAD_PRIORITY_TIME_CRITICAL = 15
End Enum

Public Enum ThreadCallType
[UsingAPC]
[UsingCallback]
[UsingPostMessage]
[DirectCall]
End Enum

Public Enum TransferArguments
[UsingCopy]
[UsingMove]
End Enum

Public Enum ThreadingType
[ObjectEnabled]
[NonObject]
End Enum


Private InternalData As String * 48
Sub InitCaller(ByVal HInstance As Long)
End Sub
Sub CloseCaller(ByVal HInstance As Long)
End Sub
Function CreateNewThread(ByRef ThreadId As Long, ByVal StackSize As Long, ByVal Priority As ThreadPriority, ByVal Reason As Long, ByVal Message As Long, ByRef Args As Variant, ByVal ThreadType As ThreadingType) As Long
End Function
Sub AddThread(CallerInterface As InThreadCall)
End Sub
Sub RemoveThread(ByVal ThreadId As Long)
End Sub
Public Function CheckThreadExist(ByVal ThreadId As Long) As Boolean
'FALSE=Thread doesn't exist,TRUE=exist..
End Function
Public Function CallThread(ByVal ThreadId As Long, ByVal Reason As Long, ByVal Message As Long, ByRef Args As Variant, ByVal CallingTransfer As TransferArguments, ByVal CallingType As ThreadCallType, ByRef IsValidCall As Boolean) As Long
'Return Value from [CallBack or Direct] Type of calling from called thread!
'IsValidCall notify if call was successfull!
End Function
Public Function EnterSynchronization(ByVal UsingTry As Boolean) As Boolean
'Return False if sync object was in use [Using Try only],True if not!
End Function
Public Sub LeaveSynchronization()
End Sub
Public Function WaitForLocalObject(ByVal ThreadId As Long, ByVal WaitTime As Long) As Long
'Returns:
'Object Evented=0, Thread No longer exist =-1, The time-out interval elapsed=258
End Function
Public Function SignalLocalObject(ByVal ThreadId As Long) As Long
'Return Success=1,Fail=0
End Function
Public Sub TranslateArguments(ByVal CallArgs As Long, ByRef Reason As Long, ByRef Message As Long, ByRef Arguments As Variant)
End Sub
Public Sub GetThreadParams(ByRef Reason As Long, ByRef Message As Long, ByRef Arguments As Variant)
End Sub
Public Function AttachThreadCallArguments(ByVal ThreadId As Long, ByVal Index As Long, ByVal Reason As Long, ByVal Message As Long, ByRef Arguments As Variant, ByVal AttachTransfer As TransferArguments, ByRef PreviousCallArguments As Long, ByVal AttachIfExist As Boolean) As Long
'Returns:
'Succes=0,Invalid ThreadId=-1
End Function
Public Function DetachThreadCallArguments(ByVal ThreadId As Long, ByVal Index As Long, ByRef Reason As Long, ByRef Message As Long, ByRef Arguments As Variant, ByVal RemoveFromIndex As Boolean) As Long
'Returns:
'Succes=0,Invalid ThreadId or Index dosn't exist= -1
End Function
Public Function RemoveThreadCallArguments(ByVal ThreadId As Long, ByVal Index As Long) As Long
'Returns:
'Succes=0,Invalid ThreadId or Index dosn't exist= -1
End Function
Public Sub FastDoEvents(ByVal hWnd As Long)
End Sub
Public Sub EnterMessagePump()
End Sub
Public Sub ExitMessagePump()
End Sub
Public Function UpdateVirtualMachine() As Boolean
End Function
Public Sub AboutBox()
End Sub

