VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InThreadCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***Do not change anything and keep it as TEMPLATE!!!!!!!!!!!

'Public Function ThreadCall(ByVal CallArgs As Long) As Long
'Only Callback And Direct type returns value to Caller!
'APC and Post does not return value!
'End Function
'Public Function EventCall(ByVal ThreadNotify As Long) As Long
'If ThreadNotify < 0 Then
    ''worker thread termination notify
    'ThreadId = ThreadNotify And &H7FFFFFFF
'Else
    ''worker thread creation notify
    'ThreadId = ThreadNotify
'End If
'End Function

Public Function ThreadCall(ByVal CallArgs As Long) As Long
End Function
Public Function EventCall(ByVal ThreadNotify As Long) As Long
End Function
