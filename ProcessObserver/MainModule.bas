Attribute VB_Name = "MainModule"
Declare Sub GatherObject Lib "vb6multithread.dll" (OBJ As Object)

Sub Main()
Static IsInit As Boolean

If Not IsInit Then
    Dim IsUpdated As Boolean
    'Main Thread
    IsInit = True
    GatherObject MULTITHREADER 'First Step required!
    IsUpdated = MULTITHREADER.UpdateVirtualMachine
    MULTITHREADER.InitCaller App.hInstance 'Second Step required!
    MULTITHREADER.AddThread FrmMain

    
Else
    
    'New Threads
    Dim Args As Variant
    Dim Reason As Long
    Dim Message As Long

    MULTITHREADER.GetThreadParams Reason, Message, Args 'Get Thread Parameters! {U can get it only once!}
    MULTITHREADER.AddThread WorkerC
    
    
End If


End Sub
