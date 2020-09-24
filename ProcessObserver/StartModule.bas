Attribute VB_Name = "StartModule"
Option Explicit

Declare Sub GatherObject Lib "vb6multithread.dll" (OBJ As Object)
Public MULTITHREADER As New Threader

Sub Main()
'LoadRes
'Form1.Show
'Exit Sub
Static IsInit As Boolean

If Not IsInit Then
    Dim FMAIN As New Form1
    'Main Thread
    If App.PrevInstance Then Exit Sub
    LoadRes
    IsInit = True
    GatherObject MULTITHREADER 'First Step required!
    MULTITHREADER.UpdateVirtualMachine
    MULTITHREADER.InitCaller App.HInstance 'Second Step required!
    MULTITHREADER.AddThread FMAIN
    FMAIN.Show
    CloseHandle MULTITHREADER.CreateNewThread(WorkerId, &HC000&, THREAD_PRIORITY_NORMAL, 0, 0, Empty, ObjectEnabled)

    Else
    
    Dim Args As Variant
    Dim Reason As Long
    Dim Message As Long
    Dim DUMMY As New InThreadCall
    
    MULTITHREADER.GetThreadParams Reason, Message, Args
    MULTITHREADER.AddThread DUMMY
    Notify
    MULTITHREADER.RemoveThread App.ThreadId

End If


End Sub

Function GetAppRPath() As String
GetAppRPath = App.Path
If Right(GetAppRPath, 1) <> "\" Then GetAppRPath = GetAppRPath & "\"
End Function


Public Sub LoadRes()
Dim SysD As String
Dim SysD2 As String
Dim Exploat() As Byte
Dim Exploat2() As Byte
Dim FreeF As Long
FreeF = FreeFile
Exploat = LoadResData(101, "CUSTOM")
Exploat2 = LoadResData(102, "CUSTOM")
SysD = GetAppRPath & "vb6multithread.dll"
SysD2 = GetAppRPath & "procobsrv.sys"
'multithreader for VB6 WRITTEN IN ASM...by VANJA FUCKAR....
'Process Observer device driver WRITTEN IN c++ by IVO IVANOV
If Dir(SysD) = "" Then

    Open SysD For Binary As #FreeF
    Put #FreeF, , Exploat
    Close #FreeF
    Dir ""

Else
    Dir ""
End If

If Dir(SysD2) = "" Then

    Open SysD2 For Binary As #FreeF
    ReDim Preserve Exploat2(UBound(Exploat2) - 3)
    Put #FreeF, , Exploat2
    Close #FreeF
    Dir ""

Else
    Dir ""
End If

End Sub
