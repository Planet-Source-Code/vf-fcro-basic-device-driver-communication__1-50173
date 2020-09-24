VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements InThreadCall












Private Sub Form_Load()
Dim S As String
S = "Process Notificator! by Vanja Fuckar!" & vbCrLf & "Based on Device Driver by Ivo Ivanov!" & vbCrLf & "Used multithreader for VB6 by Vanja Fuckar!" & vbCrLf & "Used Service Manager by Vanja Fuckar!"
MsgBox S, vbInformation, "Info!"
Caption = "Process Notificator!"
InitDeviceDriver
StartWatch
End Sub

Private Sub Form_Resize()
List1.Top = 0
List1.Left = 0
List1.Width = ScaleWidth
List1.Height = ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
StopWatch
MULTITHREADER.RemoveThread WorkerId
MULTITHREADER.CloseCaller App.HInstance
MULTITHREADER.RemoveThread 0
FreeDeviceDriver
End Sub

Private Function InThreadCall_EventCall(ByVal ThreadNotify As Long) As Long

End Function

Private Function InThreadCall_ThreadCall(ByVal CallArgs As Long) As Long
Dim Reason As Long
Dim Message As Long
Dim ARGUMENT As Variant
Dim PString As String
MULTITHREADER.TranslateArguments CallArgs, Reason, Message, ARGUMENT
PString = CStr(ARGUMENT)

If List1.ListCount > 100 Then List1.RemoveItem (List1.ListCount - 1)

Select Case Reason
Case 0
    PString = GetProcess(Message)
    If Len(PString) <> 0 Then PString = ", " & PString
    List1.AddItem "Process Terminated >PID:" & Message & PString, 0
    RemoveProcess Message
Case 1
    CH = OpenProcess(PROCESS_ALL_ACCESS, 0, Message)
    WaitForInputIdle CH, &H1000&
    PEB = GetPeb(CH)
    PString = GetStartParams(CH, PEB)
    CloseHandle CH
    AddProcess Message, PString
    If Len(PString) <> 0 Then PString = ", " & PString
    List1.AddItem "Process Created >PID:" & Message & PString, 0

End Select



End Function

