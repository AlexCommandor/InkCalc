VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MTInk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Sub NewFormThread(Optr As Long, MainFormHandle As Long)
        '<EhHeader>
        On Error GoTo NewFormThread_Err
        '</EhHeader>

        'The interesting part is here
        'We store the WindowHandle of the MainWindow and a pointer
        'to the Main Window Form by using the SetProp function against
        'The Hwnd of the New Form as a virtual property
        Dim Frm As New FrmSysTrayMT
100     Load Frm
102     SetProp Frm.hWnd, "MAINHWND", MainFormHandle
104     SetProp Frm.hWnd, "OBJPTR", Optr
106     Set Frm = Nothing

        '<EhFooter>
        Exit Sub

NewFormThread_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.MTInk.NewFormThread " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub



