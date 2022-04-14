VERSION 5.00
Begin VB.Form FrmSysTray 
   BorderStyle     =   0  'None
   ClientHeight    =   1110
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3165
   Icon            =   "FrmSys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox pictProcess 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1320
      Picture         =   "FrmSys.frx":0742
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   240
      Width           =   540
   End
   Begin VB.PictureBox Flash2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   720
      Picture         =   "FrmSys.frx":0A4C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox Flash1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "FrmSys.frx":0D56
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Timer TmrFlash 
      Interval        =   1000
      Left            =   2280
      Top             =   120
   End
   Begin VB.Menu mPopupMenu 
      Caption         =   "&PopupMenu"
      Begin VB.Menu mAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mMaximize 
         Caption         =   "Ma&ximize"
         Enabled         =   0   'False
      End
      Begin VB.Menu mRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mMinimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public WithEvents FSys As Form
Attribute FSys.VB_VarHelpID = -1
Public Event Click(ClickWhat As String)
Public Event TIcon(F As Form)

Private nid As NOTIFYICONDATA
Private LastWindowState As Integer

Public Property Let Tooltip(Value As String)
        '<EhHeader>
        On Error GoTo Tooltip_Err
        '</EhHeader>

100     nid.szTip = Value & vbNullChar

        '<EhFooter>
        Exit Property

Tooltip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.Tooltip " & _
               "at line " & Erl
        End
        '</EhFooter>
End Property

Public Property Get Tooltip() As String
        '<EhHeader>
        On Error GoTo Tooltip_Err
        '</EhHeader>

100     Tooltip = nid.szTip

        '<EhFooter>
        Exit Property

Tooltip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.Tooltip " & _
               "at line " & Erl
        End
        '</EhFooter>
End Property

Public Property Let Interval(Value As Integer)
        '<EhHeader>
        On Error GoTo Interval_Err
        '</EhHeader>

100     TmrFlash.Interval = Value
102     UpdateIcon NIM_MODIFY

        '<EhFooter>
        Exit Property

Interval_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.Interval " & _
               "at line " & Erl
        End
        '</EhFooter>
End Property

Public Property Get Interval() As Integer
        '<EhHeader>
        On Error GoTo Interval_Err
        '</EhHeader>

100     Interval = TmrFlash.Interval

        '<EhFooter>
        Exit Property

Interval_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.Interval " & _
               "at line " & Erl
        End
        '</EhFooter>
End Property

Public Property Let TrayIcon(Value)
        '<EhHeader>
        On Error GoTo TrayIcon_Err
        '</EhHeader>

100     TmrFlash.Enabled = False
        On Error Resume Next
        ' Value can be a picturebox, image, form or string

102     Select Case TypeName(Value)

            Case "PictureBox", "Image"
104             Me.Icon = Value.Picture
106             TmrFlash.Enabled = False
108             RaiseEvent TIcon(Me)

110         Case "String"

112             If (UCase(Value) = "DEFAULT") Then

114                 TmrFlash.Enabled = True
116                 Me.Icon = Flash2.Picture
118                 RaiseEvent TIcon(Me)


                ElseIf (UCase(Value) = "PROCESSING") Then
                  TmrFlash.Enabled = False
                  Me.Icon = pictProcess.Picture
                  RaiseEvent TIcon(Me)
                  
                Else

                    ' Sting is filename; load icon from picture file.
120                 TmrFlash.Enabled = True
122                 Me.Icon = LoadPicture(Value)
124                 RaiseEvent TIcon(Me)

                End If

126         Case Else
                ' It's a form ?
128             Me.Icon = Value.Icon
130             RaiseEvent TIcon(Me)

        End Select

132     If Err.Number <> 0 Then TmrFlash.Enabled = True

134     UpdateIcon NIM_MODIFY

        '<EhFooter>
        Exit Property

TrayIcon_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.TrayIcon " & _
               "at line " & Erl
        End
        '</EhFooter>
End Property

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     Me.Icon = Flash1
102     RaiseEvent TIcon(Me)
104     Me.Visible = False
106     TmrFlash.Enabled = True
108     Tooltip = App.EXEName
110     mAbout.Caption = "About " & App.EXEName
112     UpdateIcon NIM_ADD

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.Form_Load " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo Form_MouseMove_Err
        '</EhHeader>

        Dim Result As Long
        Dim msg As Long
   
        ' The Form_MouseMove is intercepted to give systray mouse events.

100     If Me.ScaleMode = vbPixels Then

102         msg = X

        Else

104         msg = X / Screen.TwipsPerPixelX

        End If
      
106     Select Case msg

            Case WM_RBUTTONDBLCLK
108             RaiseEvent Click("RBUTTONDBLCLK")

110         Case WM_RBUTTONDOWN
112             RaiseEvent Click("RBUTTONDOWN")

114         Case WM_RBUTTONUP
                ' Popup menu: selectively enable items dependent on context.

116             Select Case FSys.Visible

                    Case True

118                     Select Case FSys.WindowState

                            Case vbMaximized
120                             mMaximize.Enabled = False
122                             mMinimize.Enabled = True
124                             mRestore.Enabled = False

126                         Case vbNormal
128                             mMaximize.Enabled = False
130                             mMinimize.Enabled = True
132                             mRestore.Enabled = False

134                         Case vbMinimized
136                             mMaximize.Enabled = False
138                             mMinimize.Enabled = False
140                             mRestore.Enabled = True

142                         Case Else
144                             mMaximize.Enabled = False
146                             mMinimize.Enabled = True
148                             mRestore.Enabled = True

                        End Select

150                 Case Else
152                     mRestore.Enabled = True
154                     mMaximize.Enabled = False
156                     mMinimize.Enabled = False

                End Select
         
158             RaiseEvent Click("RBUTTONUP")
160             PopupMenu mPopupMenu

162         Case WM_LBUTTONDBLCLK
164             RaiseEvent Click("LBUTTONDBLCLK")
166             mRestore_Click

168         Case WM_LBUTTONDOWN
170             RaiseEvent Click("LBUTTONDOWN")

172         Case WM_LBUTTONUP
174             RaiseEvent Click("LBUTTONUP")

176         Case WM_MBUTTONDBLCLK
178             RaiseEvent Click("MBUTTONDBLCLK")

180         Case WM_MBUTTONDOWN
182             RaiseEvent Click("MBUTTONDOWN")

184         Case WM_MBUTTONUP
186             RaiseEvent Click("MBUTTONUP")

188         Case WM_MOUSEMOVE
190             RaiseEvent Click("MOUSEMOVE")

192         Case Else
194             RaiseEvent Click("OTHER....: " & Format$(msg))

        End Select

        '<EhFooter>
        Exit Sub

Form_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.Form_MouseMove " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
   UpdateIcon NIM_DELETE
End Sub

Private Sub FSys_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    ' Event generated my main form. WindowState is stored in LastWindowState, so that
    ' it may be re- set when the menu item "Restore" is selected.

    If (FSys.WindowState <> vbMinimized) Then LastWindowState = FSys.WindowState

End Sub

Private Sub FSys_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo FSys_Unload_Err
        '</EhHeader>

        ' Important: remove icon from tray, and unload this form when
        ' the main form is unloaded.
100     UpdateIcon NIM_DELETE
102     Unload Me

        '<EhFooter>
        Exit Sub

FSys_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.FSys_Unload " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub mAbout_Click()
        '<EhHeader>
        On Error GoTo mAbout_Click_Err
        '</EhHeader>

   '     MsgBox "Ink Calculator. Copyright 2009, Alex Commandor (kalen@inbox.ru)", vbInformation, "About Ink Calculator"
         frmAbout.Show 1
        '<EhFooter>
        Exit Sub

mAbout_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.mAbout_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub mMaximize_Click()
        '<EhHeader>
        On Error GoTo mMaximize_Click_Err
        '</EhHeader>

100     FSys.WindowState = vbMaximized
102     FSys.Show

        '<EhFooter>
        Exit Sub

mMaximize_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.mMaximize_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub mMinimize_Click()
        '<EhHeader>
        On Error GoTo mMinimize_Click_Err
        '</EhHeader>

100     FSys.WindowState = vbMinimized

        '<EhFooter>
        Exit Sub

mMinimize_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.mMinimize_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Public Sub mExit_Click()
        '<EhHeader>
        On Error GoTo mExit_Click_Err
        '</EhHeader>

100     UpdateIcon NIM_DELETE
102     Unload FSys
104     End

        '<EhFooter>
        Exit Sub

mExit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.mExit_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub mRestore_Click()
        '<EhHeader>
        On Error GoTo mRestore_Click_Err
        '</EhHeader>

        ' Don't "restore"  FSys is visible and not minimized.

100     If (FSys.Visible And FSys.WindowState <> vbMinimized) Then Exit Sub

        ' Restore LastWindowState
102     FSys.WindowState = LastWindowState
104     FSys.Visible = True
106     SetForegroundWindow FSys.hWnd

        '<EhFooter>
        Exit Sub

mRestore_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.mRestore_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub UpdateIcon(Value As Long)
        '<EhHeader>
        On Error GoTo UpdateIcon_Err
        '</EhHeader>

        ' Used to add, modify and delete icon.

100     With nid

102         .cbSize = Len(nid)
104         .hWnd = Me.hWnd
106         .uID = vbNull
108         .uFlags = NIM_DELETE Or NIF_TIP Or NIM_MODIFY
110         .uCallbackMessage = WM_MOUSEMOVE
112         .hIcon = Me.Icon

        End With

114     Shell_NotifyIcon Value, nid

        '<EhFooter>
        Exit Sub

UpdateIcon_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.UpdateIcon " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Public Sub MeQueryUnload(ByRef F As Form, Cancel As Integer, UnloadMode As Integer)
        '<EhHeader>
        On Error GoTo MeQueryUnload_Err
        '</EhHeader>

100     If UnloadMode = vbFormControlMenu Then

            ' Cancel by setting Cancel = 1, minimize and hide main window.
102         Cancel = 1
104         F.WindowState = vbMinimized
106         F.Hide

        End If

        '<EhFooter>
        Exit Sub

MeQueryUnload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.MeQueryUnload " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Public Sub MeResize(ByRef F As Form)
        '<EhHeader>
        On Error GoTo MeResize_Err
        '</EhHeader>

100     Select Case F.WindowState

            Case vbNormal, vbMaximized
                ' Store LastWindowState
102             LastWindowState = F.WindowState

104         Case vbMinimized
106             F.Hide

        End Select

        '<EhFooter>
        Exit Sub

MeResize_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.MeResize " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub mStart_Click()
        '<EhHeader>
        On Error GoTo mStart_Click_Err
        '</EhHeader>

100     Call FSys.btnStart_Click

        '<EhFooter>
        Exit Sub

mStart_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTray.mStart_Click " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub TmrFlash_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    ' Change icon.
    Static LastIconWasFlash1 As Boolean
    LastIconWasFlash1 = Not LastIconWasFlash1

    Select Case LastIconWasFlash1

        Case True
            Me.Icon = Flash2

        Case Else
            Me.Icon = Flash1

    End Select

    RaiseEvent TIcon(Me)
    UpdateIcon NIM_MODIFY

End Sub

