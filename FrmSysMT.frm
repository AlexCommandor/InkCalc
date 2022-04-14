VERSION 5.00
Begin VB.Form FrmSysTrayMT 
   BorderStyle     =   0  'None
   ClientHeight    =   1965
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4710
   Icon            =   "FrmSysMT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox IcoProcess 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "FrmSysMT.frx":0742
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   720
      Width           =   540
   End
   Begin VB.Timer TimerDoCollect 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2040
      Top             =   120
   End
   Begin VB.PictureBox Flash2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   720
      Picture         =   "FrmSysMT.frx":0CCC
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
      Picture         =   "FrmSysMT.frx":1256
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Timer TmrFlash 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   120
   End
   Begin VB.Menu mPopupMenu 
      Caption         =   "&PopupMenu"
      Begin VB.Menu mExit 
         Caption         =   "&Abort current operation"
      End
      Begin VB.Menu mSep 
         Caption         =   "-"
      End
      Begin VB.Menu mContinue 
         Caption         =   "&Continue working"
      End
   End
End
Attribute VB_Name = "FrmSysTrayMT"
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

Private Enum nType
     PDSaveIncremental = 0  ' write changes only
     PDSaveFull = &H1 ' write entire file
     PDSaveCopy = &H2 ' write copy w/o affecting current state
     PDSaveLinearized = &H4 ' writes the file linearized
     PDSaveWithPSHeader = &H8 'writes a PostScript header as part of the saved file
     PDSaveBinaryOK = &H10 ' specifies that it's OK to save binary file
     PDSaveCollectGarbage = &H20 ' perform garbage collection on unreferenced objects
End Enum

'Public WithEvents FSys As Form
Public Event Click(ClickWhat As String)
Public Event TIcon(F As Form)

Private nid As NOTIFYICONDATA
Private LastWindowState As Integer
Private sInPath As String
Private sOutPath As String
Private sTimeOutFuck As String
Private aFSO As Object, aFO As Object, aFI As Object
Private sPages() As String

Public Property Let Tooltip(Value As String)
        '<EhHeader>
        On Error GoTo Tooltip_Err
        '</EhHeader>

100     nid.szTip = Value & vbNullChar

        '<EhFooter>
        Exit Property

Tooltip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTrayMT.Tooltip " & _
               "at line " & Erl
        UpdateIcon NIM_DELETE: End
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
               "in InkCalc.FrmSysTrayMT.Tooltip " & _
               "at line " & Erl
        UpdateIcon NIM_DELETE: End
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
               "in InkCalc.FrmSysTrayMT.Interval " & _
               "at line " & Erl
        UpdateIcon NIM_DELETE: End
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
               "in InkCalc.FrmSysTrayMT.Interval " & _
               "at line " & Erl
        UpdateIcon NIM_DELETE: End
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

134     Err.Clear
        On Error GoTo TrayIcon_Err
136     UpdateIcon NIM_MODIFY

        '<EhFooter>
        Exit Property

TrayIcon_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTrayMT.TrayIcon " & _
               "at line " & Erl
        UpdateIcon NIM_DELETE: End
        '</EhFooter>
End Property

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
   
        'On Error Resume Next

100     Me.Icon = Flash2
102     RaiseEvent TIcon(Me)
104     Me.Visible = False
'106     TmrFlash.Enabled = True
108     Tooltip = App.EXEName
110     UpdateIcon NIM_ADD
112     Me.TimerDoCollect.Enabled = True
   
        Dim iFNn As Integer
    
114     iFNn = FreeFile
116     Open App.Path & "\thread.txt" For Input As iFNn
118     Line Input #iFNn, sInPath
120     Line Input #iFNn, sOutPath
121     Line Input #iFNn, sTimeOutFuck

122     Close iFNn
124     Kill App.Path & "\thread.txt"
   
        '   If Err.Number <> 0 Then MsgBox "Error in Form_Load"

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTrayMT.Form_Load " & _
               "at line " & Erl
        UpdateIcon NIM_DELETE: End
        '</EhFooter>
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo Form_MouseMove_Err
        '</EhHeader>

        Dim Result As Long
        Dim msg As Long
   
        'On Error Resume Next

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
116             RaiseEvent Click("RBUTTONUP")
118             PopupMenu mPopupMenu

120         Case WM_LBUTTONDBLCLK
122             RaiseEvent Click("LBUTTONDBLCLK")
124             Result = MsgBox("Abort this thread?", vbQuestion + vbYesNo, "Ink Calculator")

126             If Result = vbYes Then Call mExit_Click

128         Case WM_LBUTTONDOWN
130             RaiseEvent Click("LBUTTONDOWN")

132         Case WM_LBUTTONUP
134             RaiseEvent Click("LBUTTONUP")

136         Case WM_MBUTTONDBLCLK
138             RaiseEvent Click("MBUTTONDBLCLK")

140         Case WM_MBUTTONDOWN
142             RaiseEvent Click("MBUTTONDOWN")

144         Case WM_MBUTTONUP
146             RaiseEvent Click("MBUTTONUP")

148         Case WM_MOUSEMOVE
150             RaiseEvent Click("MOUSEMOVE")

152         Case Else
154             RaiseEvent Click("OTHER....: " & Format$(msg))

        End Select

        'If Err.Number <> 0 Then MsgBox "Error in Form_MouseMove"

        '<EhFooter>
        Exit Sub

Form_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTrayMT.Form_MouseMove " & _
               "at line " & Erl
        UpdateIcon NIM_DELETE: End
        '</EhFooter>
End Sub

Public Sub mExit_Click()

On Error Resume Next
'    Set aFSO = CreateObject("Scripting.FileSystemObject")
'    Set aFO = aFSO.GetFolder(sInPath)
    
'    aFSO.DeleteFolder sInPath, True

    UpdateIcon NIM_DELETE
On Error GoTo 0
    End
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
               "in InkCalc.FrmSysTrayMT.UpdateIcon " & _
               "at line " & Erl
        UpdateIcon NIM_DELETE: End
        '</EhFooter>
End Sub

Private Sub TimerDoCollect_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    Me.TimerDoCollect.Enabled = False
    Call DoCollecting

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

Public Function DoCollecting() As Boolean '(ByVal sInPath As String, ByVal sOutPath As String) As Boolean
        '<EhHeader>
        On Error GoTo DoCollecting_Err
        '</EhHeader>

        Dim aTXTFile As String, timeNow As Date, i As Long, bRes As Boolean
        Dim nFiles As Integer, iFN As Integer, iFFNN As Integer


Set aFSO = CreateObject("Scripting.FileSystemObject")
Set aFO = aFSO.GetFolder(sInPath)
iFFNN = FreeFile

Open App.Path & "\" & aFO.Name & "_folder.log" For Append As iFFNN
    
100     If Right$(sInPath, Len(sInPath) - InStrRev(sInPath, "\")) = "New folder" Or _
           Right$(sInPath, Len(sInPath) - InStrRev(sInPath, "\")) = "Новая папка" Then

102         Call mExit_Click

        End If
   
'106     Me.Tooltip = "Waiting for PDF files in " & sInPath
107     UpdateIcon NIM_MODIFY
1071    DoEvents
    
        'On Error Resume Next
    
108     Set aFSO = CreateObject("Scripting.FileSystemObject")
110     Set aFO = aFSO.GetFolder(sInPath)


113     If Err.Number <> 0 Then _
            Print #iFFNN, Time() & Chr$(9) & "Start checking for files in " & sInPath & _
                " failed! Folder is missing???": Call mExit_Click 'Folder missing?????????????
        
1131    timeNow = Time()

115 Print #iFFNN, Time() & Chr$(9) & "Start checking for files in " & sInPath


            
117      aTXTFile = sOutPath & "\" & aFO.Name & ".txt"
119      If aFSO.FileExists(aTXTFile) Then Kill aTXTFile

         If aFO.Files.Count = 0 Then 'no files in folder!
            Print #iFFNN, Time() & Chr$(9) & "No files found in folder! Exiting..."
            Close #iFFNN
            Set aFI = Nothing
            Set aFO = Nothing
            Set aFSO = Nothing
            Call mExit_Click
         End If

121      iFN = FreeFile()

123      Open aTXTFile For Append As iFN
      
125      For Each aFI In aFO.Files
         
127 Print #iFFNN, Time() & Chr$(9) & "Analizing " & aFI.Path & " file."

129         If CalculateInks(aFI.Path) = True Then
131 Print #iFFNN, Time() & Chr$(9) & "CalculateInks for " & aFI.Path & " proceeded. Saving output..."
133            Print #iFN, Join(sPages, vbCrLf)
135 Print #iFFNN, Time() & Chr$(9) & "Saved data into " & aTXTFile & "."
137            aFSO.DeleteFolder App.Path & "\" & aFO.Name, True
139 Print #iFFNN, Time() & Chr$(9) & "Deleted folder " & App.Path & "\" & aFO.Name & "."
141         End If
            
            'Operation timeout !!!
145         If Minute(Time() - timeNow) > Val(sTimeOutFuck) Then
                Print #iFFNN, Time() & Chr$(9) & "Timeout waiting for files!!! Aborting ...   wait: " _
                    & Minute(Time() - timeNow) & ", timeout: " & sTimeOutFuck
                Call mExit_Click
            End If
            
      Next
      Close #iFN
      
156     Me.Icon = Me.IcoProcess
158     RaiseEvent TIcon(Me)
160     UpdateIcon NIM_MODIFY

Print #iFFNN, Time() & Chr$(9) & "Processed " & nFiles & "files."

      
210     Set aFI = Nothing
212     Set aFO = Nothing
'214     aFSO.DeleteFolder sInPath, True
216     Set aFSO = Nothing
    
Print #iFFNN, Time() & Chr$(9) & "Pizdets! :)"
    
218     Call mExit_Click

        '<EhFooter>
        Exit Function

DoCollecting_Err:

    Print #iFFNN, Time() & Chr$(9) & Err.Description & " at line " & Erl
    Close #iFFNN
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.FrmSysTrayMT.DoCollecting " & _
               "at line " & Erl
        UpdateIcon NIM_DELETE: End
        '</EhFooter>
End Function


Public Function CalculateInks(ByVal sFile As String) As Boolean
   'Const pcxKoeff As Double = 2.83464567
   Dim MyObj As GflAx.GflAx, W As Currency, H As Currency, i As Currency, j As Currency
   Dim SUM As Currency, FULL As Currency, DPI As Integer
   Dim Wcm As Currency, Hcm As Currency, FULLcm2 As Currency
   Dim tmp As Currency, FFSO As Object
   Dim bBuff() As Byte, WSH As Object
   
   On Error Resume Next
   
   DPI = 100
   
   Set MyObj = New GflAx.GflAx
   Set WSH = CreateObject("WScript.Shell")
   Set FFSO = CreateObject("Scripting.FileSystemObject")
   
   If GetLogicalPages(sFile) = False Then CalculateInks = False: Exit Function
   
  For i = LBound(sPages) To UBound(sPages)
  
   If WSH.Run(App.Path & "\pdftoppm.exe -aa no -aaVector no -gray -f " & CStr(i) & " -l " & CStr(i) & " -r " & _
         CStr(DPI) & " -q " & sFile & " " & sFile & "-page", 0&, True) <> 0 Then GoTo NNEXTT
   
   MyObj.LoadBitmap sFile & "-page-" & Format$(i, "000000") & ".pgm"
   
   W = MyObj.Width
   H = MyObj.Height

   FULL = W * H
   Wcm = W * 2.54 / DPI
   Hcm = H * 2.54 / DPI
   FULLcm2 = Wcm * Hcm
   
   bBuff = MyObj.SendBinary
   
   FFSO.DeleteFile sFile & "-page-" & Format$(i, "000000") & ".pgm", True
   
   SUM = 0
   
   For j = 0 To UBound(bBuff)
         SUM = SUM + 255 - bBuff(j)
   Next j
   
   FULL = FULL * 255
   
   sPages(i) = sFile & vbCrLf & vbCrLf & sPages(i) & " - page size " & Format$(Wcm, "0.00") & "cm x " & Format$(Hcm, "0.00") & _
            "cm, realtive ink coverage " & Format$(100 * SUM / FULL, "0.00") & _
            "%, area of ink coverage " & Format$(FULLcm2 * SUM / FULL, "0.00") & " cm2" & vbCrLf
NNEXTT:
  Next i
  Set MyObj = Nothing
  Set WSH = Nothing
  Set FFSO = Nothing
  
  If Err.Number = 0 Then CalculateInks = True Else Err.Clear: CalculateInks = False
  On Error GoTo 0
End Function

Private Function GetLogicalPages(ByVal sFile As String) As Boolean
   Dim FFSO As Object, TS As Object, sTmp As String, vArr() As String, i As Long
   
   On Error Resume Next
   
   Set FFSO = CreateObject("Scripting.FileSystemObject")
   Set TS = FSO.OpenTextFile(sFile)
   vArr = Split(Replace$(TS.ReadAll, vbCr, vbLf), vbLf)
   TS.Close
   Set TS = Nothing
   ReDim sPages(1 To 1)
   For i = LBound(vArr) To UBound(vArr)
      If (i < UBound(vArr) - 3) And (vArr(i) Like "/P (*") Then
         sTmp = Replace$(vArr(i), "/P (", vbNullString)
         sTmp = Left$(sTmp, Len(sTmp) - 1)
         i = i + 2
         sTmp = sTmp & "page " & Replace$(vArr(i), "/St ", vbNullString)
         sTmp = Replace$(sTmp, ":", " ")
         sPages(UBound(sPages)) = sTmp
         ReDim Preserve sPages(1 To UBound(sPages) + 1)
      End If
   Next i
   ReDim Preserve sPages(1 To UBound(sPages) - 1)
   Set FFSO = Nothing
   If Err.Number = 0 Then GetLogicalPages = True Else Err.Clear: GetLogicalPages = False
   On Error GoTo 0
End Function

Private Function BitsSum(ByVal bByte As Byte) As Byte
   Dim bSum As Byte
   bSum = 0
   bSum = bSum + (bByte And &H1)
   bSum = bSum + (bByte And &H2) \ &H2
   bSum = bSum + (bByte And &H4) \ &H4
   bSum = bSum + (bByte And &H8) \ &H8
   bSum = bSum + (bByte And &H10) \ &H10
   bSum = bSum + (bByte And &H20) \ &H20
   bSum = bSum + (bByte And &H40) \ &H40
   bSum = bSum + (bByte And &H80) \ &H80
   BitsSum = bSum
End Function



