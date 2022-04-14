Attribute VB_Name = "modMain"

Option Explicit


Const WM_NCACTIVATE = &H86
Const SC_CLOSE = &HF060
Const MF_BYCOMMAND = &H0&
Const TXT_FILE_WITH_RESULTS As String = "IncCalc.cvs"

Const TXT_INI_FILE As String = "inkcalc.ini"

Public Const GROUPS_MAX_QUANTITY = &H9

' Structure to hold IDispatch GUID
Public Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

'new parameter - group koefficients by paper type :(
Public numGroups As Byte

Public bInProcessing As Boolean
Public sInputFolder() As String
Public sGroupName() As String
Public lLANG As Integer
Public bUseOneFolder As Boolean

Public sCurrentWorkFolder As String
Public sOutputFolder As String
Public sTimeOut As String, sRemoveOlderThan As String
Public bCleanupAfter As Boolean
'Public sMass_Coeff(1 To 2) As Variant
Public sMass_Coeff_ROLL() As Variant
Public sMass_Coeff_LIST As Variant
Public sResultFile As String
Public sListPlates As Variant
Public FSO As Object, FO As Object, FI As Object, sPages() As String
'Public sINP As String, sOUTP As String, sTIME_OUT_FILE As String, sCLEANUP_AFTER As String
'Public sCOEFF As String, sLISTSIZES As String
Dim iFN As Integer

Public fMain As Form

'Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, _
        ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, _
            ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
            ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As GUID)
Public Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal rguid As Long) As Long

Public Declare Function WaitForMultipleObjects Lib "kernel32" _
  (ByVal nCount As Long, lpHandles As Any, ByVal bWaitAll As Long, _
  ByVal dwMilliseconds As Long) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Declare Function FindFirstChangeNotification Lib "kernel32" _
  Alias "FindFirstChangeNotificationA" _
  (ByVal lpPathName As String, ByVal bWatchSubtree As Long, _
   ByVal dwNotifyFilter As EFILE_NOTIFY) As Long

Public Declare Function FindNextChangeNotification Lib "kernel32" _
  (ByVal hChangeHandle As Long) As Long

Public Declare Function FindCloseChangeNotification Lib "kernel32" _
  (ByVal hChangeHandle As Long) As Long

Public Enum EFILE_NOTIFY
    FILE_NOTIFY_CHANGE_FILE_NAME = &H1
    FILE_NOTIFY_CHANGE_DIR_NAME = &H2
    FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
    FILE_NOTIFY_CHANGE_SIZE = &H8
    FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
    FILE_NOTIFY_CHANGE_LAST_ACCESS = &H20
    FILE_NOTIFY_CHANGE_CREATION = &H40
    FILE_NOTIFY_CHANGE_SECURITY = &H100
End Enum

Public Const WAIT_TIMEOUT = 258
Public Const WAIT_FAILED = -1
Public Const INVALID_HANDLE_VALUE = -1

Sub Main()
        '<EhHeader>
        On Error GoTo Main_Err
        '</EhHeader>

        Dim hWnd As Long
        Dim ProcessID As Long, curProcessID As Long
        
        If App.PrevInstance Then End
        

114      Call Init


        '<EhFooter>
        Exit Sub

Main_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.modMain.Main " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Private Sub Init()
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>

        Dim A As Object, i As Byte
    
        On Error Resume Next
100     Set A = CreateObject("GflAx.GflAx")

102     If Err.Number <> 0 Then 'GflAx missing???

104         Err.Clear
106         MsgBox sLangStrings(StringIDs.remErrAccGflAxInstCorr) & _
               sLangStrings(StringIDs.remTryToGoOnGflAx) & _
               sLangStrings(StringIDs.remAndInstallGflAxManually), vbCritical, sLangStrings(StringIDs.remIncCalcERR)
108         End

        End If
         
         Set A = Nothing
         
110     Set A = CreateObject("Scripting.FileSystemObject")

112     If Err.Number <> 0 Then 'Scriptingmissing???

114         Err.Clear
116         MsgBox sLangStrings(StringIDs.remErrAccMSSCRRUN) & _
            sLangStrings(StringIDs.remTryToGoOnMSSCRRUN) & _
            sLangStrings(StringIDs.remAndInstallItManually), vbCritical, sLangStrings(StringIDs.remIncCalcERR)
118         End

        End If
    
120     Set A = Nothing

      Set A = CreateObject("WScript.Shell")

     If Err.Number <> 0 Then 'Scriptingmissing???

         Err.Clear
            MsgBox sLangStrings(StringIDs.remErrAccWSH) & _
            sLangStrings(StringIDs.remTryToGoOnMSWSH) & _
            sLangStrings(StringIDs.remAndInstallItManually), vbCritical, sLangStrings(StringIDs.remIncCalcERR)
         End

        End If
    
     Set A = Nothing

    
'125 Kill App.Path & "\*_folder.log"
    
126     Set FSO = CreateObject("Scripting.FileSystemObject")
         FSO.DeleteFile App.Path & "\*_folder.log", True
         FSO.DeleteFile App.Path & "\*_process.txt", True
         For Each FO In FSO.GetFolder(App.Path).SubFolders
            FSO.DeleteFolder FO.Path, True
         Next
 
    Dim arrData() As Byte
    If Not FSO.FileExists(App.Path & "\pdftoppm.exe") Then
        arrData = LoadResData(101, "CUSTOM")
        Open App.Path & "\pdftoppm.exe" For Binary Access Write As #1
            Put #1, , arrData
        Close #1
    End If
      
    Erase arrData
    If Not FSO.FileExists(App.Path & "\pdfimages.exe") Then
        arrData = LoadResData(102, "CUSTOM")
        Open App.Path & "\pdfimages.exe" For Binary Access Write As #1
            Put #1, , arrData
        Close #1
    End If
 
Err.Clear

        On Error GoTo Init_Err
 
'        sINP = App.Path & "\" & TXT_INPUT
'        sOUTP = App.Path & "\" & TXT_OUTPUT
'        sTIME_OUT_FILE = App.Path & "\" & TXT_TIMEOUT
'        sCLEANUP_AFTER = App.Path & "\" & TXT_CLEANUP_AFTER
'        sCOEFF = App.Path & "\" & TXT_COEFF
'        sLISTSIZES = App.Path & "\" & TXT_LIST_SIZES
        If Not modINI.InitINI(TXT_INI_FILE) Then
            MsgBox sLangStrings(StringIDs.remErrCreateINIChekAndTryAgain), _
                  vbCritical, sLangStrings(StringIDs.remIncCalcInitError)
            End
        End If
            
132     If ReadHotFolders = False Then

134         MsgBox sLangStrings(StringIDs.remDiskAccErr), _
               vbCritical + vbOKOnly, sLangStrings(StringIDs.remIncCalcInitError)
136         End

        End If

        Call EnsureHotFoldersExists
            
        If Right$(sOutputFolder, 1) = "\" Then
            sResultFile = sOutputFolder & TXT_FILE_WITH_RESULTS
         Else
            sResultFile = sOutputFolder & "\" & TXT_FILE_WITH_RESULTS
        End If
        
146     bInProcessing = False
148     Set fMain = New InkCalculator
150     Load fMain
152     DisableCloseButton fMain.hWnd
154     Set fMain = Nothing

        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.modMain.Init " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Public Sub EnsureHotFoldersExists()
    Dim i As Byte
    For i = 1 To GROUPS_MAX_QUANTITY
        If Not FSO.FolderExists(sInputFolder(i)) Then
            sInputFolder(i) = "C:\"
        End If
    Next i
    If Not FSO.FolderExists(sOutputFolder) Then
        sOutputFolder = "C:\"
    End If
End Sub

Public Function ReadHotFolders() As Boolean

    Dim sTmp As String, i As Byte
         
    On Error Resume Next
    
    numGroups = ReadINI("GroupsQuantity", "3")
    
    If numGroups <= 0 Or numGroups > 9 Then numGroups = 3
    
    ReDim sInputFolder(1 To GROUPS_MAX_QUANTITY)
    ReDim sGroupName(1 To GROUPS_MAX_QUANTITY)
    ReDim sMass_Coeff_ROLL(1 To GROUPS_MAX_QUANTITY)
    For i = 1 To GROUPS_MAX_QUANTITY
        sGroupName(i) = ReadINI("GroupName" & CStr(i), "Group " & CStr(i))
        sInputFolder(i) = ReadINI("InputFolder" & CStr(i), "C:\")
        sTmp = ReadINI("MassCoeffROLL" & CStr(i), "1.55" & vbTab & "1.55" & vbTab & "1.95" & vbTab & "1.35" & vbTab & "2.00")
        sMass_Coeff_ROLL(i) = Split(sTmp, vbTab)
    Next i
    
    lLANG = ReadINI("Language", "1000")
    sOutputFolder = ReadINI("OutputFolder", "C:\")
    sTimeOut = ReadINI("TimeOut", "30")
    sRemoveOlderThan = ReadINI("RemoveOlderThan", "60")
    sTmp = ReadINI("ListPlates")
    If Len(sTmp) > 0 Then sListPlates = Split(sTmp, vbTab) Else Erase sListPlates
    bCleanupAfter = CBool(ReadINI("CleanupAfter", "True"))
    bUseOneFolder = CBool(ReadINI("UseOneFolder", "False"))
    'sTmp = ReadINI("MassCoeffROLL", "1.55" & vbTab & "1.55" & vbTab & "1.95" & vbTab & "1.35" & vbTab & "2.00")
    'sMass_Coeff(1) = Split(sTmp, vbTab)
    sTmp = ReadINI("MassCoeffLIST", "0.90" & vbTab & "0.90" & vbTab & "1.10" & vbTab & "1.00" & vbTab & "2.00")
    sMass_Coeff_LIST = Split(sTmp, vbTab)
       
    If Right$(sOutputFolder, 1) = "\" Then
        sResultFile = sOutputFolder & TXT_FILE_WITH_RESULTS
    Else
        sResultFile = sOutputFolder & "\" & TXT_FILE_WITH_RESULTS
    End If

    If Err.Number <> 0 Then
        ReadHotFolders = False
    Else
        ReadHotFolders = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Public Function WriteHotFolders() As Boolean
    Dim bRes As Boolean, i As Byte
    On Error Resume Next
    bRes = True
    '-------
    bRes = bRes And WriteINI("GroupsQuantity", CStr(numGroups))
    '-------
    For i = 1 To GROUPS_MAX_QUANTITY
        bRes = bRes And WriteINI("GroupName" & CStr(i), sGroupName(i))
        bRes = bRes And WriteINI("InputFolder" & CStr(i), sInputFolder(i))
        bRes = bRes And WriteINI("MassCoeffROLL" & CStr(i), Join(sMass_Coeff_ROLL(i), vbTab))
    Next i
    bRes = bRes And WriteINI("Language", lLANG)
    bRes = bRes And WriteINI("OutputFolder", sOutputFolder)
    bRes = bRes And WriteINI("TimeOut", sTimeOut)
    bRes = bRes And WriteINI("RemoveOlderThan", sRemoveOlderThan)
    bRes = bRes And WriteINI("CleanupAfter", CStr(bCleanupAfter))
    bRes = bRes And WriteINI("UseOneFolder", CStr(bUseOneFolder))
    'bRes = bRes And WriteINI("MassCoeffROLL", Join(sMass_Coeff_ROLL, vbTab))
    bRes = bRes And WriteINI("MassCoeffLIST", Join(sMass_Coeff_LIST, vbTab))
    If IsArray(sListPlates) Then
        bRes = bRes And WriteINI("ListPlates", Join(sListPlates, vbTab))
    Else
        bRes = bRes And WriteINI("ListPlates", vbNullString)
    End If

    If Right$(sOutputFolder, 1) = "\" Then
        sResultFile = sOutputFolder & TXT_FILE_WITH_RESULTS
    Else
        sResultFile = sOutputFolder & "\" & TXT_FILE_WITH_RESULTS
    End If
    If bRes = True And Err.Number = 0 Then
        WriteHotFolders = True
    Else
        WriteHotFolders = False
    End If
    Err.Clear
    On Error GoTo 0
End Function

Public Sub DisableCloseButton(ByVal FormHWND As Long)
        '<EhHeader>
        On Error GoTo DisableCloseButton_Err
        '</EhHeader>

        Dim hMenu As Long, Success As Long
100     hMenu = GetSystemMenu(FormHWND, 0)
102     Success = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
104     SendMessage FormHWND, WM_NCACTIVATE, 0&, 0&
106     SendMessage FormHWND, WM_NCACTIVATE, 1&, 0

        '<EhFooter>
        Exit Sub

DisableCloseButton_Err:
        MsgBox Err.Description & vbCrLf & _
               "in InkCalc.modMain.DisableCloseButton " & _
               "at line " & Erl
        End
        '</EhFooter>
End Sub

Public Function GenerateRandomGUID() As String
    Dim myGUID As GUID, NewGUID As GUID
    Dim GUIDByte() As Byte, sGUID As String
    Dim GuidLen As Long
    
    CoCreateGuid myGUID
    
    ReDim GUIDByte(80)
    GuidLen = StringFromGUID2(VarPtr(myGUID.Data1), VarPtr(GUIDByte(0)), UBound(GUIDByte))
    
    sGUID = Left(GUIDByte, GuidLen)
    
    
    GuidLen = CLSIDFromString(StrPtr(sGUID), VarPtr(NewGUID.Data1))
    
    If Asc(Right$(sGUID, 1)) = 0 Then sGUID = Left$(sGUID, Len(sGUID) - 1)
    
    GenerateRandomGUID = sGUID
End Function

Public Function GetGUIDfromString(ByVal sGUID As String, ByRef Result As GUID) As Boolean
    Dim NewGUID As GUID
    Dim GuidRes As Long
    GuidRes = CLSIDFromString(StrPtr(sGUID), VarPtr(NewGUID.Data1))
    If GuidRes = 0 Then
        Result = NewGUID
        GetGUIDfromString = True
    Else
        GetGUIDfromString = False
    End If
End Function

Public Function GetStringFromGUID(ByRef tGUID As GUID) As String
    Dim myGUID As GUID
    Dim GUIDByte() As Byte, sGUID As String
    Dim GuidLen As Long
    ReDim GUIDByte(80)
    myGUID = tGUID
    GuidLen = StringFromGUID2(VarPtr(myGUID.Data1), VarPtr(GUIDByte(0)), UBound(GUIDByte))
    sGUID = Left(GUIDByte, GuidLen)
    If Asc(Right$(sGUID, 1)) = 0 Then sGUID = Left$(sGUID, Len(sGUID) - 1)
    GetStringFromGUID = sGUID
End Function
