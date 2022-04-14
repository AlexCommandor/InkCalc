Attribute VB_Name = "modLANG"
Option Explicit
Option Compare Text

Public Enum LANGUAGE
    ENG = 1000
    RUS = 2000
End Enum

Public Enum StringIDs
    resInkCalcCaption = 1
    resPDFMustBe = 2
    resDblClkToChangeName = 3
    resOutFolderPath = 4
    resRemPDFafter = 5
    resSelOutFolder = 6
    remTimeout = 7
    remRemRecsOlderThen = 8
    remDays = 9
    remKoefSheet = 10
    remSheetPlateSizes = 11
    remAddNew = 12
    remRemChecked = 13
    remSaveSett = 14
    remStartProc = 15
    remExitProg = 16
    remHideThis = 17
    remAbout = 18
    remGroupQuant = 19
    remK = 20
    remC = 21
    remM = 22
    remY = 23
    remO = 24
    remSelInFolder = 25
    remUseOneInout = 26
    remInThisCase = 27
    remWatchFolder = 28
    remKoef = 29
    remForRoll = 30
    remStopProc = 31
    
    remSomeFoldStillTheSame = 101
    remPleaseSelectProperFolders = 102
    remIncCalcWarn = 103
    remSomeFoldersTheSame = 104
    remSomeFoldersIsRoot = 105
    remItsImpossAllWillDel = 106
    remIncCalcERR = 107
    remWatchFold = 108
    remEnterWidth = 109
    remInkCalcAddNewPlate = 110
    remYouAreAboutWrongNumber = 111
    remDoYouAbortAddNewSize = 112
    remEnterHeight = 113
    remChNameOfSelGroup = 114
    remInkCalcGrControl = 115
    remAnalFile = 116
    remUnableWriteFile = 117
    remMaybeItLockCloseAndTry = 118
    remErrAccGflAxInstCorr = 119
    remTryToGoOnGflAx = 120
    remAndInstallGflAxManually = 121
    remErrAccMSSCRRUN = 122
    remTryToGoOnMSSCRRUN = 123
    remAndInstallItManually = 124
    remErrAccWSH = 125
    remTryToGoOnMSWSH = 126
    remErrCreateINIChekAndTryAgain = 127
    remIncCalcInitError = 128
    remDiskAccErr = 129
    remBrowseForFolder = 130
End Enum


Public sLangStrings(1 To 200) As String

Public Enum kbLayout
  kbdENG = 0
  kbdRUS = 1
  kbdUKR = 2
End Enum

     
Private Const LOCALE_SDECIMAL = &HE         '  decimal separator
Private Const LOCALE_SLIST = &HC                  '  list item separator

Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" _
  (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
 
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
Private Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long

Public Sub ChangeLanguage(ByVal LangID As LANGUAGE, ByRef fForm As Form)
    Dim i As Integer
    
    On Error Resume Next 'we must turn off errors because we dont have all strings filled
    For i = 1 To UBound(sLangStrings)
        sLangStrings(i) = LoadResString(LangID + i)
        'If Err.Number > 0 Then Exit For ' no more strings in resource - exiting
    Next i
    Err.Clear
    On Error GoTo 0
    
    With fForm
        .Caption = sLangStrings(StringIDs.resInkCalcCaption)
        .Label4.AutoSize = True
        .Label4.Caption = sLangStrings(StringIDs.resPDFMustBe)
        .Text2.Text = sLangStrings(StringIDs.resDblClkToChangeName)
        .Label2.AutoSize = True
        .Label2.Caption = sLangStrings(StringIDs.resOutFolderPath)
        .chkCleanupAfter.Caption = sLangStrings(StringIDs.resRemPDFafter)
        .btnSelectOutput.Caption = sLangStrings(StringIDs.resSelOutFolder)
        .Label3.AutoSize = True
        .Label3.Caption = sLangStrings(StringIDs.remTimeout)
        .Label6.AutoSize = True
        .Label6.Caption = sLangStrings(StringIDs.remRemRecsOlderThen)
        .Label7.AutoSize = True
        .Label7.Caption = sLangStrings(StringIDs.remDays)
        .Frame2.Caption = sLangStrings(StringIDs.remKoefSheet)
        .Frame3.Caption = sLangStrings(StringIDs.remSheetPlateSizes)
        .cmdAddListSize.Caption = sLangStrings(StringIDs.remAddNew)
        .cmdRemoveListSize.Caption = sLangStrings(StringIDs.remRemChecked)
        .cmdSave.Caption = sLangStrings(StringIDs.remSaveSett)
        .btnStart.Caption = sLangStrings(StringIDs.remStartProc)
        .cmdDiscard.Caption = sLangStrings(StringIDs.remExitProg)
        .cmdHide.Caption = sLangStrings(StringIDs.remHideThis)
        .cmdAbout.Caption = sLangStrings(StringIDs.remAbout)
        .txtGroupsQ.Text = sLangStrings(StringIDs.remGroupQuant)
        For i = 0 To (modMain.GROUPS_MAX_QUANTITY + 1) * 5 - 1 Step 5
            .Label5(i).AutoSize = True
            .Label5(i + 1).AutoSize = True
            .Label5(i + 2).AutoSize = True
            .Label5(i + 3).AutoSize = True
            .Label5(i + 4).AutoSize = True
            .Label5(i).Caption = sLangStrings(StringIDs.remC)
            .Label5(i + 1).Caption = sLangStrings(StringIDs.remM)
            .Label5(i + 2).Caption = sLangStrings(StringIDs.remY)
            .Label5(i + 3).Caption = sLangStrings(StringIDs.remK)
            .Label5(i + 4).Caption = sLangStrings(StringIDs.remO)
        Next i
        .chkUseOneHotfolder.Caption = sLangStrings(StringIDs.remUseOneInout)
        .txtOneHot.Text = sLangStrings(StringIDs.remInThisCase)
        
        For i = 0 To modMain.GROUPS_MAX_QUANTITY - 1
            .btnSelectInput(i).Caption = sLangStrings(StringIDs.remSelInFolder) & " N" & CStr(i + 1)
            .Label1(i).Caption = sLangStrings(StringIDs.remWatchFolder) & " N" & CStr(i + 1) & ":"
            .Frame1(i).Caption = sLangStrings(StringIDs.remKoef) & " N" & CStr(i + 1) & sLangStrings(StringIDs.remForRoll)
        Next i
        
        .Refresh
    End With
    
   
End Sub

Public Function UC_NZ(chkString, Optional DefaultValue As String = vbNullString) As String
  Dim SS As String
  On Error GoTo ErrHand
  If Not IsNull(chkString) Then
    SS = CStr(chkString)
  Else
    SS = DefaultValue
  End If
  UC_NZ = SS
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Sub SetKbdrdLayout(KeybLayout As kbLayout)
  Dim i As Long
  On Error GoTo ErrHand
  Select Case KeybLayout
    Case kbdENG
      i = LoadKeyboardLayout("00000409", 1)
    Case kbdRUS
      i = LoadKeyboardLayout("00000419", 1)
    Case kbdUKR
      i = LoadKeyboardLayout("00000422", 1)
  End Select
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Public Sub SetDecimalSeparator(ByVal sDecSeparator As String)
  Dim iLocale As Integer, sTmpStr As String, lRes As Long
  On Error GoTo ErrHand
  If Len(sDecSeparator) = 0 Then Exit Sub
  sTmpStr = sDecSeparator
  If Len(sTmpStr) > 4 Then sTmpStr = Left$(sTmpStr, 4)
  sTmpStr = sTmpStr & Chr$(0)
  iLocale = GetUserDefaultLCID()
  lRes = SetLocaleInfo(iLocale, LOCALE_SDECIMAL, sTmpStr)
  If lRes = 0 Then MsgBox "Ошибка изменения разделителя целой и дробной части на «" _
      & sDecSeparator & "».", vbCritical
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Public Function GetDecimalSeparator() As String
  Dim iLocale As Integer, sTmpStr As String * 4, lRes As Long
  On Error GoTo ErrHand
  iLocale = GetUserDefaultLCID()
  lRes = GetLocaleInfo(iLocale, LOCALE_SDECIMAL, sTmpStr, 0)
  lRes = GetLocaleInfo(iLocale, LOCALE_SDECIMAL, sTmpStr, lRes)
  sTmpStr = Replace$(sTmpStr, Chr$(0), "", , 4)
  GetDecimalSeparator = Trim$(sTmpStr)
  If lRes = 0 Then MsgBox "Ошибка получения разделителя целой и дробной части!", vbCritical
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Sub SetListSeparator(ByVal sListSeparator As String)
  Dim iLocale As Integer, sTmpStr As String, lRes As Long
  On Error GoTo ErrHand
  If Len(sListSeparator) = 0 Then Exit Sub
  sTmpStr = sListSeparator
  If Len(sTmpStr) > 4 Then sTmpStr = Left$(sTmpStr, 4)
  sTmpStr = sTmpStr & Chr$(0)
  iLocale = GetUserDefaultLCID()
  lRes = SetLocaleInfo(iLocale, LOCALE_SLIST, sTmpStr)
  If lRes = 0 Then MsgBox "Ошибка изменения разделителя списков на «" _
      & sListSeparator & "».", vbCritical
  Exit Sub
ErrHand:
  ErrMsgO Err, , True
End Sub

Public Function GetListSeparator() As String
  Dim iLocale As Integer, sTmpStr As String * 4, lRes As Long
  On Error GoTo ErrHand
  iLocale = GetUserDefaultLCID()
  lRes = GetLocaleInfo(iLocale, LOCALE_SLIST, sTmpStr, 0)
  lRes = GetLocaleInfo(iLocale, LOCALE_SLIST, sTmpStr, lRes)
  sTmpStr = Replace$(sTmpStr, Chr$(0), "", , 4)
  GetListSeparator = Trim$(sTmpStr)
  If lRes = 0 Then MsgBox "Ошибка получения разделителя списков!", vbCritical
  Exit Function
ErrHand:
  ErrMsgO Err, , True
End Function

Public Sub ErrMsgO(ByRef oErr As ErrObject, _
      Optional ByVal sMsgCaption As String = "Ошибка!", _
      Optional bStop As Boolean = False)
  MsgBox "Ошибка выполнения <" & oErr.Description & ">, код " & _
      oErr.Number & "." & vbCrLf & _
      "По старой доброй привычке считаем это недопустимой операцией и закрываемся." & vbCrLf & _
      "Если слабо разобраться в причинах ошибки, обратитесь к разработчику :) .", _
      vbCritical, sMsgCaption
  If bStop Then End
End Sub



