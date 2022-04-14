Attribute VB_Name = "modINI"
Option Explicit

Public Enum enumINIPlacement
   AppPath ' Place INI file in application directory
   WinPath ' Place INI file in windows directory
End Enum

Private sINIFileName As String ' full path to INI file (path + name + ext)

' InitINI recievs ONLY file name+ext (INIFileName) and its placement (App or Win dir)
' and generate inner full INI path (sINIFileName)
Public Function InitINI(ByVal INIFileName As String, _
      Optional ByVal INIPlacemnet As enumINIPlacement = AppPath) As Boolean
   Dim sPath As String
   If Len(INIFileName) = 0 Then
      ' if creating failed - stop processing :(
      InitINI = False
      Exit Function
   End If
   If INIPlacemnet = AppPath Then
      sPath = App.Path
   Else
      sPath = Environ("windir")
   End If
   ' we must be shure that path is ended with "\"
   If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
   ' create full INI path
   sINIFileName = sPath & INIFileName
   ' check if INI is exists
   If Not FileExists(sINIFileName) Then
      ' if not exists - trying to create
      If Not CreateFile(sINIFileName) Then
         ' if creating failed - stop processing :(
         InitINI = False
         Exit Function
      End If
   End If
   InitINI = True
End Function

' Simply func to check file availability. We just trying to open file for reading
' and if no error then file is exists
Public Function FileExists(ByVal sFileName As String) As Boolean
   Dim iNN As Integer
   On Error Resume Next
   iNN = FreeFile
   Open sFileName For Input As iNN
   Close iNN
   If Err.Number <> 0 Then FileExists = False Else FileExists = True
   Err.Clear
   On Error GoTo 0
End Function

' Simply func to (re)create file. We just trying to open file for writing
' and closing it immediatly. If no errors - file is created successfully
Public Function CreateFile(ByVal sFileName As String) As Boolean
   Dim iNN As Integer
   On Error Resume Next
   iNN = FreeFile
   Open sFileName For Output As iNN
   Close iNN
   If Err.Number <> 0 Then CreateFile = False Else CreateFile = True
   Err.Clear
   On Error GoTo 0
End Function

'Reads one variable from INI file. If no success - return empty string
Public Function ReadINI(ByVal sVariable As String, Optional ByVal sValueIfEmpty As String = vbNullString) As String
   Dim sData As String, iNNN As Integer
   If Not FileExists(sINIFileName) Then
      MsgBox "INI file is not initialised! Aborting...", vbCritical, "InkCalc INI error"
      End
   End If
   ' if file is empty or var is empty - nothing to do :(
   If FileLen(sINIFileName) = 0 Then
        If Len(sVariable) = 0 Then
            ReadINI = vbNullString
        Else
            ReadINI = sValueIfEmpty
        End If
        Exit Function
   End If
   'opens INI file and reads line by line until found our var name
   iNNN = FreeFile
   sData = vbNullString
   Open sINIFileName For Input As iNNN
      Do
         Line Input #iNNN, sData
         If InStr(UCase$(sData), UCase$(sVariable & " = ")) = 1 Then
            'we found string with var name!
            'Trim left part of string, because it containts var name, not var value
            sData = Right$(sData, Len(sData) - Len(sVariable) - 3)
         Else
            sData = vbNullString
         End If
         'loops until end of file or until var name was found
      Loop Until EOF(iNNN) Or Len(sData) > 0
   Close iNNN
   If Len(sValueIfEmpty) > 0 And Len(sData) = 0 Then
      ReadINI = sValueIfEmpty
   Else
      ReadINI = sData
   End If
End Function

Public Function WriteINI(ByVal sVariableName As String, ByVal sVariableData As String) As Boolean
   Dim sData() As String, iNNN As Integer, i As Long, bVarFound As Boolean
   If Not FileExists(sINIFileName) Then
      MsgBox "INI file is not initialised! Aborting...", vbCritical, "InkCalc INI error"
      End
   End If
   If Len(sVariableName) = 0 Then
      WriteINI = False
      Exit Function
   End If
   iNNN = FreeFile
   ReDim sData(1 To 1)
   bVarFound = False
   ' if file is not empty - read ALL data from file and REWRITE INI with new var value
   If FileLen(sINIFileName) > 0 Then
      ' if INI file is early initialized -
      ' open it and reads ALL lines, correct our var name and write out clearly
      i = 1
      Open sINIFileName For Input As iNNN
         Do
            ReDim Preserve sData(1 To i)
            Line Input #iNNN, sData(i)
            If InStr(sData(i), " = ") > 1 Then 'check line for structure VARNAME = BLABLABLA
               If InStr(UCase$(sData(i)), UCase$(sVariableName & " = ")) = 1 Then
                  If Not bVarFound Then ' if it is first var position
                     'we found string with var name!
                     'Trim left part of string, because it containts var name, not var value
                     sData(i) = sVariableName & " = " & sVariableData
                     ' bVarFound is indicate than var is found and when we found same name we must skip it
                     bVarFound = True
                  Else ' we found a duplicate of var - delete it!
                     i = i - 1
                     ReDim Preserve sData(1 To i)
                  End If
               End If
            Else 'we found some strange data, not var - clear it
               i = i - 1
               ReDim Preserve sData(1 To i)
            End If
            i = i + 1
         'loops until end of file
         Loop Until EOF(iNNN)
      Close iNNN
      If Not bVarFound Then ' we dont found any var name - var is NEW, just add it
         ReDim Preserve sData(1 To i)
         sData(i) = sVariableName & " = " & sVariableData
      End If
   Else ' INI file is still empty - just write out new FIRST val
      sData(1) = sVariableName & " = " & sVariableData
   End If
   'rewrite clean data stream
   iNNN = FreeFile
   Open sINIFileName For Output As iNNN
      Print #iNNN, Join$(sData, vbCrLf)
   Close iNNN
   WriteINI = True
End Function



