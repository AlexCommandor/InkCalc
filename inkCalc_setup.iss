; 'D:\!_Work\Alex_K\Progs\InkCalc\Package\Support\Setup.Lst' imported by ISTool version 5.2.1

[Setup]
AppName=Ink Coverage Calculator
AppVerName=Ink Coverage Calculator
DefaultDirName={pf}\Ink Calculator
DefaultGroupName=InkCalc
PrivilegesRequired=admin

[Files]
; [Bootstrap Files]
; @COMCAT.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,5/31/98 1:00:00 AM,22288,4.71.1460.1
Source: COMCAT.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @STDOLE2.TLB,$(WinSysPathSysFile),$(TLBRegister),,6/3/99 1:00:00 AM,17920,2.40.4275.1
Source: STDOLE2.TLB; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
; @ASYCFILT.DLL,$(WinSysPathSysFile),,,3/8/99 1:00:00 AM,147728,2.40.4275.1
Source: ASYCFILT.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @OLEPRO32.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,3/8/99 1:00:00 AM,164112,5.0.4275.1
Source: OLEPRO32.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @OLEAUT32.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,4/12/00 1:00:00 AM,598288,2.40.4275.1
Source: OLEAUT32.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @msvbvm60.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,4/14/08 2:00:00 PM,1384479,6.0.98.2
Source: msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver

; [Setup1 Files]
; @GflAx.dll,$(WinSysPath),$(DLLSelfRegister),$(Shared),2/27/08 5:32:14 PM,1167360,2.82.0.0
Source: GflAx.dll; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @InkCalc.exe,$(AppPath),,,5/14/09 4:38:43 PM,1982464,2.0.0.3
Source: InkCalc.exe; DestDir: {app}; Flags: promptifolder

[Icons]
Name: {group}\InkCalc; Filename: {app}\InkCalc.exe; WorkingDir: {app}
