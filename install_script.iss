;InnoSetupVersion=4.2.6

[Setup]
AppName=Malcode Analyst Pack
AppVerName=Malcode Analyst Pack v0.24
DefaultDirName=c:\iDefense\MAP\
DefaultGroupName=Malcode Analyst Pack
OutputBaseFilename=./map_setup
OutputDir=./


[Files]
Source: ./dependancies\vbDevKit.dll; DestDir: {win}; Flags: regserver
Source: ./dependancies\spSubclass2.dll; DestDir: {win}; Flags: regserver
Source: ./dependancies\MSWINSCK.OCX; DestDir: {win}; Flags: uninsneveruninstall regserver promptifolder
Source: ./dependancies\mscomctl.ocx; DestDir: {win}; Flags: uninsneveruninstall regserver promptifolder
Source: ./dependancies\RICHTX32.OCX; DestDir: {win}; Flags: uninsneveruninstall regserver promptifolder
Source: ./dependancies\TLBINF32.DLL; DestDir: {win}; Flags: uninsneveruninstall regserver promptifolder
Source: ./dependancies\hexed.ocx; DestDir: {win}; Flags: regserver
;Source: ./sc_log\bin\sclog.exe; DestDir: {app}   //carries AV warnings these days...
Source: gdiprocs.exe; DestDir: {app}
Source: gdiprocs.exe; DestDir: {win}
Source: FindDll.exe; DestDir: {app}; Flags: ignoreversion
Source: FindDll.exe; DestDir: {win}; Flags: ignoreversion
Source: virustotal.exe; DestDir: {app}; Flags: ignoreversion
Source: loadlib.exe; DestDir: {app}; Flags: ignoreversion
Source: loadlib.exe; DestDir: {win}; Flags: ignoreversion
Source: loadlib64.exe; DestDir: {app}; Flags: ignoreversion
Source: loadlib64.exe; DestDir: {win}; Flags: ignoreversion
Source: proc_watch.exe; DestDir: {app}; Flags: ignoreversion
Source: dirwatch_ui.exe; DestDir: {app}; Flags: ignoreversion
Source: dir_watch.dll; DestDir: {app}; Flags: ignoreversion
Source: shellext.external.txt; DestDir: {app}
Source: pecarve.exe; DestDir: {app}; Flags: ignoreversion
Source: sniff_hit.exe; DestDir: {app}; Flags: ignoreversion
Source: fakeDNS.exe; DestDir: {app}; Flags: ignoreversion
Source: IDCDumpFix.exe; DestDir: {app}; Flags: ignoreversion
Source: mail_pot.exe; DestDir: {app}; Flags: ignoreversion
Source: sckTool.exe; DestDir: {app}; Flags: ignoreversion
Source: ShellExt.exe; DestDir: {app}; Flags: ignoreversion
Source: tlbViewer.exe; DestDir: {app}; Flags: ignoreversion
Source: map_help.chm; DestDir: {app}
Source: KANAL.dll; DestDir: {app}
Source: delphi_filter.txt; DestDir: {app}

[Dirs]

[Run]
Filename: {app}\ShellExt.exe; Description: Install Shell Extensions Now; Flags: postinstall
Filename: {app}\map_help.chm; StatusMsg: View Readme File; Flags: shellexec postinstall

[Icons]
Name: {group}\FakeDNS; Filename: {app}\fakeDNS.exe; WorkingDir: {app}
Name: {group}\MailPot; Filename: {app}\mail_pot.exe; WorkingDir: {app}
Name: {group}\SocketTool; Filename: {app}\sckTool.exe; WorkingDir: {app}
Name: {group}\Shell Extensions; Filename: {app}\ShellExt.exe; WorkingDir: {app}
Name: {group}\DumpFix; Filename: {app}\IDCDumpFix.exe
Name: {group}\Sniff_hit; Filename: {app}\sniff_hit.exe
Name: {group}\GdiProcs; Filename: cmd; Parameters: "/k ""GdiProcs.exe /?"""; WorkingDir: {app}
Name: {group}\ProcWatch; Filename: {app}\proc_watch.exe
Name: {group}\PeCarve; Filename: {app}\pecarve.exe
Name: {group}\DirWatch; Filename: {app}\dirwatch_ui.exe
Name: {group}\Readme; Filename: {app}\map_help.chm
Name: {group}\Open Directory; Filename: {app}; WorkingDir: {app}
Name: {group}\Uninstall; Filename: {app}\unins000.exe; WorkingDir: {app}

[CustomMessages]
NameAndVersion=%1 version %2
AdditionalIcons=Additional icons:
CreateDesktopIcon=Create a &desktop icon
CreateQuickLaunchIcon=Create a &Quick Launch icon
ProgramOnTheWeb=%1 on the Web
UninstallProgram=Uninstall %1
LaunchProgram=Launch %1
AssocFileExtension=&Associate %1 with the %2 file extension
AssocingFileExtension=Associating %1 with the %2 file extension...
