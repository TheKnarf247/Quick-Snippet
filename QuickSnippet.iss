[Setup]
AppName=Quick Snippet
AppVersion=1.0.2
DefaultDirName={autopf}\Quick Snippet
DefaultGroupName=Quick Snippet
InfoBeforeFile=readme.txt
OutputDir=installer
OutputBaseFilename=QuickSnippetInstaller
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=quicksnippet.ico

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "dist\QuickSnippet.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Quick Snippet"; Filename: "{app}\QuickSnippet.exe"
Name: "{group}\README"; Filename: "notepad.exe"; Parameters: """{app}\README.txt"""
Name: "{group}\Uninstall Quick Snippet"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Quick Snippet"; Filename: "{app}\QuickSnippet.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\QuickSnippet.exe"; Description: "Launch Quick Snippet"; Flags: nowait postinstall skipifsilent
Filename: "notepad.exe"; Parameters: """{app}\README.txt"""; Description: "View README"; Flags: postinstall skipifsilent unchecked