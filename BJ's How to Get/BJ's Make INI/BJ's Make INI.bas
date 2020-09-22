Attribute VB_Name = "modMakeINI"

' Two Windows API calls used to read and write private .INI files.
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


'This can also be placed in the form.
'All you have to do is remove Global from Global Const


'Info to Write to .ini file                                 Info in .ini file

Global Const SECTION = "BJ's How to Get... Make .INI"       '[BJ's How to Get... Make .INI]
Global Const Entry = "E-Mail me"                            'E-Mail me=
Global Const ENTRY1 = "Entry1"                              'Entry1=
Global Const ENTRY2 = "Entry2"                              'Entry2=
Global Const SECTION1 = "Section1"                          '[Section1]
Global Const ENTRY3 = "Entry3"                              'Entry3=
Global Const SECTION2 = "Program Information"               '[Program Information]
Global Const ENTRY4 = "INI File Path"                       'INI File Path=
Global Const ENTRY5 = "Application Path"                    'Application Path=
Global Const ENTRY6 = "Application EXE Name"                'Application EXE Name=
Global Const ENTRY7 = "Application Version"                 'Application Version=
Global Const SECTION3 = "BJ's How to Get... All Apps"       '[BJ's How to Get Apps]
Global Const ENTRY8 = "Basic Calender"                      'Basic Calender=
Global Const ENTRY9 = "Colors"                             'Colors=
Global Const ENTRY10 = "Copy Files"                         'Copy Files=
Global Const ENTRY11 = "Delete Files"                       'Delete Files=
Global Const ENTRY12 = "Disable [X]"                        'Disable [X]=
Global Const ENTRY13 = "E-Mail"                             'E-Mail=
Global Const ENTRY14 = "Image Combo"                        'Image Combo=
Global Const ENTRY15 = "Make INI"                           'Make INI=
Global Const ENTRY16 = "Multi Undo"                         'Multi Undo=
Global Const ENTRY17 = "Popup Menu"                         'Popup Menu=
Global Const ENTRY18 = "Registry Example"                   'Registry Example=
Global Const ENTRY19 = "Rename Files"                       'Rename Files=
Global Const ENTRY20 = "System About"                       'System About=
Global Const ENTRY21 = "Time Elapse"                        'Time Elapse=
Global Const ENTRY22 = "Tray Icon"                          'Tray Icon=
Global Const ENTRY23 = "Trial Version"                      'Trial Version=
Global Const ENTRY24 = "Windows Directory"                  'Windows Directory=
Global Const INI_FILE = "BJ 's How to Get... Make INI.ini"  '.ini file itself found
                                                            'in your Windows Directory
'also .ini file don't have to be .ini. It can be any extention you want it to be.
