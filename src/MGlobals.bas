Attribute VB_Name = "MGlobals"
''' based on VBA Code Cleaner 4.4 © 1996-2006 by Rob Bovey,
''' all rights reserved. May be redistributed for free but
''' may not be sold without the author's explicit permission.
Option Explicit
Option Private Module

' SVN control commands, if needed, replace with other
Public Const COMMITCMD = "C:\Programme\TortoiseSVN\bin\TortoiseProc.exe /command:commit /notempfile /path:"
Public Const UPDATECMD = "C:\Programme\TortoiseSVN\bin\TortoiseProc.exe /command:update /rev /notempfile /path:"
Public Const REVERTCMD = "C:\Programme\TortoiseSVN\bin\TortoiseProc.exe /command:revert /notempfile /path:"
Public Const DIFFCMD = "C:\Programme\TortoiseSVN\bin\TortoiseProc.exe /command:diff /path:"
Public Const PATHCONCAT = "*"
Public Const DIFFERCMD = "C:\Program Files\WinMerge\WinMergeU.exe "

' file/project specific settings
Public wbPath As String                     ' the path where source files are stored
Public cmpComponents As VBIDE.VBComponents  ' currently selected project
Public szFilename As String                 ' filename of current project (worbookname)
Public Const folderExtension = ""       ' extension to be applied to filename (which is the folder of all modules' source files)

''' Error message constants.
Public Const gszERR_VBPROJ_PROTECT As String = "Your version of Excel appears to have VB Project Protection enabled. The Code Cleaner cannot run with this setting. To solve this problem please take the following steps:" & vbLf & vbLf & _
       "1) Choose Tools > Macro > Security from the Excel menu." & vbLf & _
       "2) In the Security dialog select the Trusted Sources tab." & vbLf & _
       "3) Check the 'Trust access to Visual Basic Project' checkbox."

''' File extension constants.
Public Const gszEXT_BACKUP As String = ".bak"
Public Const gszEXT_CLASS As String = ".cls"
Public Const gszEXT_FORM As String = ".frm"
Public Const gszEXT_FORM_BINARY As String = ".frx"
Public Const gszEXT_MODULE As String = ".bas"
Public Const gszEXT_DOC As String = ".xwk"

''' Misc. constants
Public Const gdWAIT As Double = 0.00001
Public Const gszSOURCE_SAFE As String = "VSSODE"    ''' The name of the module added to Excel 2000 projects by Visual Source Safe.
Public Const gszFORCE_CALC As String = "^%{F9}"     ''' Used to do a SendKeys Ctrl+Alt+F9
Public Const gszREM As String = "Rem"
Public Const gszEMPTY_STRING As String = ""
Public Const gszBANG As String = "!"
Public Const gszSTAR As String = "*"
Public Const gszSEP As String = "\"
Public Const gszSQ As String = "'"
Public Const gszDQ As String = """"
Public Const gszCONTINUED As String = " _"
Public Const gszDOT As String = "."
Public Const gszSPACE As String = " "

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public gclsMenuHandler As CMenuHandler  ''' The VBE menu handler class.
Public glVBEHWnd As Long                ''' The VBE window handle.
Public glVersion As Long                ''' The version of Excel we're running under.
Public gszErrMsg As String              ''' Used to pass error messages between procedures.

Public Sub InitGlobals()
''' Initialize global variables
  glVBEHWnd = Application.VBE.MainWindow.hwnd
  glVersion = Val(Application.Version)
  gszErrMsg = vbNullString
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Comments:   Sets all critical Application properties. This procedure is
'''             used to isolate application-specific settings to make the
'''             code cleaner easier to port to other applications.
Public Sub SetAppProperties()
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  Application.EnableCancelKey = xlDisabled
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Comments:   Resets all critical Application properties. This procedure is
'''             used to isolate application-specific settings to make the
'''             code cleaner easier to port to other applications.
Public Sub ResetAppProperties()
  Application.ScreenUpdating = True
  Application.StatusBar = False
  Application.DisplayAlerts = True
  Application.EnableEvents = True
  Application.EnableCancelKey = xlInterrupt
End Sub

