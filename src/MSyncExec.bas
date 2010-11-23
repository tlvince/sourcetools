Attribute VB_Name = "MSyncExec"
Option Explicit

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessID As Long
  dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
                                                        ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
                                                        ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
                                                        ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
                                                        lpStartupInfo As STARTUPINFO, lpProcessInformation As _
                                                                                      PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Public Sub ExecCmd(cmdline$)
  Dim Proc As PROCESS_INFORMATION
  Dim start As STARTUPINFO
  Dim ReturnValue As Integer
  ' Initialisiert die STARTUPINFO Struktur:
  start.cb = Len(start)
  ' Startet die Shell-Anwendung:
  ReturnValue = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
                               NORMAL_PRIORITY_CLASS, 0&, 0&, start, Proc)
  ' Wartet bis Shell-Anwendung geschlossen ist:
  Do
    ReturnValue = WaitForSingleObject(Proc.hProcess, 0)
    DoEvents
  Loop Until ReturnValue <> 258
  ReturnValue = CloseHandle(Proc.hProcess)
End Sub


