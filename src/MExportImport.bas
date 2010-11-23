Attribute VB_Name = "MExportImport"
'Copyright (C) 2007 Roland Kapl
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

''' based on VBA Code Cleaner 4.4 © 1996-2006 by Rob Bovey,
''' all rights reserved. May be redistributed for free but
''' may not be sold without the author's explicit permission.
Option Explicit

' setwbPathAndName
'
' sets global path of project and (if module is selected), the filename of the module (ext. see globals)
' sideeffect: sets componentSelected, depending whether module is selected (for context dependent commandbars)
Function setwbPathAndName(wbproj As VBProject, ByRef componentSelected As Boolean) As Boolean
  componentSelected = True
  On Error Resume Next
  szFilename = wbproj.Parent.SelectedVBComponent.Name
  If Err <> 0 Then
    componentSelected = False
  Else
    setName wbproj.Parent.SelectedVBComponent
  End If
  Err.Clear
  wbPath = Left(wbproj.FileName, InStrRev(wbproj.FileName, ".") - 1)
  If Err <> 0 Then
    MsgBox "You haven't saved the workbook yet"
    Exit Function
  End If
  
  Set cmpComponents = wbproj.VBComponents

  If Err <> 0 Then
    MsgBox "couldn't determine path of project: " & Err.Description, vbCritical
    setwbPathAndName = False
  End If
  setwbPathAndName = True
End Function

Function setName(cmpComponent As VBIDE.VBComponent) As Boolean
  On Error GoTo setName_Error
  setName = True
  szFilename = cmpComponent.Name
  ''' Concatenate the correct filename for export.
  Select Case cmpComponent.Type
  Case vbext_ct_ClassModule
    szFilename = szFilename & gszEXT_CLASS
  Case vbext_ct_MSForm
    szFilename = szFilename & gszEXT_FORM
  Case vbext_ct_StdModule
    szFilename = szFilename & gszEXT_MODULE
  Case vbext_ct_Document
    szFilename = szFilename & gszEXT_DOC
  End Select
  Exit Function
setName_Error:
  setName = False
  Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in setName"
End Function

' cmdbutton diff modules/project
Sub Differ()
  Dim componentSelected As Boolean
  Dim origAddInName, addInName, wbOrigPath, wbComparePath, pathAndFilename, onlyFileName As String
  Dim compareWB As Workbook

  If Not setwbPathAndName(ThisWorkbook.VBProject.VBE.ActiveVBProject, componentSelected) Then Exit Sub
  Save
  wbOrigPath = wbPath
Selection:
  pathAndFilename = Application.GetOpenFilename( _
                    "All files (*.*), *.*", 1, _
                    "Choose Excel files to compare ...", MultiSelect:=False)
  If VarType(pathAndFilename) = vbBoolean Then
    If MsgBox("You haven't selected any file." & vbLf & "Cancel comparison ?", vbYesNo, "No file selected...") = vbNo Then
      GoTo Selection
    Else
      Exit Sub
    End If
  End If
  ' load addin or file...
  onlyFileName = Mid(pathAndFilename, InStrRev(pathAndFilename, "\") + 1)
  wbComparePath = Environ("TEMP") & "\" & "Compare_" & onlyFileName
  origAddInName = Replace(onlyFileName, ".xla", "")
  'Workbooks(onlyFileName).Close False
  FileCopy pathAndFilename, wbComparePath
  If Right(pathAndFilename, 3) = "xla" Then
    Application.AddIns.add wbComparePath, False
    addInName = Replace("Compare_" & onlyFileName, ".xla", "")
    Application.AddIns(addInName).Installed = True
  Else
    Set compareWB = Workbooks.Open(wbComparePath, False, , , , , True, , , True, True)
  End If
  ' save source
  setwbPathAndName ThisWorkbook.VBProject.VBE.ActiveVBProject, componentSelected
  Save
  ' close again
  If compareWB Is Nothing Then
    Application.AddIns(addInName).Installed = False
  Else
    compareWB.Close False
  End If
  ' invoke Differ
  ExecCmd DIFFERCMD & """" & wbPath & folderExtension & """ """ & wbOrigPath & folderExtension & """"
  ' clean up afterwards
  Kill wbComparePath
  Kill wbComparePath & folderExtension & "\*.*"
  RmDir wbComparePath & folderExtension
End Sub

' cmdbutton diff modules/project
Sub Diff()
  Dim componentSelected As Boolean

  If Not setwbPathAndName(ThisWorkbook.VBProject.VBE.ActiveVBProject, componentSelected) Then Exit Sub
  If componentSelected Then
    Save ThisWorkbook.VBProject.VBE.ActiveVBProject.Parent.SelectedVBComponent
    Shell DIFFCMD & """" & wbPath & folderExtension & gszSEP & szFilename & """", vbNormalFocus
  Else
    Save
    Shell DIFFCMD & """" & wbPath & folderExtension & """", vbNormalFocus
  End If
End Sub

' cmdbutton commit modules/project
Sub Commit()
  Dim componentSelected As Boolean
  Dim curFileName, workbookSavePath As String

  If Not setwbPathAndName(ThisWorkbook.VBProject.VBE.ActiveVBProject, componentSelected) Then Exit Sub
  If componentSelected Then
    Save ThisWorkbook.VBProject.VBE.ActiveVBProject.Parent.SelectedVBComponent
    Shell COMMITCMD & """" & wbPath & folderExtension & gszSEP & szFilename & """", vbNormalFocus
  Else
    curFileName = ThisWorkbook.VBProject.VBE.ActiveVBProject.FileName
    curFileName = Mid(curFileName, InStrRev(curFileName, gszSEP) + 1)
    ' also commit whole file...
    Workbooks(curFileName).Save
    workbookSavePath = Workbooks(curFileName).Path
    ' this is a workaround for a bug in tortoiseProc crashing/behaving abnormally when giving
    ' network paths with one being UNC and the other having a drive letter ...
    If IsPathNetPath(workbookSavePath) Then _
       workbookSavePath = GetUncFullPathFromMappedDrive(workbookSavePath)
    Save
    Shell COMMITCMD & """" & wbPath & folderExtension & PATHCONCAT & workbookSavePath & gszSEP & curFileName & """", vbNormalFocus
  End If
End Sub

' cmdbutton revert modules/project
Sub Revert()
  Dim componentSelected As Boolean
  Dim curComponent As VBIDE.VBComponent

  If Not setwbPathAndName(ThisWorkbook.VBProject.VBE.ActiveVBProject, componentSelected) Then Exit Sub
  If componentSelected Then
    Set curComponent = ThisWorkbook.VBProject.VBE.ActiveVBProject.Parent.SelectedVBComponent
    ExecCmd REVERTCMD & """" & wbPath & folderExtension & gszSEP & szFilename & """"
    Load curComponent
  Else
    ExecCmd REVERTCMD & """" & wbPath & folderExtension & """"
    Load
  End If
End Sub

' cmdbutton update modules/project
Sub Update()
  Dim componentSelected As Boolean
  Dim curComponent As VBIDE.VBComponent

  If Not setwbPathAndName(ThisWorkbook.VBProject.VBE.ActiveVBProject, componentSelected) Then Exit Sub
  If componentSelected Then
    Set curComponent = ThisWorkbook.VBProject.VBE.ActiveVBProject.Parent.SelectedVBComponent
    Save curComponent
    ExecCmd UPDATECMD & """" & wbPath & folderExtension & gszSEP & szFilename & """"
    Load curComponent
    curComponent.Activate
  Else
    Save
    ExecCmd UPDATECMD & """" & wbPath & folderExtension & """"
    Load
  End If
End Sub

' cmdbutton load modules/project
Sub LoadFrom()
  Dim componentSelected As Boolean

  If Not setwbPathAndName(ThisWorkbook.VBProject.VBE.ActiveVBProject, componentSelected) Then Exit Sub
Selection:
  wbPath = displayFolderOpen("Select folder to load module sources from...", wbPath)
  If wbPath = "" Then Exit Sub

  If componentSelected Then
    Load ThisWorkbook.VBProject.VBE.ActiveVBProject.Parent.SelectedVBComponent
  Else
    Load
  End If
  ChDir wbPath
End Sub

' cmdbutton save modules/project
Sub SaveTo()
  Dim componentSelected As Boolean

  If Not setwbPathAndName(ThisWorkbook.VBProject.VBE.ActiveVBProject, componentSelected) Then Exit Sub
  If componentSelected Then
    Save ThisWorkbook.VBProject.VBE.ActiveVBProject.Parent.SelectedVBComponent
  Else
    Save
  End If
End Sub

' Load
'
' Load actually loads the component / or all components into Excel VBA
Public Sub Load(Optional cmpComponent As VBIDE.VBComponent = Nothing)
  Dim lReturn As Long

  ' save ourselves from restoring ...
  If InStr(1, wbPath, "SourceTools.xla") > 0 Then Exit Sub

  On Error GoTo ErrorHandler
  If Not bIsProjectProtected() Then
    LockWindowUpdate (glVBEHWnd)
    SetAppProperties
    If cmpComponent Is Nothing Then
      ImportFiles
    Else
      ImportFilesSingle cmpComponent
    End If
    ResetAppProperties
    Application.Visible = True
    lReturn = LockWindowUpdate(0&)
  Else
    MsgBox gszERR_VBPROJ_PROTECT, vbExclamation
  End If
  Exit Sub
ErrorHandler:
  Debug.Print Err.Description
End Sub

' Save
'
' Save actually saves the component / or all components to disk
Public Sub Save(Optional cmpComponent As VBIDE.VBComponent = Nothing)
  Dim bCheckNames As Boolean

  If Not bIsProjectProtected() Then
    SetAppProperties
    On Error Resume Next
    ' for excel 5/95 workbooks convert module names to valid chars...
    bCheckNames = CBool(ThisWorkbook.VBProject.VBE.ActiveVBProject.VBComponents("ThisWorkbook").Properties("Modules").Value("Count"))
    If Err <> 0 Then
      bCheckNames = CBool(ThisWorkbook.VBProject.VBE.ActiveVBProject.VBComponents("DieseArbeitsmappe").Properties("Modules").Value("Count"))
    End If
    On Error GoTo ErrorHandler
    If cmpComponent Is Nothing Then
      ExportFiles bCheckNames
    Else
      ExportFiles bCheckNames, cmpComponent
    End If
    ResetAppProperties
  Else
    MsgBox gszERR_VBPROJ_PROTECT, vbExclamation
  End If
  Exit Sub
ErrorHandler:
  Debug.Print Err.Description
End Sub

' ExportFiles
'
' This procedure exports select/all VBComponents to text files
Public Sub ExportFiles(ByVal bCheckNames As Boolean, Optional cmpComponent As VBIDE.VBComponent = Nothing)
' check if we already have a folder for storing
  If Dir(wbPath & folderExtension, vbDirectory) = "" Then MkDir wbPath & folderExtension

  If cmpComponent Is Nothing Then
    ''' Loop the collection of VBComponents
    For Each cmpComponent In cmpComponents
      DoEvents
      exportComponent bCheckNames, cmpComponent
    Next cmpComponent
  Else
    ' single module export
    DoEvents
    exportComponent bCheckNames, cmpComponent
  End If
  Exit Sub

ErrorHandler:
  Debug.Print Err.Description
End Sub

' exportComponent
'
' This procedure exports one Component to a text file (help proc for ExportFiles)
Sub exportComponent(ByVal bCheckNames As Boolean, cmpComponent As VBIDE.VBComponent)
  On Error GoTo ErrorHandler
  szFilename = cmpComponent.Name
  If bCheckNames Then
    If cmpComponent.Type = vbext_ct_StdModule Then
      ProcessName szFilename
      cmpComponent.Name = szFilename
    End If
  End If
  ''' Concatenate the correct filename for export.
  Select Case cmpComponent.Type
  Case vbext_ct_ClassModule
    szFilename = szFilename & gszEXT_CLASS
  Case vbext_ct_MSForm
    szFilename = szFilename & gszEXT_FORM
  Case vbext_ct_StdModule
    szFilename = szFilename & gszEXT_MODULE
  Case vbext_ct_Document
    szFilename = szFilename & gszEXT_DOC
  End Select
  cmpComponent.Export wbPath & folderExtension & gszSEP & szFilename
  If cmpComponent.Type <> vbext_ct_Document Then _
     stripAttributeLines wbPath & folderExtension & gszSEP & szFilename, Trim(cmpComponent.CodeModule.Lines(1, 1))
  Exit Sub
ErrorHandler:
  Debug.Print Err.Description
End Sub

' ImportFiles
'
' This procedure removes and imports previously exported VBComponents back into the project.
Public Sub ImportFiles()
  Dim cmpComponent As VBIDE.VBComponent

  For Each cmpComponent In cmpComponents
    DoEvents
    Select Case cmpComponent.Type
    Case vbext_ct_ClassModule
      cmpComponents.Remove cmpComponent
    Case vbext_ct_MSForm
      cmpComponents.Remove cmpComponent
    Case vbext_ct_StdModule
      cmpComponents.Remove cmpComponent
    Case vbext_ct_Document
      ''' This is a worksheet or workbook object. Don't remove
    End Select
  Next
  On Error GoTo ErrorHandler
  ResetAppProperties
  Application.OnTime Now(), ThisWorkbook.Name & "!ImportFiles2"
  Exit Sub

ErrorHandler:
  Debug.Print Err.Description
End Sub

Public Sub ImportFiles2()
  Dim szFullName As String
  Dim szComponentName As String

  On Error GoTo ErrorHandler
  SetAppProperties
  szComponentName = Dir(wbPath & folderExtension & gszSEP & gszSTAR)
  If szComponentName <> "" Then
    ''' Loop each item on the exported files
    Do While szComponentName <> ""
      ' dont include workbook file itself, frx (form binary) and document code files !!
      If szComponentName <> Mid(wbPath, InStrRev(wbPath & folderExtension, gszSEP)) And _
         Right(szComponentName, 4) = gszEXT_CLASS Or _
         Right(szComponentName, 4) = gszEXT_FORM Or _
         Right(szComponentName, 4) = gszEXT_MODULE Then
        ''' Concatenate the path and name of the file to import.
        szFullName = wbPath & folderExtension & gszSEP & szComponentName
        ''' Import the file to the VBComponents collection.
        cmpComponents.Import szFullName
      End If
      szComponentName = Dir
    Loop
  End If
  RemoveInitialUserFormBlanks cmpComponents
  Exit Sub

ErrorHandler:
  Debug.Print Err.Description
End Sub


Public Sub ImportFilesSingle(cmpComponent As VBIDE.VBComponent)
  DoEvents
  If Not setName(cmpComponent) Then Exit Sub
  Select Case cmpComponent.Type
  Case vbext_ct_ClassModule
    cmpComponents.Remove cmpComponent
  Case vbext_ct_MSForm
    cmpComponents.Remove cmpComponent
  Case vbext_ct_StdModule
    cmpComponents.Remove cmpComponent
  Case vbext_ct_Document
    ''' This is a worksheet or workbook object. Don't remove
  End Select
  On Error GoTo ErrorHandler
  ResetAppProperties
  Application.OnTime Now(), ThisWorkbook.Name & "!ImportFilesSingle2"
  Exit Sub

ErrorHandler:
  Debug.Print Err.Description
End Sub

Public Sub ImportFilesSingle2()
  Dim szFullName As String

  On Error GoTo ErrorHandler
  SetAppProperties
  If Right(szFilename, 4) = gszEXT_CLASS Or _
     Right(szFilename, 4) = gszEXT_FORM Or _
     Right(szFilename, 4) = gszEXT_MODULE Then
    ''' Concatenate the path and name of the file to import.
    szFullName = wbPath & folderExtension & gszSEP & szFilename
    ''' Import the file to the VBComponents collection.
    cmpComponents.Import szFullName
  End If
  RemoveInitialUserFormBlanks cmpComponents
  Exit Sub

ErrorHandler:
  Debug.Print Err.Description
End Sub


' RemoveInitialUserFormBlanks
'
' The importing process causes extra blank lines to be inserted
' at the top of UserForm code modules. This function removes those blank lines.
'
' Arguments:  cmpComponents   [in] The VBComponents collection of the project being cleaned.
Public Sub RemoveInitialUserFormBlanks(ByRef cmpComponents As VBIDE.VBComponents)

  Dim szInputLine As String
  Dim modCode As VBIDE.CodeModule
  Dim cmpComponent As VBIDE.VBComponent

  For Each cmpComponent In cmpComponents
    If cmpComponent.Type = vbext_ct_MSForm Then

      Set modCode = cmpComponent.CodeModule
      szInputLine = Trim$(modCode.Lines(1, 1))

      Do While Len(szInputLine) = 0 And modCode.CountOfLines > 0
        modCode.DeleteLines 1
        szInputLine = Trim$(modCode.Lines(1, 1))
      Loop
    End If
  Next cmpComponent
  Exit Sub

ErrorHandler:
  Debug.Print Err.Description
End Sub

' bIsProjectProtected
'
' Determines if VBProject protection is enabled in versions of Excel higher than 2000.
Public Function bIsProjectProtected() As Boolean
  Dim objProject As VBIDE.VBProject
  On Error Resume Next
  If Val(Application.Version) >= 10 Then
    Set objProject = ThisWorkbook.VBProject
    bIsProjectProtected = (objProject Is Nothing)
  End If
End Function


' ProcessName
'
' Takes the name of an Excel 5/95 object and removes any
' characters that are illegal in Excel 97/2000 VBA object names.
'
' Arguments:  szOldName   [in] The name of the Excel 5/95 object to process.
Public Sub ProcessName(ByRef szOldName As String)
  Dim lCount As Long
  Dim szChar As String
  Dim szNewName As String
  szNewName = vbNullString
  For lCount = 1 To Len(szOldName)
    szChar = Mid$(szOldName, lCount, 1)
    Select Case Asc(szChar)
      ''' Upper case letters, lower case letters and underscore - all OK.
    Case 65 To 90, 95, 97 To 122
      szNewName = szNewName & szChar
      ''' Numbers - OK as long as they aren't first character.
    Case 48 To 57
      If Len(szNewName) > 0 Then szNewName = szNewName & szChar
    Case Else
      ''' No other characters are valid.
    End Select
  Next lCount
  ''' Pass back the processed name.
  szOldName = szNewName
End Sub

' stripAttributeLines
'
' strips unwanted Attribute VB_.. lines from the exported modules, which make diffing/versioning
' much harder than necessary...
Public Sub stripAttributeLines(FileName As String, firstCodeLine As String)
  Dim currentLine As String

  On Error GoTo err1:
  Open FileName For Input As 1
  Open FileName & ".tmp" For Output As 2
  ' first retain all Attributes before official beginning of file:
  Do
    Line Input #1, currentLine
    Print #2, currentLine
  Loop Until currentLine = firstCodeLine Or EOF(1)
  ' then skip all lines beginning with "Attribute ":
  If Not EOF(1) Then
    Do
      Line Input #1, currentLine
      If Left(currentLine, 10) <> "Attribute " Then _
         Print #2, currentLine
    Loop Until EOF(1)
  End If
  ' clean up
  Close 1: Close 2
  On Error Resume Next
  Kill FileName
  Name FileName & ".tmp" As FileName
  Exit Sub
err1:
  Close 1: Close 2
End Sub


