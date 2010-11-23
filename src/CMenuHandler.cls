VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMenuHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' based on VBA Code Cleaner 4.4 © 1996-2006 by Rob Bovey,
''' all rights reserved. May be redistributed for free but
''' may not be sold without the author's explicit permission.
Option Explicit

Private WithEvents theSaveMenu As VBIDE.CommandBarEvents
Private WithEvents theOpenMenu As VBIDE.CommandBarEvents
Private WithEvents theDifferMenu As VBIDE.CommandBarEvents
Private WithEvents theDiffMenu As VBIDE.CommandBarEvents
Private WithEvents theCommitMenu As VBIDE.CommandBarEvents
Private WithEvents theUpdateMenu As VBIDE.CommandBarEvents
Private WithEvents theRevertMenu As VBIDE.CommandBarEvents

Private mctlToolsMenu As Office.CommandBar

Private Function existsCommandBar(barName As String) As Boolean
  Dim check As Integer

  existsCommandBar = True
  On Error GoTo err1:
  check = Application.VBE.CommandBars("Source").Index
  Exit Function
err1:
  existsCommandBar = False
End Function

Private Sub Class_Initialize()
  If existsCommandBar("Source") Then
    Set mctlToolsMenu = Application.VBE.CommandBars("Source")
    deleteMenus
  Else
    Set mctlToolsMenu = Application.VBE.CommandBars.add
    With mctlToolsMenu
      .Name = "Source"
      .Position = msoBarTop
      .Visible = True
    End With
  End If
  AddMenuItems
End Sub

Private Sub Class_Terminate()
  Set theSaveMenu = Nothing
  Set theOpenMenu = Nothing
  Set theDiffMenu = Nothing
  Set theDifferMenu = Nothing
  Set theCommitMenu = Nothing
  Set theUpdateMenu = Nothing
  Set theRevertMenu = Nothing
  deleteMenus
  Set mctlToolsMenu = Nothing
End Sub

Private Sub theSaveMenu_Click(ByVal cmdBar As Object, handled As Boolean, Cancel As Boolean)
  Application.OnTime Now(), ThisWorkbook.Name & "!SaveTo"
  handled = True
End Sub


Private Sub theOpenMenu_Click(ByVal cmdBar As Object, handled As Boolean, Cancel As Boolean)
  Application.OnTime Now(), ThisWorkbook.Name & "!LoadFrom"
  handled = True
End Sub

Private Sub theDiffMenu_Click(ByVal cmdBar As Object, handled As Boolean, Cancel As Boolean)
  Application.OnTime Now(), ThisWorkbook.Name & "!Diff"
  handled = True
End Sub

Private Sub theDifferMenu_Click(ByVal cmdBar As Object, handled As Boolean, Cancel As Boolean)
  Application.OnTime Now(), ThisWorkbook.Name & "!Differ"
  handled = True
End Sub

Private Sub theCommitMenu_Click(ByVal cmdBar As Object, handled As Boolean, Cancel As Boolean)
  Application.OnTime Now(), ThisWorkbook.Name & "!Commit"
  handled = True
End Sub

Private Sub theUpdateMenu_Click(ByVal cmdBar As Object, handled As Boolean, Cancel As Boolean)
  Application.OnTime Now(), ThisWorkbook.Name & "!Update"
  handled = True
End Sub

Private Sub theRevertMenu_Click(ByVal cmdBar As Object, handled As Boolean, Cancel As Boolean)
  Application.OnTime Now(), ThisWorkbook.Name & "!Revert"
  handled = True
End Sub

Private Sub AddMenuItems()
  Dim ctlOpen As Office.CommandBarButton
  Dim ctlSave As Office.CommandBarButton
  Dim ctlCommit As Office.CommandBarButton
  Dim ctlUpdate As Office.CommandBarButton
  Dim ctlDiff As Office.CommandBarButton
  Dim ctlDiffer As Office.CommandBarButton
  Dim ctlRevert As Office.CommandBarButton

  deleteMenus

  Set ctlRevert = mctlToolsMenu.Controls.add(msoControlButton, , , 1)
  ctlRevert.Caption = "Revert Module/Project"
  ctlRevert.FaceId = 455
  Set ctlCommit = mctlToolsMenu.Controls.add(msoControlButton, , , 1)
  ctlCommit.Caption = "Commit Module/Project"
  ctlCommit.FaceId = 270
  Set ctlUpdate = mctlToolsMenu.Controls.add(msoControlButton, , , 1)
  ctlUpdate.Caption = "Update Module/Project"
  ctlUpdate.FaceId = 271
  Set ctlDiff = mctlToolsMenu.Controls.add(msoControlButton, , , 1)
  ctlDiff.Caption = "Diff Module/Project"
  ctlDiff.FaceId = 176
  ctlDiff.BeginGroup = True

  Set ctlDiffer = mctlToolsMenu.Controls.add(msoControlButton, , , 1)
  ctlDiffer.Caption = "WinDiff Project"
  ctlDiffer.FaceId = 2042
  Set ctlSave = mctlToolsMenu.Controls.add(msoControlButton, , , 1)
  ctlSave.Caption = "Save Module/Project"
  ctlSave.FaceId = 3
  Set ctlOpen = mctlToolsMenu.Controls.add(msoControlButton, , , 1)
  ctlOpen.Caption = "Open Module/Project"
  ctlOpen.FaceId = 23

  Set theOpenMenu = Application.VBE.Events.CommandBarEvents(ctlOpen)
  Set theSaveMenu = Application.VBE.Events.CommandBarEvents(ctlSave)
  Set theDiffMenu = Application.VBE.Events.CommandBarEvents(ctlDiff)
  Set theDifferMenu = Application.VBE.Events.CommandBarEvents(ctlDiffer)
  Set theCommitMenu = Application.VBE.Events.CommandBarEvents(ctlCommit)
  Set theUpdateMenu = Application.VBE.Events.CommandBarEvents(ctlUpdate)
  Set theRevertMenu = Application.VBE.Events.CommandBarEvents(ctlRevert)
End Sub

Sub deleteMenus()
  On Error Resume Next
  mctlToolsMenu.Controls("Open Module/Project").Delete
  mctlToolsMenu.Controls("Save Module/Project").Delete
  mctlToolsMenu.Controls("Diff Module/Project").Delete
  mctlToolsMenu.Controls("WinDiff Project").Delete
  mctlToolsMenu.Controls("Commit Module/Project").Delete
  mctlToolsMenu.Controls("Update Module/Project").Delete
  mctlToolsMenu.Controls("Revert Module/Project").Delete
  On Error GoTo 0
End Sub
