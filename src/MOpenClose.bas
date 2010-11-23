Attribute VB_Name = "MOpenClose"
''' based on VBA Code Cleaner 4.4 © 1996-2006 by Rob Bovey,
''' all rights reserved. May be redistributed for free but
''' may not be sold without the author's explicit permission.
Option Explicit
Option Private Module

Sub Auto_Open()
''' Create the VBE menu.
  On Error Resume Next
  Set gclsMenuHandler = New CMenuHandler
  On Error GoTo 0
End Sub

Sub Auto_Close()
  Set gclsMenuHandler = Nothing
End Sub


