VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVBEOnKey 
   Caption         =   "VBEOnKey-ID"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   OleObjectBlob   =   "frmVBEOnKey.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVBEOnKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'***************************************************************************
'*
'* PROJECT NAME:    VBEONKEY
'* AUTHOR & DATE:   STEPHEN BULLEN, Business Modelling Solutions Ltd.
'*                  10 May 2000
'*
'*                  COPYRIGHT © 2000 BY BUSINESS MODELLING SOLUTIONS LTD
'*
'* CONTACT:         Stephen@BMSLtd.co.uk
'* WEB SITE:        http://www.BMSLtd.co.uk
'*
'* DESCRIPTION:     Provides functionality similar to Application.OnKey, for the VBE
'*
'* USAGE:           To use in other projects, copy the following components into your
'*                  project, in their entirety:
'*                     - modVBEOnKey
'*                     - frmVBEOnKey
'*
'*                  You can then use the following lines to turn key trapping on and off:
'*                     VBEOnKey "%X", "RunProcedureX"
'*                     VBEOnKey "%X"
'*
'* THIS MODULE:     API functions to provide the shortcut key functionality.
'*
'* PROCEDURES:
'*
'* UserForm_Initialize  Give the form a unique caption and get it's hWnd
'* UserForm_Terminate   Ensure all hooks are destroyed if the from is destroyed
'*
'***************************************************************************
'*
'* CHANGE HISTORY
'*
'*  DATE        NAME                DESCRIPTION
'*  10/05/2000  Stephen Bullen      Initial version
'*
'***************************************************************************

Option Explicit
Option Compare Text
Option Base 1

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public mFormhWnd As Long

Private Sub UserForm_Initialize()
  Dim sCaption As String

  'Create a unique caption for our form
  Randomize
  Do
    sCaption = "VBEOnKey-" & Rnd
  Loop Until FindWindow(vbNullString, sCaption) = 0

  Me.Caption = sCaption
  mFormhWnd = FindWindow(vbNullString, sCaption)

End Sub

Private Sub UserForm_Terminate()
  UnHookAll
End Sub

