VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbWrkbk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' -------------------------------------------------------------------
' Standard Module mExists
'
' Requires: Reference to: "Microsoft Scripting Runtime"
'
' W. Rauschenberger, Berlin Sept 2019
' -------------------------------------------------------------------
Const HOSTED_RAWS = "mWrkbk"

Private Sub Workbook_Open()
    mCompManClient.CompManService "UpdateRawClones", HOSTED_RAWS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
End Sub

