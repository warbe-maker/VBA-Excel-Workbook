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
Const HOSTED_COMM_COMPS = "mWrkbk"

Private Sub Workbook_Open()
#If CompMan Then
    mCompMan.UpdateUsedCommCompsTheOriginHasChanged wbTarget:=ThisWorkbook, sHosted:=HOSTED_COMM_COMPS
#End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
#If CompMan Then
    mCompMan.ExportChangedComponents wb:=ThisWorkbook, sHosted:=HOSTED_COMM_COMPS
#End If
End Sub