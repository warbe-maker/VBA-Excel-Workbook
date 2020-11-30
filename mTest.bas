Attribute VB_Name = "mTest"
Option Explicit
' ------------------------------------------------------------
' Standard Module mTest Test of all Existence checks variants
'                       in module mExists
' -----------------------------------------------------------
'Declare API
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Integer) As Integer
Const SHIFT_KEY = 16

Public Sub Regression()
' ---------------------------------------------------------
' All results are asserted and there is no intervention
' required for the whole test. When an assertion fails the
' test procedure will stop and indicates the problem with
' the called procedure.
' An execution trace is displayed for each test procedure.
'
' Please note:
' There is a lot of Workbook open and close going on during
' this test which will take some 20 seconds for the whole
' test to finish.
' ---------------------------------------------------------
    Const PROC = "Regression"

    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    mTest.Test_01_IsOpen
    mTest.Test_02_GetOpen
    mTest.Test_03_GetOpen_Errors
    mTest.Test_04_Is_
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_02_GetOpen()
    Const PROC = "Test_02_GetOpen"  ' This procedure's name for the error handling and execution tracking

    On Error GoTo eh
    Dim wb              As Workbook
    Dim wb1             As Workbook
    Dim sWb1Name        As String
    Dim sWb1FullName    As String
    Dim wb2             As Workbook
    Dim sWb2Name        As String
    Dim sWb2FullName    As String
    Dim wb3             As Workbook
    Dim sWb3Name        As String
    Dim sWb3FullName    As String
    Dim sFullName       As String
    
    mErH.BoP ErrSrc(PROC)
    
    '~~ Ensure precondition
    On Error Resume Next
    wb1.Close
    wb2.Close
    On Error GoTo eh
    sWb1Name = "Test1.xlsm"
    sWb1FullName = ThisWorkbook.Path & "\" & sWb1Name
    sWb2Name = sWb1Name
    sWb2FullName = ThisWorkbook.Path & "\Test\" & sWb2Name
    sWb3Name = "Test2.xlsm"
    sWb3FullName = ThisWorkbook.Path & "\Test\" & sWb3Name
    
    Set wb1 = Workbooks.Open(sWb1FullName) ' open the test Workbook
    
    '~~ --------------------------------------------
    '~~ Run tests (all not raising an error
    '~~ --------------------------------------------
    '~~ Test 1: GetOpen Workbook by object (open)
    Debug.Assert GetOpen(wb1) Is wb1

    '~~ Test 2: GetOpen Workbook by name (open)
    Debug.Assert GetOpen(sWb1Name) Is wb1

    '~~ Test 3: GetOpen Workbook by fullname (open)
    Debug.Assert GetOpen(sWb1FullName) Is wb1

    '~~ Test 4: GetOpen Workbook by fullname (not open)
    sFullName = wb1.FullName
    wb1.Close
    Debug.Assert GetOpen(sFullName).FullName = sFullName

    '~~ Test 5: GetOpen Workbook by full name (not open)
    '~~         A Workbook with the same name but from a different location is already open.
    On Error Resume Next
    wb1.Close
    wb2.Close
    sFullName = wb3.FullName
    wb3.Close

    On Error GoTo eh
    Debug.Assert GetOpen(sFullName).FullName = sFullName

    '~~ Test 6: GetOpen Workbook by full name (not open)
    '~~         A Workbook with the same name but from a different location is already open
    '~~         and the file does not/no longer exist at the provided location.
    Set wb3 = Workbooks.Open(sWb3FullName)
    Debug.Assert GetOpen(sWb1FullName & "\Test2.xlsm").Name = sWb3Name
    wb3.Close
    
xt: '~~ Cleanup
    On Error Resume Next
    wb1.Close
    wb2.Close
    wb3.Close
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_03_GetOpen_Errors()
    Const PROC = "Test_03_GetOpen_Errors"
    
    On Error GoTo eh
    Dim wb              As Workbook
    Dim wb1             As Workbook
    Dim sWb1Name        As String
    Dim sWb1FullName    As String
    Dim wb2             As Workbook
    Dim sWb2Name        As String
    Dim sWb2FullName    As String
    Dim wb3             As Workbook
    Dim sWb3Name        As String
    Dim sWb3FullName    As String
    Dim sFullName       As String
        
    ' Prepare
    sWb1Name = "Test1.xlsm"
    sWb1FullName = ThisWorkbook.Path & "\" & sWb1Name
    sWb2Name = sWb1Name
    sWb2FullName = ThisWorkbook.Path & "\Test\" & sWb2Name
    sWb3Name = "Test2.xlsm"
    sWb3FullName = ThisWorkbook.Path & "\Test\" & sWb3Name
    
    '~~ Test : GetOpen Workbook is object never opened
    mErH.BoTP ErrSrc(PROC), mErH.AppErr(1) ' Bypass this error as the one asserted
    mWrkbk.GetOpen wb1
    mErH.EoP ErrSrc(PROC)
    Debug.Assert mErH.MostRecentError = mErH.AppErr(1)

    '~~ Test E-2: Parameter is Nothing
    Set wb1 = Workbooks.Open(sWb1FullName) ' open the test Workbook
    wb1.Close
    mErH.BoTP ErrSrc(PROC), mErH.AppErr(2) ' Bypass this error as the one asserted
    mWrkbk.GetOpen wb1
    mErH.EoP ErrSrc(PROC)
    Debug.Assert mErH.MostRecentError = mErH.AppErr(2)

    '~~ Test E-2: Parameter is Nothing
    Set wb1 = Nothing
    
    mErH.BoTP ErrSrc(PROC), mErH.AppErr(1) ' Bypass this error as the one asserted
    mWrkbk.GetOpen wb1
    mErH.EoP ErrSrc(PROC)
    Debug.Assert mErH.MostRecentError = mErH.AppErr(1)
    
    '~~ Test E-3: Parameter is a not open Workbook's name
    On Error Resume Next
    wb1.Close
    wb2.Close
    
    mErH.BoTP ErrSrc(PROC), mErH.AppErr(5) ' Bypass this error as the one asserted
    GetOpen sWb1Name
    mErH.EoP ErrSrc(PROC)
    Debug.Assert mErH.MostRecentError = mErH.AppErr(5)

    '~~ Test E-4: Parameter is a Workbook's full name but the file does't exist
    mErH.BoTP ErrSrc(PROC), mErH.AppErr(4) ' Bypass this error as the one asserted
    mWrkbk.GetOpen Replace(sWb1FullName, sWb1Name, "not-existing.xls")
    mErH.EoP ErrSrc(PROC)
    Debug.Assert mErH.MostRecentError = mErH.AppErr(4)

    '~~ Test E-5: A Workbook with the provided name is open but from a different location
    '             and the Workbook file still exists at the provided location
    Close wb1
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & "Test3.xlsm")
    On Error Resume Next
    Set wb1 = GetOpen(ThisWorkbook.Path & "\Test\" & "Test3.xlsm")
    Debug.Assert mErH.AppErr(err.Number) = 3
    wb.Close
    
    '~~ Test E-6: Parameter is neither a Workbook object nor a string
    On Error Resume Next
    Set wb = GetOpen(ThisWorkbook.ActiveSheet)
    Debug.Assert mErH.AppErr(1)

    '~~ Cleanup
    On Error Resume Next
    wb1.Close
    wb2.Close
    wb3.Close

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_01_IsOpen()
    Const PROC = "Test_01_IsOpen"  ' This procedure's name for the error handling and execution tracking

    On Error GoTo eh
    Dim wb          As Workbook
    Dim sName       As String
    Dim o           As Object
    Dim wb1         As Workbook
    Dim wb2         As Workbook
    Dim wb3         As Workbook
    Dim wbResult    As Workbook
    
    '~~ Prepare test environment
    On Error Resume Next
    wb1.Close
    wb2.Close
    wb3.Close
    
    With Workbooks
        Set wb1 = .Open(ThisWorkbook.Path & "\Test1.xlsm")
        Set wb2 = .Open(ThisWorkbook.Path & "\Test\Test2.xlsm")
        Set wb3 = .Open(ThisWorkbook.Path & "\Test\Test3.xlsm")
    End With
    
    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    '~~ 1. Test IsOpen by object
    Debug.Assert IsOpen(wb1, wbResult) = True

    '~~ 2. Test IsOpen by Name
    Debug.Assert IsOpen(wb1.Name, wbResult) = True

    '~~ 3. Test IsOpen by FullName
    Debug.Assert IsOpen(wb1.FullName, wbResult) = True

    '~~ 4. A Workbook with the givven nmae is open but from a different location
    '~~    Since the Workbook does not or no longer exist at the requested location it regarded moved and comsidered open
    Debug.Assert IsOpen(wb1.Path & "\Test2.xlsm", wbResult) = True
    Debug.Assert wbResult.FullName = wb1.Path & "\Test\Test2.xlsm"
    
    '~~ 4b No Workbook object is returned since the parameter is not Variant    Debug.Assert vWb Is wb2
    Debug.Assert IsOpen(wb1.Path & "\Test2.xlsm", wbResult) = True
    
    '~~ 5. Workbook does not exist. When a fullname is provided an error is raised
    Debug.Assert IsOpen(wb1.Path & "\Test\Test.xlsm", wbResult) = False
        
    '~~ 6. A Workbook with the given Name is open but from a different location
    '~~    Since it still exists at the requested location it is regarde not open
    wb3.Close
    Debug.Assert IsOpen(wb1.Path & "\Test3.xlsm", wbResult) = False
    
xt: wb1.Close
    wb2.Close
    On Error Resume Next: wb3.Close
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_04_Is_()
    Const PROC = "Test_04_Is_"
    
    On Error GoTo eh
    Dim wb  As Workbook
    Dim wb1 As Workbook

    Set wb = mWrkbk.GetOpen(ThisWorkbook.Path & "\" & "Test1.xlsm")
    
    mErH.BoP ErrSrc(PROC)
    Debug.Assert IsName(wb.Name) = True
    Debug.Assert IsName(wb.FullName) = False
    Debug.Assert IsName(wb.Path) = False
    Debug.Assert IsName(ThisWorkbook) = False
    
    Debug.Assert IsFullName(wb.Name) = False
    Debug.Assert IsFullName(wb.FullName) = True
    Debug.Assert IsFullName(wb.Path) = False
    Debug.Assert IsFullName(ThisWorkbook) = False

    Debug.Assert IsObject(wb.Name) = False
    Debug.Assert IsObject(wb.FullName) = False
    Debug.Assert IsObject(wb.Path) = False
    Debug.Assert IsObject(ThisWorkbook) = True
    
    Debug.Assert IsObject(wb) = True
    wb.Close
    Debug.Assert IsObject(wb) = True                              ' A closed Workbook is still a Workbook object
    Set wb = Nothing
    Debug.Assert IsObject(wb) = False                             ' A set to Nothing is no longer a Workbook object
    Debug.Assert IsObject(wb1) = False                             ' A never assigned Workbook is not a Workbook object
        
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub
 
Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest" & "." & sProc
End Function
