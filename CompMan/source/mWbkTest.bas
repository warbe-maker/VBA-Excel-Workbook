Attribute VB_Name = "mWbkTest"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mWbkTest: Common Excel Workbook Services - Test
' =========================
'
' Uses (for this test only): mBasic, mErH, fMsg/mMsg, mTrc
'
' W. Rauschenberger Berlin, Jun 2023
' See: https://github.com/warbe-maker/VBA-Excel-Workbook
' ------------------------------------------------------------------------------
'Declare API
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Integer) As Integer
Const SHIFT_KEY = 16

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function
  
Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mWbkTest" & "." & sProc
End Function

Public Sub Regression()
' ----------------------------------------------------------------------------
' All results are asserted and there is no intervention required for the whole
' test. When an assertion fails the test procedure will stop and indicates the
' problem with the called procedure. An execution trace log is displayed at
' the end of the test.
'
' Please note: There is a lot of Workbook open and close going on during this
'              test. This will take some 20 seconds for the whole test to
'              finish and comes with some unavoidable flicker.
' ----------------------------------------------------------------------------
    Const PROC = "Regression"

    On Error GoTo eh
    
    '~~ Initialization (must be done prior the first BoP !)
    mTrc.FileName = "RegressionTest.ExecTrace.log"
    mTrc.Title = "Regression Test mWbk module"
    mTrc.NewFile
    mErH.Regression = True
    
    mBasic.BoP ErrSrc(PROC)
    mWbkTest.Test_00_Opened_Service
    mWbkTest.Test_01_IsOpen_Service
    mWbkTest.Test_02_GetOpen_Service
    mWbkTest.Test_03_GetOpen_Service_Error_Conditions
    mWbkTest.Test_04_Is_Services
    mWbkTest.Test_05_Value_Service
    mWbkTest.Test_06_Exists_Service
    
xt: mBasic.EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_01_IsOpen_Service()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_01_IsOpen_Service"  ' This procedure's name for the error handling and execution tracking

    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim wbk              As Workbook
    Dim sName           As String
    Dim o               As Object
    Dim wb1             As Workbook
    Dim wb2             As Workbook
    Dim wb3             As Workbook
    Dim wbResult        As Workbook
    Dim sWb1FullName    As String
    Dim sWb2FullName    As String
    Dim sWb3FullName    As String
    Dim sWb1Name        As String
    Dim sWb2Name        As String
    Dim sWb3Name        As String
    
    mBasic.BoP ErrSrc(PROC)
    sWb1FullName = ThisWorkbook.Path & "\Test\Test1.xlsm"
    sWb2FullName = ThisWorkbook.Path & "\Test\TestSubFolder\Test2.xlsm"
    sWb3FullName = ThisWorkbook.Path & "\Test\TestSubFolder\Test3.xlsm"
    sWb1Name = fso.GetFileName(sWb1FullName)
    sWb2Name = fso.GetFileName(sWb2FullName)
    sWb3Name = fso.GetFileName(sWb3FullName)
    
    '~~ Prepare test environment
    mWbk.WbClose sWb1Name
    mWbk.WbClose sWb2Name
    mWbk.WbClose sWb3Name
    
    With Workbooks
        Set wb1 = .Open(sWb1FullName)
        Set wb2 = .Open(sWb2FullName)
        Set wb3 = .Open(sWb3FullName)
    End With
    
    '~~ 1. Test IsOpen by object
    Debug.Assert mWbk.IsOpen(wb1, wbResult) = True

    '~~ 2. Test IsOpen by Name
    Debug.Assert mWbk.IsOpen(wb1.Name, wbResult) = True

    '~~ 3. Test IsOpen by FullName
    Debug.Assert mWbk.IsOpen(wb1.FullName, wbResult) = True

    '~~ 4. A Workbook with the given name is open but from a different location
    '~~    Since the Workbook does not or no longer exist at the requested location it regarded moved and considered open
    Debug.Assert mWbk.IsOpen(wb1.Path & "\Test2.xlsm", wbResult) = True
    Debug.Assert wbResult.FullName = wb1.Path & "\TestSubFolder\Test2.xlsm"
    
    '~~ 4b No Workbook object is returned since the parameter is not Variant    Debug.Assert vWb Is wb2
    Debug.Assert mWbk.IsOpen(wb1.Path & "\Test2.xlsm", wbResult) = True
    
    '~~ 5. Workbook does not exist. When a fullname is provided an error is raised
    Debug.Assert mWbk.IsOpen(wb1.Path & "\Test\Test.xlsm", wbResult) = False
        
    '~~ 6. A Workbook with the given Name is open but from a different location
    '~~    Since it still exists at the requested location it is regarde not open
    wb3.Close
    Debug.Assert mWbk.IsOpen(wb1.Path & "\Test3.xlsm", wbResult) = False
    
xt: On Error Resume Next
    With Application
        .Workbooks(sWb1Name).Close
        .Workbooks(sWb2Name).Close
        .Workbooks(sWb3Name).Close
    End With
    Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_02_GetOpen_Service()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_02_GetOpen_Service"  ' This procedure's name for the error handling and execution tracking

    On Error GoTo eh
    Dim wbk              As Workbook
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
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Ensure precondition
    On Error Resume Next
    wb1.Close False
    wb2.Close False
    wb3.Close False
    
    On Error GoTo eh
    sWb1Name = "Test1.xlsm"
    sWb1FullName = ThisWorkbook.Path & "\Test\" & sWb1Name
    sWb2Name = sWb1Name
    sWb2FullName = ThisWorkbook.Path & "\Test\TestSubFolder\" & sWb2Name
    sWb3Name = "Test2.xlsm"
    sWb3FullName = ThisWorkbook.Path & "\Test\TestSubFolder\" & sWb3Name
    
    Set wb1 = Workbooks.Open(sWb1FullName) ' open the first test Workbook
    
    '~~ --------------------------------------------
    '~~ Run tests (all not raising an error)
    '~~ --------------------------------------------
    '~~ Test 1: GetOpen Workbook by object (open)
    mBasic.BoP ErrSrc(PROC) & ".Test_ArgIsWorkbookObject"
    Debug.Assert GetOpen(wb1) Is wb1
    mBasic.EoP ErrSrc(PROC) & ".Test_ArgIsWorkbookObject"

    '~~ Test 2: GetOpen Workbook by name (open)
    mBasic.BoP ErrSrc(PROC) & ".Test_ArgIsNameOfOpenWorkbook"
    Debug.Assert GetOpen(sWb1Name) Is wb1
    mBasic.EoP ErrSrc(PROC) & ".Test_ArgIsNameOfOpenWorkbook"

    '~~ Test 3: GetOpen Workbook by fullname (open)
    mBasic.BoP ErrSrc(PROC) & ".Test_ArgIsWorkbookFullNameOpen"
    Debug.Assert GetOpen(sWb1FullName) Is wb1
    mBasic.EoP ErrSrc(PROC) & ".Test_ArgIsWorkbookFullNameOpen"

    '~~ Test 4: GetOpen Workbook by fullname (not open)
    mBasic.BoP ErrSrc(PROC) & ".Test_ArgIsWorkbookFullNameNotOpen"
    sFullName = wb1.FullName
    wb1.Close False
    Set wb1 = GetOpen(sFullName)
    Debug.Assert wb1.FullName = sFullName
    wb1.Close False
    mBasic.EoP ErrSrc(PROC) & ".Test_ArgIsWorkbookFullNameNotOpen"
    
    '~~ Test 5: GetOpen Workbook by full name
    '~~         A Workbook with the same name but from a different location is already open
    '~~         and the file does not/no longer exist at the provided location.
    mBasic.BoP ErrSrc(PROC) & ".Test_ArgIsWorkbookFullNameOpenFromMovedToLocation"
    Set wb3 = Workbooks.Open(sWb3FullName)
    Debug.Assert GetOpen(Replace(sWb1FullName, "Test1", "Test2")).Name = sWb3Name
    wb3.Close False
    mBasic.EoP ErrSrc(PROC) & ".Test_ArgIsWorkbookFullNameOpenFromMovedToLocation"
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_03_GetOpen_Service_Error_Conditions()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_03_GetOpen_Service_Error_Conditions"
    
    On Error GoTo eh
    Dim wbk              As Workbook
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
        
    mBasic.BoP ErrSrc(PROC)
    ' Prepare
    sWb1Name = "Test1.xlsm"
    sWb1FullName = ThisWorkbook.Path & "\" & sWb1Name
    sWb2Name = sWb1Name
    sWb2FullName = ThisWorkbook.Path & "\Test\" & sWb2Name
    sWb3Name = "Test2.xlsm"
    sWb3FullName = ThisWorkbook.Path & "\Test\" & sWb3Name
    
    '~~ Test 1a: GetOpen Workbook is object never opened
    mBasic.BoP ErrSrc(PROC) & ".Test_ObjectIsNothing"
    mErH.Asserted AppErr(1) ' skip display of error message when mBasic.Regression = True
    mWbk.GetOpen wb1
    mBasic.EoP ErrSrc(PROC) & ".Test_ObjectIsNothing"
    
    '~~ Test 1b: Parameter is neither a Workbook object nor a string
    mBasic.BoP ErrSrc(PROC) & ".Test_NeitherWorkbookObjectNorString"
    mErH.Asserted AppErr(1)
    Set wbk = GetOpen(ThisWorkbook.ActiveSheet)
    mBasic.EoP ErrSrc(PROC) & ".Test_NeitherWorkbookObjectNorString"
    
    '~~ Test 2: A Workbook with the provided name is open but from a different location
    '             and the Workbook file still exists at the provided location
    mBasic.BoP ErrSrc(PROC) & ".Test_OpenButDiffLocationNoLongerExisting"
    If Not wb1 Is Nothing Then wb1.Close False
    Set wbk = Workbooks.Open(ThisWorkbook.Path & "\Test\TestSubFolder\Test3.xlsm")
    mErH.Asserted AppErr(2)
    Set wb1 = GetOpen(ThisWorkbook.Path & "\Test\" & "Test3.xlsm")
    wb1.Close False
    mBasic.EoP ErrSrc(PROC) & ".Test_OpenButDiffLocationNoLongerExisting"
    
    '~~ Test 3: Parameter is a not open Workbook's name
    mBasic.BoP ErrSrc(PROC) & ".Test_WorkbookNameNotAnOpenWorkbook"
    mErH.Asserted AppErr(3) ' skip display of error message when mBasic.Regression = True
    GetOpen sWb1Name
    mBasic.EoP ErrSrc(PROC) & ".Test_WorkbookNameNotAnOpenWorkbook"

    '~~ Test 4: Parameter is a Workbook's full name but the file does't exist
    mBasic.BoP ErrSrc(PROC) & ".Test_WorkbookFullNameNotExisting"
    mErH.Asserted AppErr(4) ' skip display of error message when mBasic.Regression = True
    mWbk.GetOpen Replace(sWb1FullName, sWb1Name, "not-existing.xls")
    mBasic.EoP ErrSrc(PROC) & ".Test_WorkbookFullNameNotExisting"
    
    '~~ Cleanup
    On Error Resume Next
    If Not wb1 Is Nothing Then wb1.Close False
    If Not wb2 Is Nothing Then wb2.Close False
    If Not wb3 Is Nothing Then wb3.Close False

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_04_Is_Services()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_04_Is_Services"
    
    On Error GoTo eh
    Dim wbk              As Workbook
    Dim wb1             As Workbook
    Dim fso             As New FileSystemObject
    Dim sWb1FullName    As String
    Dim sWb1Name        As String
    
    mBasic.BoP ErrSrc(PROC)
    
    sWb1FullName = ThisWorkbook.Path & "\Test\Test1.xlsm"
    sWb1Name = fso.GetFileName(sWb1FullName)
    Set wbk = mWbk.GetOpen(sWb1FullName)
    
    '~~ Test 1: IsName
    Debug.Assert IsName(wbk.Name) = True
    Debug.Assert IsName(wbk.FullName) = False
    Debug.Assert IsName(wbk.Path) = False
    Debug.Assert IsName(ThisWorkbook) = False
    Debug.Assert IsName(fso.GetBaseName(wbk.FullName)) = False
    
    '~~ Test 2: IsFullName
    Debug.Assert mWbk.IsFullName(wbk.Name) = False
    Debug.Assert mWbk.IsFullName(wbk.FullName) = True
    Debug.Assert mWbk.IsFullName(wbk.Path) = False
    Debug.Assert mWbk.IsFullName(ThisWorkbook) = False

    '~~ Test 3: IsWbObject
    Debug.Assert mWbk.IsWbObject(wbk.Name) = False
    Debug.Assert mWbk.IsWbObject(wbk.FullName) = False
    Debug.Assert mWbk.IsWbObject(wbk.Path) = False
    Debug.Assert mWbk.IsWbObject(ThisWorkbook) = True
    Debug.Assert mWbk.IsWbObject(wbk) = True
    wbk.Close
    Debug.Assert mWbk.IsWbObject(wbk) = False               ' A closed Workbook is still an object but not an object Type Workbook
    Set wbk = Nothing
    Debug.Assert mWbk.IsWbObject(wbk) = False               ' A set to Nothing is no longer a Workbook object
    Debug.Assert mWbk.IsWbObject(wb1) = False              ' A never assigned Workbook is not a Workbook object
        
xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_00_Opened_Service()
' ----------------------------------------------------------------------------
' This test works with the current Excel environment. I.e. it is insecure
' how many and which Workbooks are open. Certain only is ThisWorkbook.
' Because the function is used with the IsOpen and the GetOpen service the
' Regression test will tested it first.
' ----------------------------------------------------------------------------
    Const PROC = "Test_00_Opened_Service"
    
    On Error GoTo eh
    Dim dct As Dictionary
    Dim v   As Variant
    
    mBasic.BoP ErrSrc(PROC)
    
    Set dct = mWbk.Opened
    Debug.Assert Opened.Count >= 1
    
    For Each v In dct
        If v = ThisWorkbook.Name Then
            Debug.Assert dct(v) Is ThisWorkbook
            Exit For
        End If
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_05_Value_Service()
    Const PROC = "Test_05_Value_Service"
    
    On Error GoTo eh
    Dim Rng As Range
    
    mBasic.BoP ErrSrc(PROC)

    '~~ Test 1: Range name as string
    wshWbkTest.Unprotect
    wshWbkTest.UsedRange.Cells.ClearContents
    wshWbkTest.Protect
    
    mWbk.Value(wshWbkTest, "celValueUnlocked") = "Test-Value-Unlocked"
    Debug.Assert mWbk.Value(wshWbkTest, "celValueUnlocked") = "Test-Value-Unlocked"
    mWbk.Value(wshWbkTest, "celValueLocked") = "Test-Value-Locked"
    Debug.Assert mWbk.Value(wshWbkTest, "celValueLocked") = "Test-Value-Locked"
    
    '~~ Test 2: Range as object
    wshWbkTest.Unprotect
    wshWbkTest.UsedRange.Cells.ClearContents
    wshWbkTest.Protect
    
    Set Rng = Range("celValueUnlocked")
    mWbk.Value(wshWbkTest, Rng) = "Test-Value-Unlocked"
    Debug.Assert mWbk.Value(wshWbkTest, Rng) = "Test-Value-Unlocked"
    Set Rng = Range("celValueLocked")
    mWbk.Value(wshWbkTest, Rng) = "Test-Value-Locked"
    Debug.Assert mWbk.Value(wshWbkTest, Rng) = "Test-Value-Locked"
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_06_Exists_Service()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC              As String = "Test_06_Exists_Service"
    Const TEST_RANGE_NAME   As String = "celTestRangeName"
    Const TEST_WS_NAME      As String = "Exists-Test"
    Const TEST_WS_CODE_NAME As String = "wsAny"
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim wb1             As Workbook
    Dim wb2             As Workbook
    Dim sWb1FullName    As String
    Dim sWb2FullName    As String
    Dim sWb1Name        As String
    Dim sWb2Name        As String
    
    mBasic.BoP ErrSrc(PROC)
    sWb1FullName = ThisWorkbook.Path & "\Test\Test1.xlsm"
    sWb1Name = fso.GetFileName(sWb1FullName)
    Set wb1 = mWbk.GetOpen(sWb1FullName)
    sWb2FullName = ThisWorkbook.Path & "\Test\TestSubFolder\Test2.xlsm"
    sWb2Name = fso.GetFileName(sWb2FullName)
    ThisWorkbook.Activate

    '~~ Test 1: Workbook exists
    Debug.Assert mWbk.Exists(sWb1FullName) = True
    Debug.Assert mWbk.Exists(sWb1FullName & "x") = False
    Debug.Assert mWbk.Exists(sWb1Name) = True ' able to check existence though onl the name is provided because the Workbook is open
    
    '~~ Test 2: Worksheet exists
    Debug.Assert mWbk.Exists(sWb1Name, TEST_WS_NAME & "x") = False
    Debug.Assert mWbk.Exists(sWb1Name, TEST_WS_CODE_NAME & "x") = False
    Debug.Assert mWbk.Exists(sWb1Name, TEST_WS_NAME) = True
    Debug.Assert mWbk.Exists(sWb1Name, TEST_WS_CODE_NAME) = True
    
    
    '~~ Test 3: Range Name exists
    Debug.Assert mWbk.Exists(ex_wbk:=sWb1Name, ex_wsh:=TEST_WS_NAME, ex_range_name:=TEST_RANGE_NAME) = True
    Debug.Assert mWbk.Exists(ex_wbk:=sWb1Name, ex_wsh:=TEST_WS_CODE_NAME, ex_range_name:=TEST_RANGE_NAME) = True
    
    '~~ Test 4: Error conditions
    '~~ Test 4-1: Workbook is not open (AppErr(1)
    mErH.Asserted AppErr(1)
    Debug.Print mWbk.Exists(sWb2Name, TEST_WS_CODE_NAME)

xt: mWbk.WbClose wb1
    Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
