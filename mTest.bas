Attribute VB_Name = "mTest"
Option Explicit
' ------------------------------------------------------------
' Standard Module mTest Test of all Existence checks variants
'                       in module mExists
' -----------------------------------------------------------
'Declare API
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Integer) As Integer
Const SHIFT_KEY = 16

Private Sub Regression_Test()
' ---------------------------------------------------------
' All results are asserted and there is no intervention
' required for the whole test. When an assertion fails the
' test procedure will stop and indicates the problem with
' the called procedure.
' An execution trace is displayed for each test procedure.
' ---------------------------------------------------------
Const PROC = "Test_All"

    On Error GoTo on_error
    
    Test_IsOpen
    Test_GetOpen
    Test_GetOpen_Errors
    
exit_proc:
    Exit Sub
    
on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Sub Test_GetOpen()
Const PROC          As String = "Test_GetOpen"  ' This procedure's name for the error handling and execution tracking
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

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    '~~ Ensure precondition
    On Error Resume Next
    wb1.Close
    wb2.Close
    On Error GoTo on_error
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

    On Error GoTo on_error
    Debug.Assert GetOpen(sFullName).FullName = sFullName

    '~~ Test 6: GetOpen Workbook by full name (not open)
    '~~         A Workbook with the same name but from a different location is already open
    '~~         and the file does not/no longer exist at the provided location.
    Set wb3 = Workbooks.Open(sWb3FullName)
    Debug.Assert GetOpen(sWb1FullName & "\Test2.xlsm").Name = sWb3Name
    wb3.Close
    
exit_proc:
    '~~ Cleanup
    On Error Resume Next
    wb1.Close
    wb2.Close
    wb3.Close
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Sub Test_GetOpen_Errors()
Const PROC = "Test_GetOpen_Errors"
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

    BoP ErrSrc(PROC)
    On Error GoTo on_error
    
    '~~ Ensure precondition
    sWb1Name = "Test1.xlsm"
    sWb1FullName = ThisWorkbook.Path & "\" & sWb1Name
    sWb2Name = sWb1Name
    sWb2FullName = ThisWorkbook.Path & "\Test\" & sWb2Name
    sWb3Name = "Test2.xlsm"
    sWb3FullName = ThisWorkbook.Path & "\Test\" & sWb3Name
    
    '~~ Test : GetOpen Workbook is object never opened
    On Error Resume Next
    GetOpen wb1
    Debug.Assert AppErr(Err.Number) = 1

    '~~ Test E-2: Parameter is Nothing
    Set wb1 = Workbooks.Open(sWb1FullName) ' open the test Workbook
    wb1.Close
    On Error Resume Next
    GetOpen wb1
    Debug.Assert AppErr(Err.Number) = 2

    '~~ Test E-2: Parameter is Nothing
    Set wb1 = Nothing
    On Error Resume Next
    GetOpen wb1
    Debug.Assert AppErr(Err.Number) = 1
    
    '~~ Test E-3: Parameter is a not open Workbook's name
    On Error Resume Next
    wb1.Close
    wb2.Close
    On Error Resume Next
    GetOpen sWb1Name
    Debug.Assert AppErr(Err.Number) = 5

    '~~ Test E-4: Parameter is a Workbook's full name but the file does't exist
    On Error Resume Next
    GetOpen Replace(sWb1FullName, sWb1Name, "not-existing.xls")
    Debug.Assert AppErr(Err.Number) = 4

    '~~ Test E-5: A Workbook with the provided name is open but from a different location
    '             and the Workbook file still exists at the provided location
    Close wb1
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\" & "Test3.xlsm")
    On Error Resume Next
    Set wb1 = GetOpen(ThisWorkbook.Path & "\Test\" & "Test3.xlsm")
    Debug.Assert AppErr(Err.Number) = 3
    wb.Close
    
    '~~ Test E-6: Parameter is neither a Workbook object nor a string
    On Error Resume Next
    Set wb = GetOpen(ThisWorkbook.ActiveSheet)
    Debug.Assert AppErr(Err.Number) = 1

    '~~ Cleanup
    On Error Resume Next
    wb1.Close
    wb2.Close
    wb3.Close
exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Sub Test_IsOpen()
Const PROC      As String = "Test_IsOpen"  ' This procedure's name for the error handling and execution tracking
Dim wb          As Workbook
Dim sName       As String
Dim o           As Object
Dim wb1         As Workbook
Dim wb2         As Workbook
Dim wb3         As Workbook
Dim wbResult    As Workbook

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    '~~ Prepare test environment
    On Error Resume Next
    wb1.Close
    wb2.Close
    wb3.Close
    On Error GoTo on_error
    With Workbooks
        Set wb1 = .Open(ThisWorkbook.Path & "\Test1.xlsm")
        Set wb2 = .Open(ThisWorkbook.Path & "\Test\Test2.xlsm")
        Set wb3 = .Open(ThisWorkbook.Path & "\Test\Test3.xlsm")
    End With
    
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
    Debug.Assert IsOpen(wb1.Path & "\Test3.xlsm", wbResult) = False
    
exit_proc:
    wb1.Close
    wb2.Close
    wb3.Close
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Sub Test_Is_()
Dim wb  As Workbook
Dim wb1 As Workbook

    Set wb = mWrkbk.GetOpen(ThisWorkbook.Path & "\" & "Test.xlsm")
    
    Debug.Assert IsName(wb.Name) = True
    Debug.Assert IsName(wb.FullName) = False
    Debug.Assert IsName(wb.Path) = False
    Debug.Assert IsName(ThisWorkbook) = False
    
    Debug.Assert IsPath(wb.Name) = False
    Debug.Assert IsPath(wb.FullName) = False
    Debug.Assert IsPath(wb.Path) = True
    Debug.Assert IsPath(ThisWorkbook) = False
    
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
        
End Sub
 
Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest" & "." & sProc
End Function
