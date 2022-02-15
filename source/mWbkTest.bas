Attribute VB_Name = "mWbkTest"
Option Explicit
' ------------------------------------------------------------
' Standard Module mTest Test of all Existence checks variants
'                       in module mExists
' -----------------------------------------------------------
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

  
Private Sub BoP(ByVal b_proc As String, _
          ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' service. When neither the Common Execution Trace
' Component (mTrc) nor the Common Error Handling Component (mErH) is installed
' (indicated by the Conditional Compile Arguments 'ExecTrace = 1' and/or the
' Conditional Compile Argument 'ErHComp = 1') this procedure does nothing.
' Else the service is handed over to the corresponding procedures.
' May be copied as Private Sub into any module or directly used when mBasic is
' installed.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case only the mTrc is installed but not the merH.
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' service. When neither the Common Execution Trace
' Component (mTrc) nor the Common Error Handling Component (mErH) is installed
' (indicated by the Conditional Compile Arguments 'ExecTrace = 1' and/or the
' Conditional Compile Argument 'ErHComp = 1') this procedure does nothing.
' Else the service is handed over to the corresponding procedures.
' May be copied as Private Sub into any module or directly used when mBasic is
' installed.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case the mTrc is installed but the merH is not.
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

Private Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_no As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option active
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section displaying text connected to an error
' message by two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

  
Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mWbkTest" & "." & sProc
End Function

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
    
    mTrc.LogTitle = "Execution trace log 'Regression Test' mWbk module"
    mErH.BoP ErrSrc(PROC)
    mErH.Regression = True
    mWbkTest.Test_01_IsOpen
    mWbkTest.Test_02_GetOpen
    mWbkTest.Test_03_GetOpen_Errors
    mWbkTest.Test_04_Is_Name_FullName_Object
    mWbkTest.Test_05_Opened
    mWbkTest.Test_06_Exists
    
xt: mErH.EoP ErrSrc(PROC)
    mErH.Regression = False
    RegressionKeepLog
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub RegressionKeepLog()
    Dim sFile As String

#If ExecTrace = 1 Then
#If MsgComp = 1 Or ErHComp = 1 Then
    '~~ avoid the error message when the Conditional Compile Argument 'MsgComp = 0'!
    mTrc.Dsply
#End If
    '~~ Keep the regression test result
    With New FileSystemObject
        sFile = .GetParentFolderName(mTrc.LogFile) & "\RegressionTest.log"
        If .FileExists(sFile) Then .DeleteFile (sFile)
        .GetFile(mTrc.LogFile).Name = "RegressionTest.log"
    End With
    mTrc.Terminate
#End If

End Sub

Public Sub Test_01_IsOpen()
    Const PROC = "Test_01_IsOpen"  ' This procedure's name for the error handling and execution tracking

    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim wb              As Workbook
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
    
    BoP ErrSrc(PROC)
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
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
    
    BoP ErrSrc(PROC)
    
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
    wb1.Close False
    Set wb1 = GetOpen(sFullName)
    Debug.Assert wb1.FullName = sFullName
    wb1.Close False
    
    '~~ Test 5: GetOpen Workbook by full name
    '~~         A Workbook with the same name but from a different location is already open
    '~~         and the file does not/no longer exist at the provided location.
    Set wb3 = Workbooks.Open(sWb3FullName)
    Debug.Assert GetOpen(Replace(sWb1FullName, "Test1", "Test2")).Name = sWb3Name
    wb3.Close False
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
        
    BoP ErrSrc(PROC)
    ' Prepare
    sWb1Name = "Test1.xlsm"
    sWb1FullName = ThisWorkbook.Path & "\" & sWb1Name
    sWb2Name = sWb1Name
    sWb2FullName = ThisWorkbook.Path & "\Test\" & sWb2Name
    sWb3Name = "Test2.xlsm"
    sWb3FullName = ThisWorkbook.Path & "\Test\" & sWb3Name
    
    '~~ Test : GetOpen Workbook is object never opened
    mErH.Asserted AppErr(1) ' skip display of error message when mErH.Regression = True
    mWbk.GetOpen wb1

    '~~ Test E-2: Parameter is Nothing
    If Not wb1 Is Nothing Then wb1.Close False
    mErH.Asserted AppErr(1) ' skip display of error message when mErH.Regression = True
    mWbk.GetOpen wb1

    '~~ Test E-2: Parameter is Nothing
    Set wb1 = Nothing
    mErH.Asserted AppErr(1) ' skip display of error message when mErH.Regression = True
    mWbk.GetOpen wb1
    
    '~~ Test E-3: Parameter is a not open Workbook's name
    mErH.Asserted AppErr(3) ' skip display of error message when mErH.Regression = True
    GetOpen sWb1Name

    '~~ Test E-4: Parameter is a Workbook's full name but the file does't exist
    mErH.Asserted AppErr(4) ' skip display of error message when mErH.Regression = True
    mWbk.GetOpen Replace(sWb1FullName, sWb1Name, "not-existing.xls")

    '~~ Test E-5: A Workbook with the provided name is open but from a different location
    '             and the Workbook file still exists at the provided location
    If Not wb1 Is Nothing Then wb1.Close False
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\Test\TestSubFolder\Test3.xlsm")
    mErH.Asserted AppErr(2)
    Set wb1 = GetOpen(ThisWorkbook.Path & "\Test\" & "Test3.xlsm")
    wb1.Close False
    
    '~~ Test E-6: Parameter is neither a Workbook object nor a string
    mErH.Asserted AppErr(1)
    Set wb = GetOpen(ThisWorkbook.ActiveSheet)

    '~~ Cleanup
    On Error Resume Next
    If Not wb1 Is Nothing Then wb1.Close False
    If Not wb2 Is Nothing Then wb2.Close False
    If Not wb3 Is Nothing Then wb3.Close False

xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_04_Is_Name_FullName_Object()
    Const PROC = "Test_04_Is_Name_FullName_Object"
    
    On Error GoTo eh
    Dim wb              As Workbook
    Dim wb1             As Workbook
    Dim fso             As New FileSystemObject
    Dim sWb1FullName    As String
    Dim sWb1Name        As String
    
    BoP ErrSrc(PROC)
    
    sWb1FullName = ThisWorkbook.Path & "\Test\Test1.xlsm"
    sWb1Name = fso.GetFileName(sWb1FullName)
    Set wb = mWbk.GetOpen(sWb1FullName)
    
    '~~ Test 1: IsName
    Debug.Assert IsName(wb.Name) = True
    Debug.Assert IsName(wb.FullName) = False
    Debug.Assert IsName(wb.Path) = False
    Debug.Assert IsName(ThisWorkbook) = False
    Debug.Assert IsName(fso.GetBaseName(wb.FullName)) = False
    
    '~~ Test 2: IsFullName
    Debug.Assert mWbk.IsFullName(wb.Name) = False
    Debug.Assert mWbk.IsFullName(wb.FullName) = True
    Debug.Assert mWbk.IsFullName(wb.Path) = False
    Debug.Assert mWbk.IsFullName(ThisWorkbook) = False

    '~~ Test 3: IsWbObject
    Debug.Assert mWbk.IsWbObject(wb.Name) = False
    Debug.Assert mWbk.IsWbObject(wb.FullName) = False
    Debug.Assert mWbk.IsWbObject(wb.Path) = False
    Debug.Assert mWbk.IsWbObject(ThisWorkbook) = True
    Debug.Assert mWbk.IsWbObject(wb) = True
    wb.Close
    Debug.Assert mWbk.IsWbObject(wb) = False               ' A closed Workbook is still an object but not an object Type Workbook
    Set wb = Nothing
    Debug.Assert mWbk.IsWbObject(wb) = False               ' A set to Nothing is no longer a Workbook object
    Debug.Assert mWbk.IsWbObject(wb1) = False              ' A never assigned Workbook is not a Workbook object
        
xt: Set fso = Nothing
    EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_05_Opened()
    Const PROC = "Test_04_Is_Name_FullName_Object"

    On Error GoTo eh
    BoP ErrSrc(PROC)
    Debug.Assert Opened.Count > 0
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_06_Exists()
    Const PROC              As String = "Test_06_Exists"
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
    
    BoP ErrSrc(PROC)
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
    Debug.Assert mWbk.Exists(ex_wb:=sWb1Name, ex_ws:=TEST_WS_NAME, ex_range_name:=TEST_RANGE_NAME) = True
    Debug.Assert mWbk.Exists(ex_wb:=sWb1Name, ex_ws:=TEST_WS_CODE_NAME, ex_range_name:=TEST_RANGE_NAME) = True
    
    '~~ Test 4: Error conditions
    '~~ Test 4-1: Workbook is not open (AppErr(1)
    mErH.Asserted AppErr(1)
    Debug.Print mWbk.Exists(sWb2Name, TEST_WS_CODE_NAME)

xt: mWbk.WbClose wb1
    Set fso = Nothing
    EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

