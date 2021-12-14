Attribute VB_Name = "mWrkbkTest"
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
    If err_source = vbNullString Then err_source = Err.Source
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
    ErrSrc = "mTest" & "." & sProc
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
    
    mErH.BoP ErrSrc(PROC)
    mTest.Test_01_IsOpen
    mTest.Test_02_GetOpen
    mTest.Test_03_GetOpen_Errors
    mTest.Test_04_Is_
    mTest.Test_05_Opened
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
        Set wb1 = .Open(ThisWorkbook.Path & "\Test\Test1.xlsm")
        Set wb2 = .Open(ThisWorkbook.Path & "\Test\TestSubFolder\Test2.xlsm")
        Set wb3 = .Open(ThisWorkbook.Path & "\Test\TestSubFolder\Test3.xlsm")
    End With
    
    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    '~~ 1. Test IsOpen by object
    Debug.Assert mWrkbk.IsOpen(wb1, wbResult) = True

    '~~ 2. Test IsOpen by Name
    Debug.Assert mWrkbk.IsOpen(wb1.Name, wbResult) = True

    '~~ 3. Test IsOpen by FullName
    Debug.Assert mWrkbk.IsOpen(wb1.FullName, wbResult) = True

    '~~ 4. A Workbook with the given name is open but from a different location
    '~~    Since the Workbook does not or no longer exist at the requested location it regarded moved and considered open
    Debug.Assert mWrkbk.IsOpen(wb1.Path & "\Test2.xlsm", wbResult) = True
    Debug.Assert wbResult.FullName = wb1.Path & "\TestSubFolder\Test2.xlsm"
    
    '~~ 4b No Workbook object is returned since the parameter is not Variant    Debug.Assert vWb Is wb2
    Debug.Assert mWrkbk.IsOpen(wb1.Path & "\Test2.xlsm", wbResult) = True
    
    '~~ 5. Workbook does not exist. When a fullname is provided an error is raised
    Debug.Assert mWrkbk.IsOpen(wb1.Path & "\Test\Test.xlsm", wbResult) = False
        
    '~~ 6. A Workbook with the given Name is open but from a different location
    '~~    Since it still exists at the requested location it is regarde not open
    wb3.Close
    Debug.Assert mWrkbk.IsOpen(wb1.Path & "\Test3.xlsm", wbResult) = False
    
xt: wb1.Close
    wb2.Close
    On Error Resume Next: wb3.Close
    mErH.EoP ErrSrc(PROC)
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
    
    mErH.BoP ErrSrc(PROC)
    
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
    Debug.Assert GetOpen(sFullName).FullName = sFullName

    '~~ Test 5: GetOpen Workbook by full name (not open)
    '~~         A Workbook with the same name but from a different location is already open.
    On Error Resume Next
    wb1.Close False
    wb2.Close False
    sFullName = wb3.FullName
    wb3.Close False

    On Error GoTo eh
    Debug.Assert GetOpen(sFullName).FullName = sFullName

    '~~ Test 6: GetOpen Workbook by full name (not open)
    '~~         A Workbook with the same name but from a different location is already open
    '~~         and the file does not/no longer exist at the provided location.
    Set wb3 = Workbooks.Open(sWb3FullName)
    Debug.Assert GetOpen(sWb1FullName & "\Test2.xlsm").Name = sWb3Name
    wb3.Close False
    
xt: '~~ Cleanup
    On Error Resume Next
    GetOpen(sWb1FullName).Close False
    wb1.Close False
    wb2.Close False
    wb3.Close False
    mErH.EoP ErrSrc(PROC)
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
        
    ' Prepare
    sWb1Name = "Test1.xlsm"
    sWb1FullName = ThisWorkbook.Path & "\" & sWb1Name
    sWb2Name = sWb1Name
    sWb2FullName = ThisWorkbook.Path & "\Test\" & sWb2Name
    sWb3Name = "Test2.xlsm"
    sWb3FullName = ThisWorkbook.Path & "\Test\" & sWb3Name
    
    '~~ Test : GetOpen Workbook is object never opened
    mErH.BoTP ErrSrc(PROC), AppErr(1) ' Bypass this error as the one asserted
    mWrkbk.GetOpen wb1
    mErH.EoP ErrSrc(PROC)
    Debug.Assert mErH.MostRecentError = AppErr(1)

    '~~ Test E-2: Parameter is Nothing
    Set wb1 = Workbooks.Open(sWb1FullName) ' open the test Workbook
    wb1.Close
    mErH.BoTP ErrSrc(PROC), AppErr(2) ' Bypass this error as the one asserted
    mWrkbk.GetOpen wb1
    mErH.EoP ErrSrc(PROC)
    Debug.Assert mErH.MostRecentError = AppErr(2)

    '~~ Test E-2: Parameter is Nothing
    Set wb1 = Nothing
    
    mErH.BoTP ErrSrc(PROC), AppErr(1) ' Bypass this error as the one asserted
    mWrkbk.GetOpen wb1
    mErH.EoP ErrSrc(PROC)
    Debug.Assert mErH.MostRecentError = AppErr(1)
    
    '~~ Test E-3: Parameter is a not open Workbook's name
    On Error Resume Next
    wb1.Close
    wb2.Close
    
    mErH.BoTP ErrSrc(PROC), AppErr(5) ' Bypass this error as the one asserted
    GetOpen sWb1Name
    mErH.EoP ErrSrc(PROC)
    Debug.Assert mErH.MostRecentError = AppErr(5)

    '~~ Test E-4: Parameter is a Workbook's full name but the file does't exist
    mErH.BoTP ErrSrc(PROC), AppErr(4) ' Bypass this error as the one asserted
    mWrkbk.GetOpen Replace(sWb1FullName, sWb1Name, "not-existing.xls")
    mErH.EoP ErrSrc(PROC)
    Debug.Assert mErH.MostRecentError = AppErr(4)

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
    Debug.Assert AppErr(1)

    '~~ Cleanup
    On Error Resume Next
    wb1.Close
    wb2.Close
    wb3.Close

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_04_Is_()
    Const PROC = "Test_04_Is_"
    
    On Error GoTo eh
    Dim wb  As Workbook
    Dim wb1 As Workbook
    Dim fso As New FileSystemObject
    
    mErH.BoP ErrSrc(PROC)
    
    Set wb = mWrkbk.GetOpen(ThisWorkbook.Path & "\Test\Test1.xlsm")
    
    Debug.Assert IsName(wb.Name) = True
    Debug.Assert IsName(wb.FullName) = True
    Debug.Assert IsName(wb.Path) = False
    Debug.Assert IsName(ThisWorkbook) = False
    Debug.Assert IsName(fso.GetBaseName(wb.FullName)) = False
    
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
    Debug.Assert IsObject(wb) = False               ' A closed Workbook is still an object but not an object Type Workbook
    Set wb = Nothing
    Debug.Assert IsObject(wb) = False               ' A set to Nothing is no longer a Workbook object
    Debug.Assert IsObject(wb1) = False              ' A never assigned Workbook is not a Workbook object
        
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

 
Public Sub Test_05_Opened()
    Const PROC = "Test_04_Is_"

    On Error GoTo eh
    mErH.BoP ErrSrc(PROC)
    Debug.Assert Opened.Count > 0
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
