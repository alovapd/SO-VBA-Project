VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImmediateReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' ImmediateReporter v2.0.0-beta
' (c) Tim Hall - https://github.com/VBA-tools/Excel-TDD
'
' Report results to Immediate Window
'
' @class ImmediateReporter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private WithEvents pSpecs As SpecSuite
Attribute pSpecs.VB_VarHelpID = -1
Private Finished As Boolean

' ============================================= '
' Public Methods
' ============================================= '

''
' Listen to given SpecSuite
'
' @method ListenTo
' @param {SpecSuite} Specs
''
Public Sub ListenTo(specs As SpecSuite)
    If Not pSpecs Is Nothing Then
        Done
    End If
    
    Debug.Print "===" & IIf(specs.Description <> "", " " & specs.Description & " ===", "")
    Set pSpecs = specs
    Finished = False
End Sub

''
' Finish report for SpecSuite
'
' @method Done
''
Public Function Done()
    Finished = True
    
    Debug.Print "= " & Summary & " = " & Now & " =" & vbNewLine
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function ResultTypeToString(ResultType As SpecResultType) As String
    Select Case ResultType
    Case SpecResultType.Pass
        ResultTypeToString = "+"
    Case SpecResultType.Fail
        ResultTypeToString = "X"
    Case SpecResultType.Pending
        ResultTypeToString = "."
    End Select
End Function

Private Function Summary() As String
    Dim Total As Long
    Dim Passed As Long
    Dim Failed As Long
    Dim Pending As Long
    Total = pSpecs.specs.Count
    Passed = pSpecs.PassedSpecs.Count
    Failed = pSpecs.FailedSpecs.Count
    Pending = pSpecs.PendingSpecs.Count
    
    Dim SummaryMessage As String
    If Failed > 0 Then
        SummaryMessage = "FAIL (" & Failed & " of " & Total & " failed"
    Else
        SummaryMessage = "PASS (" & Passed & " of " & Total & " passed"
    End If
    If Pending > 0 Then
        SummaryMessage = SummaryMessage & ", " & Pending & " pending)"
    Else
        SummaryMessage = SummaryMessage & ")"
    End If
    
    Summary = SummaryMessage
End Function

Private Sub pSpecs_Result(spec As SpecDefinition)
    Debug.Print ResultTypeToString(spec.Result) & " " & spec.Description & IIf(spec.Id <> "", " [" & spec.Id & "]", "")
    
    If spec.Result = SpecResultType.Fail Then
        Dim Expectation As SpecExpectation
        For Each Expectation In spec.FailedExpectations
            Debug.Print "  " & Expectation.FailureMessage
        Next Expectation
    End If
End Sub

Private Sub Class_Terminate()
    If Not Finished Then
        Done
    End If
End Sub
