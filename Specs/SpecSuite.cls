VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' SpecSuite v2.0.0-beta
' (c) Tim Hall - https://github.com/VBA-tools/Excel-TDD
'
' A collection of specs and results
'
' @class SpecSuite
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Types, Events, and Properties
' --------------------------------------------- '

Public Enum SpecResultType
    Pass
    Fail
    Pending
End Enum

Public Event BeforeEach()
Public Event Result(Definition As SpecDefinition)
Public Event AfterEach()

''
' (Optional) description of suite for display in runners
'
' @property Description
' @type String
''
Public Description As String

''
' @property Specs
' @type Collection<SpecDefinition>
''
Public specs As VBA.Collection

''
' @property Result
' @type SpecResultType
''
Public Property Get Result() As SpecResultType
    Result = SpecResultType.Pending

    Dim spec As SpecDefinition
    For Each spec In Me.specs
        If spec.Result = SpecResultType.Pass Then
            Result = SpecResultType.Pass
        ElseIf spec.Result = SpecResultType.Fail Then
            Result = SpecResultType.Fail
            Exit For
        End If
    Next spec
End Property

''
' @property PassedSpecs
' @type Collection<SpecDefinition>
''
Public Property Get PassedSpecs() As Collection
    Dim spec As SpecDefinition
    Dim Filtered As New Collection
    For Each spec In Me.specs
        If spec.Result = SpecResultType.Pass Then
            Filtered.Add spec
        End If
    Next spec

    Set PassedSpecs = Filtered
End Property

''
' @property FailedSpecs
' @type Collection<SpecDefinition>
''
Public Property Get FailedSpecs() As Collection
    Dim spec As SpecDefinition
    Dim Filtered As New Collection
    For Each spec In Me.specs
        If spec.Result = SpecResultType.Fail Then
            Filtered.Add spec
        End If
    Next spec
    
    Set FailedSpecs = Filtered
End Property

''
' @property PendingSpecs
' @type Collection<SpecDefinition>
''
Public Property Get PendingSpecs() As Collection
    Dim spec As SpecDefinition
    Dim Filtered As New Collection
    For Each spec In Me.specs
        If spec.Result = SpecResultType.Pending Then
            Filtered.Add spec
        End If
    Next spec
    
    Set PendingSpecs = Filtered
End Property
''
'

' ============================================= '
' Public Methods
' ============================================= '

''
' Create a new spec definition with description
'
' @method It
' @param {String} Description
' @param {String} [Id = ""] Useful for identifying specific specs
' @returns {SpecDefinition}
''
Public Function It(Description As String, Optional Id As String = "") As SpecDefinition
    Dim spec As New SpecDefinition
    
    RaiseEvent BeforeEach
    
    ' Prepare Spec
    spec.Description = Description
    spec.Id = Id
    Set spec.Suite = Me

    Set It = spec
End Function

''
' Called at completion of SpecDefinition
'
' @internal
' @method SpecDone
' @param {SpecDefinition} Spec
''
Public Sub SpecDone(spec As SpecDefinition)
    Me.specs.Add spec
    RaiseEvent Result(spec)
    RaiseEvent AfterEach
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Sub Class_Initialize()
    Set Me.specs = New VBA.Collection
End Sub
