VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fm_Evoc1 
   Caption         =   "Evoc One"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13410
   OleObjectBlob   =   "fm_Evoc1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fm_Evoc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    control_Evoc1.initializeForm Me
    Me.Height = 600
End Sub

Private Sub btnOptionAdd_Click()
    control_Evoc1.addbuttonpush Me
End Sub

Private Sub btnOptionRemove_Click()
    control_Evoc1.removeoptions Me
End Sub

Private Sub cbox1Toggle_Click()
    control_Evoc1.checkboxSelectAllFromListbox Me.cbox1Toggle, Me.lb_TermReson
End Sub

Private Sub cbox2Toggle_Click()
    control_Evoc1.checkboxSelectAllFromListbox Me.cbox2Toggle, Me.lb_TermSelected
End Sub

Private Sub btnSubmit_Click()
    control_Evoc1.submitform Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    control_Evoc1.redXpushed Me, Cancel, CloseMode
End Sub



