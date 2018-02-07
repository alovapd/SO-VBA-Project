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

Public WithEvents SATermSelected As SelectAllListBox
Public WithEvents SATermReason As SelectAllListBox

Private Sub UserForm_Initialize()
    control_Evoc1.initializeForm Me
    Me.Height = 600
    Set SATermSelected = New SelectAllListBox
    Set SATermReason = New SelectAllListBox
    SATermSelected.Initialize Me.lb_TermSelected, Me.cbox2Toggle
    SATermReason.Initialize Me.lb_TermReson, Me.cbox1Toggle
    
End Sub

Private Sub btnOptionAdd_Click()
    SATermReason.SendSelectedItems SATermSelected.ListBox
End Sub

Private Sub btnOptionRemove_Click()
    SATermSelected.SendSelectedItems SATermReason.ListBox
End Sub

Private Sub btnSubmit_Click()
    control_Evoc1.submitform Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    control_Evoc1.redXpushed Me, Cancel, CloseMode
End Sub



