Attribute VB_Name = "RUI"
Public btnPressed As String
Public ApproveDenyComments As String
Public SheetNameData As String
Public SheetNameForm As String
Public MasterFormName As Object

Sub onAction(control As IRibbonControl)

    Dim OutApp  As Object
    Set OutApp = OutlookApp()
    
    btnPressed = control.ID
    Select Case control.ID
        Case "btn_CIDRequest": btn_Request
        Case "btn_Evoc1": btnEvoc1
        Case "btn_CreateForm": btnCreateForm
        Case "btn_EditForm": btnEditForm
        Case "btn_CheckSgtQue", "btn_CheckLTQue", "btn_CheckCptQue", "btn_CheckSheriffQue": btn_CheckQueue
    Case Else
            Debug.Print "There is no action specified for the """ & control.ID & """ button yet."
    End Select
    
    FrmUnloadAll
    
End Sub

Sub label(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case "tb_PatrolTools"
            returnedVal = "Patrol Tools"
        Case "btn_CIDRequest"
            returnedVal = "CID Referral"
        Case "Grp_One"
            returnedVal = ""
        Case "tb_SupervisorTools"
            returnedVal = "Supervisor Tools"
        Case "Grp_Two"
            returnedVal = ""
        Case "btn_CheckSgtQue"
            returnedVal = "Check Queue"
        Case "btn_CheckLTQue"
            returnedVal = "Check Queue"
        Case "btn_CheckCptQue"
            returnedVal = "Check Queue"
        Case "btn_CheckSheriffQue"
            returnedVal = "Check Queue"
        Case "btn_Evoc1"
            returnedVal = "EVOC 1"
        Case "btn_Evoc2"
            returnedVal = "EVOC 2"
        Case "btn_DT1"
            returnedVal = "DT 1"
        Case "btn_CreateForm"
            returnedVal = "Create Form"
        Case "btn_EditForm"
            returnedVal = "Edit Form"
        Case Else
            returnedVal = control.ID
            Debug.Print "There is no label specified for the """ & control.ID & """ button yet."
    End Select
End Sub

Sub Size(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        '' Big buttons
        Case "btn_CIDRequest", "btn_CheckSgtQue", "btn_CheckCplQue", "btn_CheckLTQue", "btn_CheckSheriffQue", "btn_CheckCptQue", "btn_Evoc1", "btn_Evoc2", "btn_DT1", "btn_CreateForm", "btn_EditForm"
            returnedVal = 1
        '' Small buttons
        Case ""
            returnedVal = 0
        Case Else
            returnedVal = 0
            Debug.Print "There is no size specified for the """ & control.ID & """ button yet."
    End Select
End Sub

Sub screenTip(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case "btn_CIDRequest", "btn_Evoc1", "btn_Evoc2", "btn_DT1", "btn_CreateForm", "btn_EditForm", "btn_CheckSgtQue", "btn_CheckLTQue", "btn_CheckCptQue", "btn_CheckSheriffQue"
            returnedVal = "What Does This Do?"
        Case Else
            returnedVal = control.ID
            Debug.Print "There is no screenTip specified for the """ & control.ID & """ button yet."
    End Select
End Sub

Sub superTip(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case "btn_CIDRequest"
            returnedVal = "Press this button to request CID investigate your case."
        Case "btn_Evoc1"
            returnedVal = "Complete an EVOC 1 form if you were the PRIMARY unit in a vehicle persuit."
        Case "btn_Evoc2"
            returnedVal = "Complete an EVOC 2 form if you were the SECONDARY unit in a vehicle persuit or you used a persuit intervention device or procedure."
        Case "btn_DT1"
            returnedVal = "Complete this form if you were involved in a Use of Force."
        Case "btn_CreateForm"
            returnedVal = "Click this button to create a form for printing."
        Case "btn_EditForm"
            returnedVal = "Click this button to edit a form."
        Case "btn_CheckSgtQue"
            returnedVal = "***PASSWPORD REQUIRED*** Click this button to Check the Sergeant Queue."
        Case "btn_CheckLTQue"
            returnedVal = "***PASSWPORD REQUIRED*** Click this button to Check the Lieutenant Queue."
        Case "btn_CheckCptQue"
            returnedVal = "***PASSWPORD REQUIRED*** Click this button to Check the Captain Queue."
        Case "btn_CheckSheriffQue"
            returnedVal = "***PASSWPORD REQUIRED*** Click this button to Check the Sheriff Queue."
        Case Else
            returnedVal = control.ID
            Debug.Print "There is no super tip specified for the """ & control.ID & """ button yet."
    End Select
End Sub

Sub getVisible(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        '' hide button elements by adding them here
        Case ""
            returnedVal = 0
        '' visible by default
        Case Else
            returnedVal = 1
    End Select
End Sub

Sub btnEvoc1()
    fm_Evoc1.Show
End Sub

Sub btnCreateForm()
    fm_CreateForm.Show
End Sub

Sub btnEditForm()
    fm_Password.Show
    fm_CreateForm.btnCreateForm.Caption = "Submit"
    fm_CreateForm.Show
End Sub

Sub btn_CheckQueue()
    fm_Password.Show
End Sub
