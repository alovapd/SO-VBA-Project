VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fm_Request 
   Caption         =   "CID Case Referal"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4635
   OleObjectBlob   =   "fm_Request.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fm_Request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")

    strsql = "SELECT [name],[position] FROM [Customs$] WHERE [name] IS NOT NULL AND [position] = '" & "Deputy" & "'"
    
    rst.Open strsql, cnnThisWorkbook, 3, 1
    
    Do Until rst.EOF = True
        Me.cb_ReqDep.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [name],[position] FROM [Customs$] WHERE [name] IS NOT NULL AND [position] = '" & "Sergeant" & "'"
    
    rst.Open strsql, cnnThisWorkbook, 3, 1
    
    Do Until rst.EOF = True
        If rst.fields(0).Value <> "Waldon" Then
            Me.cb_Sergeant.AddItem rst.fields(0).Value
        End If
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [name],[position] FROM [Customs$] WHERE [name] IS NOT NULL AND [position] = '" & "Corporal" & "'"
    
    rst.Open strsql, cnnThisWorkbook, 3, 1
    
    Do Until rst.EOF = True
        If rst.fields(0).Value <> "Penkava" Then
            Me.cb_Corporal.AddItem rst.fields(0).Value
        End If
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [Crimes] FROM [Customs$] WHERE [Crimes] IS NOT NULL"
    
    rst.Open strsql, cnnThisWorkbook, 3, 1
    
    Do Until rst.EOF = True
        Me.cb_Crimes.AddItem rst.fields(0).Value
        rst.movenext
    Loop
         
End Sub

Private Sub btn_Next_Click()

    Dim FormOk As Boolean
    
    fm_Request.Hide
    FormOk = CheckForm("cb_ReqDep,txt_CaseNum,cb_Crimes,cb_Sergeant,cb_Corporal")
    
    If FormOk = True Then
        fm_Comments.Show
    Else
        fm_Request.Show
    End If
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     Unload Me
     End
End Sub

