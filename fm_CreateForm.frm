VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fm_CreateForm 
   Caption         =   "Select Below"
   ClientHeight    =   2370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "fm_CreateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fm_CreateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Initialize()
    
    Dim cnn As Object: Set cnn = CreateObject("adodb.connection")
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")

    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=true"";"
    strsql = "SELECT [CreateFormSelection] FROM [Customs$] WHERE [CreateFormSelection] IS NOT NULL"
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        Me.cb_FormType.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
End Sub


Private Sub cb_FormType_Change()
    
    Dim cnn As Object: Set cnn = CreateObject("adodb.connection")
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    Dim SheetName As String
    Dim FormSelection As String
    
    FormSelection = Replace(Me.cb_FormType.Value, " ", "")
    SheetName = "Data" & Trim(FormSelection)
    
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=true"";"
    strsql = "SELECT [CaseNum] FROM [" & SheetName & "$] WHERE [CaseNum] IS NOT NULL"
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        Me.cb_CaseNum.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
End Sub

Private Sub btnCreateForm_Click()
        
        Dim WhatForm As String
        
        WhatForm = Me.cb_FormType.Value
        
        If Me.btnCreateForm.Caption = "Create Form" Then
            Select Case WhatForm
                Case "CID Referal"
                
                Case "DT 1"
                
                Case "Evoc 1"
                    CreateFormforPrint Me.cb_CaseNum.Value, Me.cb_FormType.Value
                Case "Evoc 2"
                            
            End Select
        Else
            Select Case WhatForm
                Case "CID Referal"
                
                Case "DT 1"
                
                Case "Evoc 1"
                    FillFormFromDataBase Me.cb_CaseNum.Value, Me.cb_FormType.Value, fm_Evoc1
                Case "Evoc 2"
                            
            End Select
        End If
        Me.Hide
        
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     Unload Me
     End
End Sub

