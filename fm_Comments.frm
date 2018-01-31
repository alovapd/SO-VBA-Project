VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fm_Comments 
   Caption         =   "Your Notes:"
   ClientHeight    =   6315
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13680
   OleObjectBlob   =   "fm_Comments.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fm_Comments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub UserForm_Initialize()
    
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    Dim note As String
    Dim strsql As String
    Dim SheetName As String
    Dim CaseNumber As String
    
    SheetNameData = "Data" & Mid(fm_QueResults.ListBox_DetChoices.List(0), InStrRev(fm_QueResults.ListBox_DetChoices.List(0), "(") + 1, Len(Mid(fm_QueResults.ListBox_DetChoices.List(0), InStrRev(fm_QueResults.ListBox_DetChoices.List(0), "(") + 1)) - 1)
    CaseNumber = Trim(Left(fm_QueResults.ListBox_DetChoices.List(0), InStrRev(fm_QueResults.ListBox_DetChoices.List(0), "(") - 1))
    
    strsql = "SELECT [SupComments] FROM [" & SheetNameData & "$] WHERE [CaseNum] = '" & CaseNumber & "'"
    
    rst.Open strsql, cnnThisWorkbook, 3, 1
    
    With Me.txt_Notes
        .WordWrap = True
        .MultiLine = True
        .EnterKeyBehavior = True
        .ScrollBars = fmScrollBarsVertical
        .Value = rst.fields(0).Value
        note = Me.txt_Notes.Value
    End With
    If Len(note) = 0 Then
        Me.lbl_count.Visible = False
    End If
    
End Sub

Sub txt_Notes_Change()

    Dim i As String
    Dim FirstNumInt As Long
    Dim FieldChars As Long
    Dim FirstNum As String

    FieldChars = CLng(Len(Me.txt_Notes.Value))
    i = "32767"
    
    If Len(Me.txt_Notes.Value) >= 0 Then
        Me.lbl_count.Visible = True
        FirstNumInt = i - FieldChars
        FirstNum = CStr(FirstNumInt)
        Me.lbl_count.Caption = Format(FirstNum, "#,##0") & " of 32,767 chars left"
    End If
    
End Sub

Private Sub btn_Submit_Click()

    Dim CaseThere As Boolean
    Dim x As Long
    Dim counter As Integer
    Dim workingrow As Long
    Dim WhatForm As String

'    CaseThere = CasePresent(fm_Request.txt_CaseNum.Value)
'
'    Do Until CaseThere = False
'        If Me.lbl_ErrorCount.Caption = 2 Then
'            Me.Hide
'            MsgBox "You seem to be trying to enter the same case multiple times. Please contact Deputy Osborne for further assistance.", vbOKOnly, "Error"
'            End
'        Else
'            fm_Comments.Hide
'            Me.lbl_ErrorCount.Caption = Me.lbl_ErrorCount.Caption + 1
'            MsgBox "The case number you are trying to enter seems to already have been used. Please check the case number and try again.", vbOKOnly, "Case Present"
'            fm_Request.Show
'        End If
'    Loop
    
    Me.Hide
    
    'Mailer

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     Unload Me
     End
End Sub

