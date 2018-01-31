VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fm_Password 
   Caption         =   "Select Queue"
   ClientHeight    =   2370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4860
   OleObjectBlob   =   "fm_Password.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fm_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    Dim filter As String
    
    Select Case btnPressed
        Case "btn_CheckSgtQue"
            concat filter, "'Sergeant'"
        Case "btn_CheckLTQue"
            concat filter, "'Lieutenant'"
        Case "btn_CheckCptQue"
            concat filter, "'Captain'"
        Case "btn_CheckSheriffQue"
            concat filter, "'Sheriff'"
        Case "btn_EditForm"
            concat filter, "'Deputy'"
        Case Else
            Stop '{DEV}
    End Select
    
    strsql = "SELECT [name],[position],[Team],[Division] FROM [Customs$] WHERE lcase([position]) IN (" & filter & ")"
    
    rst.Open strsql, cnnThisWorkbook, 3, 1
        
    Do Until rst.EOF = True
        If rst.fields("division").Value & "" <> "Traffic" Then
            Me.cb_User.AddItem rst.fields("name").Value
        Else
            rst.movenext
        End If
        rst.movenext
    Loop
    
        
End Sub


Private Sub btnSubmit_Click()

    Dim arrUsers As Variant
    Dim User As String
    Dim PassPass As Boolean
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    Dim strsql As String
    
    PassPass = Passes
    
    If PassPass = False Then
        MsgBox "The password you entered isn't valid. Please retry.", vbOKOnly + vbCritical, "Password Incorrect"
    Else
        fm_QueResults.Show
    End If
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     Unload Me
     End
End Sub


