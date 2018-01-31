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


Private Sub UserForm_Initialize()
    
    Dim cnn As Object: Set cnn = CreateObject("adodb.connection")
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    Dim x As Integer

    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=true"";"
    strsql = "SELECT [ReasonsTerminated] FROM [Customs$] WHERE [ReasonsTerminated] IS NOT NULL ORDER BY [ReasonsTerminated] ASC"
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        Me.lb_TermReson.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [name] FROM [Customs$] ORDER BY [name] ASC"
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        Me.cb_Deputy.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [Category] FROM [Customs$] WHERE [Category] IS NOT NULL "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        Me.cb_Category.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [Lighting] FROM [Customs$] WHERE [Lighting] IS NOT NULL "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        Me.cb_Lighting.AddItem rst.fields(0).Value
        rst.movenext
    Loop
         
    rst.Close
    
    strsql = "SELECT [Weather] FROM [Customs$] WHERE [Weather] IS NOT NULL "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        Me.cb_Weather.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [RoadSurface] FROM [Customs$] WHERE [RoadSurface] IS NOT NULL "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        Me.cb_RoadSurface.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [name],[position] FROM [Customs$] WHERE [position] = 'Sergeant' OR [position] = 'Corporal' "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        If rst.fields(0).Value = "Penkava" Then
            rst.movenext
        Else
            Me.cb_OICName.AddItem rst.fields(0).Value
            rst.movenext
        End If
    Loop
    
    rst.Close
    
    strsql = "SELECT [name],[position] FROM [Customs$] WHERE [position] = 'Sergeant'"
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        If rst.fields(0).Value = "Penkava" Then
            rst.movenext
        Else
            Me.cb_Sergeant.AddItem rst.fields(0).Value
            rst.movenext
        End If
    Loop
    
    rst.Close
    
    strsql = "SELECT [name],[position] FROM [Customs$] WHERE [position] = 'Lieutenant' "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        If rst.fields(0).Value = "Penkava" Then
            rst.movenext
        Else
            Me.cb_Lieutenant.AddItem rst.fields(0).Value
            rst.movenext
        End If
    Loop
    
    rst.Close
    
    strsql = "SELECT [name],[position] FROM [Customs$] WHERE [position] = 'Captain' "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        If rst.fields(0).Value = "Penkava" Then
            rst.movenext
        Else
            Me.cb_Captain.AddItem rst.fields(0).Value
            rst.movenext
        End If
    Loop
    
    For x = 1 To 4
        Me.cb_TeamNum.AddItem x
    Next x
    
    Me.Height = 600
    Me.ScrollHeight = 658
    If btnPressed = "btn_EditForm" Then Me.FrameApprove.Visible = False
    Me.FrameApprove.BorderColor = vbRed
    Me.FrameApprove.ForeColor = vbRed
    
End Sub

Private Sub frmbtnOptionAdd_Click()

    lb_TermSelected.Clear
    
    For i = 0 To lb_TermReson.ListCount - 1
        If lb_TermReson.Selected(i) = True Then lb_TermSelected.AddItem lb_TermReson.List(i)
    Next i

End Sub

Private Sub frmbtnOptionRemove_Click()

    Dim counter As Integer
    counter = 0
    
    For i = 0 To lb_TermSelected.ListCount - 1
        If lb_TermSelected.Selected(i - counter) Then
            lb_TermSelected.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
    
    frmcbox2Toggle.Value = False

End Sub

Private Sub frmcbox1Toggle_Click()

    If frmcbox1Toggle.Value = True Then
        For i = 0 To lb_TermReson.ListCount - 1
            lb_TermReson.Selected(i) = True
        Next i
    End If
    
    If frmcbox1Toggle.Value = False Then
        For i = 0 To lb_TermReson.ListCount - 1
            lb_TermReson.Selected(i) = False
        Next i
    End If

End Sub

Private Sub frmcbox2Toggle_Click()

    If frmcbox2Toggle.Value = True Then
        For i = 0 To lb_TermSelected.ListCount - 1
            lb_TermSelected.Selected(i) = True
        Next i
    End If
    
    If frmcbox2Toggle.Value = False Then
        For i = 0 To lb_TermSelected.ListCount - 1
            lb_TermSelected.Selected(i) = False
        Next i
    End If

End Sub

Private Sub btnSubmit_Click()
    
    With Me
        .Hide
        If .FrameApprove.Visible = True Then
            If .obDeny.Value = True Then
                If .checkBoxAddComments.Value = True Then
                    Me.Hide
                    ApproveDenyComments = "DenyWithComments"
                    FillComments ApproveDenyComments
                Else
                    Me.Hide
                    ApproveDenyComments = "DenyWithoutComments"
                    FillComments ApproveDenyComments
                End If
            Else
                If .checkBoxAddComments.Value = True Then
                    Me.Hide
                    ApproveDenyComments = "ApproveWithComments"
                    FillComments ApproveDenyComments
                Else
                    Me.Hide
                    ApproveDenyComments = "ApproveWithoutComments"
                    FillComments ApproveDenyComments
                End If
            End If
        End If
    End With

    If MasterFormName.FrameApprove.Visible = False Then
        FillDataTable "DataEvoc1", fm_Evoc1
    End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     Unload Me
     End
End Sub



