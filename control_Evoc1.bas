Attribute VB_Name = "control_Evoc1"

Public Sub initializeForm(frm As UserForm)

    Dim cnn As Object: Set cnn = CreateObject("adodb.connection")
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    
    Dim x As Integer

    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=true"";"
    strsql = "SELECT [ReasonsTerminated] FROM [Customs$] WHERE [ReasonsTerminated] IS NOT NULL ORDER BY [ReasonsTerminated] ASC"
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        frm.lb_TermReson.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [name] FROM [Customs$] ORDER BY [name] ASC"
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        frm.cb_Deputy.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [Category] FROM [Customs$] WHERE [Category] IS NOT NULL "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        frm.cb_Category.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [Lighting] FROM [Customs$] WHERE [Lighting] IS NOT NULL "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        frm.cb_Lighting.AddItem rst.fields(0).Value
        rst.movenext
    Loop
         
    rst.Close
    
    strsql = "SELECT [Weather] FROM [Customs$] WHERE [Weather] IS NOT NULL "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        frm.cb_Weather.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [RoadSurface] FROM [Customs$] WHERE [RoadSurface] IS NOT NULL "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        frm.cb_RoadSurface.AddItem rst.fields(0).Value
        rst.movenext
    Loop
    
    rst.Close
    
    strsql = "SELECT [name],[position] FROM [Customs$] WHERE [position] = 'Sergeant' OR [position] = 'Corporal' "
    
    rst.Open strsql, cnn, 3, 1
    
    Do Until rst.EOF = True
        If rst.fields(0).Value = "Penkava" Then
            rst.movenext
        Else
            frm.cb_OICName.AddItem rst.fields(0).Value
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
            frm.cb_Sergeant.AddItem rst.fields(0).Value
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
            frm.cb_Lieutenant.AddItem rst.fields(0).Value
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
            frm.cb_Captain.AddItem rst.fields(0).Value
            rst.movenext
        End If
    Loop
    
    For x = 1 To 4
        frm.cb_TeamNum.AddItem x
    Next x
    
    frm.ScrollHeight = 658
    If btnPressed = "btn_EditForm" Then frm.FrameApprove.Visible = False
    frm.FrameApprove.BorderColor = vbRed
    frm.FrameApprove.ForeColor = vbRed
    
End Sub

Public Sub addbuttonpush(frm As UserForm)
    
    frm.lb_TermSelected.Clear
    
    For i = 0 To frm.lb_TermReson.ListCount - 1
        If frm.lb_TermReson.Selected(i) = True Then frm.lb_TermSelected.AddItem frm.lb_TermReson.List(i)
    Next i
    
End Sub

Public Sub removeoptions(frm As UserForm)

    Dim counter As Integer
    counter = 0
    
    For i = 0 To frm.lb_TermSelected.ListCount - 1
        If frm.lb_TermSelected.Selected(i - counter) Then
            frm.lb_TermSelected.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
    
    frm.cbox2Toggle.Value = False

End Sub

Public Sub toggle1button(frm As UserForm)

    If frm.cbox1Toggle.Value = True Then
        For i = 0 To frm.lb_TermReson.ListCount - 1
            frm.lb_TermReson.Selected(i) = True
        Next i
    End If
    
    If frm.cbox1Toggle.Value = False Then
        For i = 0 To frm.lb_TermReson.ListCount - 1
            frm.lb_TermReson.Selected(i) = False
        Next i
    End If

End Sub

Public Sub toggle2button(frm As UserForm)

    If frm.cbox2Toggle.Value = True Then
        For i = 0 To frm.lb_TermSelected.ListCount - 1
            frm.lb_TermSelected.Selected(i) = True
        Next i
    End If
    
    If frm.cbox2Toggle.Value = False Then
        For i = 0 To frm.lb_TermSelected.ListCount - 1
            frm.lb_TermSelected.Selected(i) = False
        Next i
    End If

End Sub

Public Sub submitform(frm As UserForm)
    
    With frm
        .Hide
        If .FrameApprove.Visible = True Then
            If .obDeny.Value = True Then
                If .checkBoxAddComments.Value = True Then
                    .Hide
                    ApproveDenyComments = "DenyWithComments"
                    FillComments ApproveDenyComments
                Else
                    .Hide
                    ApproveDenyComments = "DenyWithoutComments"
                    FillComments ApproveDenyComments
                End If
            Else
                If .checkBoxAddComments.Value = True Then
                    .Hide
                    ApproveDenyComments = "ApproveWithComments"
                    FillComments ApproveDenyComments
                Else
                    .Hide
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

Public Sub redXpushed(frm As UserForm, Cancel As Integer, CloseMode As Integer)
     Unload frm
     End
End Sub




