Attribute VB_Name = "m_Mailer"
Option Explicit

Sub Mailer(SendMailTo As String, CaseNum As String, ReqDeputy As String, FormName As String, Optional WhatLevel As String, Optional ApproveDenyComments As String, Optional Sergeant As String, Optional Lieutenant As String, Optional Captain As String)

    Dim strEmailBody As String
    Dim cnn As Object: Set cnn = CreateObject("adodb.connection")
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    Dim strsql As String
    Dim SgtEmail As String
    Dim CplEmail As String
    Dim CplName As String
    Dim Comments As String
    Dim abrev1 As String
    Dim abrev2 As String
    Dim emailTo As String
    Dim ForwardTo As String
    Dim EmailToAddress As String
    Dim SubjectName As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=true"";"
    
    Select Case btnPressed
        Case "btn_CheckSgtQue"
            Select Case ApproveDenyComments
                Case "DenyWithComments"
                    Comments = fm_Comments.txt_Notes.Value
                    FillInApproveOrDeny CaseNum, "SergeantDenied", ApproveDenyComments
                    emailTo = ReqDeputy
                    abrev1 = "Dep."
                    abrev2 = "Sgt."
                Case "DenyWithoutComments"
                    FillInApproveOrDeny CaseNum, "SergeantDenied", ApproveDenyComments
                    emailTo = ReqDeputy
                    abrev1 = "Dep."
                    abrev2 = "Sgt."
                Case "ApproveWithComments"
                    Comments = fm_Comments.txt_Notes.Value
                    FillInApproveOrDeny CaseNum, "SergeantApproved", ApproveDenyComments
                    emailTo = Lieutenant
                    abrev1 = "Lt."
                Case "ApproveWithoutComments"
                    FillInApproveOrDeny CaseNum, "SergeantApproved", ApproveDenyComments
                    emailTo = Lieutenant
                    abrev1 = "Lt."
            End Select
            SubjectName = Sergeant
        Case "btn_CheckLTQue"
            Select Case ApproveDenyComments
                Case "DenyWithComments"
                    Comments = fm_Comments.txt_Notes.Value
                    FillInApproveOrDeny CaseNum, "LieutenantDenied", ApproveDenyComments
                    emailTo = Sergeant
                Case "DenyWithoutComments"
                    FillInApproveOrDeny CaseNum, "LieutenantDenied", ApproveDenyComments
                    emailTo = Sergeant
                Case "ApproveWithComments"
                    Comments = fm_Comments.txt_Notes.Value
                    FillInApproveOrDeny CaseNum, "LieutenantApproved", ApproveDenyComments
                    emailTo = Captain
                Case "ApproveWithoutComments"
                    FillInApproveOrDeny CaseNum, "LieutenantApproved", ApproveDenyComments
                    emailTo = Captain
            End Select
            SubjectName = Lieutenant
        Case "btn_CheckCptQue"
            Select Case ApproveDenyComments
                Case "DenyWithComments"
                    Comments = fm_Comments.txt_Notes.Value
                    FillInApproveOrDeny CaseNum, "CaptainDenied", ApproveDenyComments
                    emailTo = Lieutenant
                Case "DenyWithoutComments"
                    FillInApproveOrDeny CaseNum, "CaptainDenied", ApproveDenyComments
                    emailTo = Lieutenant
                Case "ApproveWithComments"
                    Comments = fm_Comments.txt_Notes.Value
                    FillInApproveOrDeny CaseNum, "CaptainApproved", ApproveDenyComments
                Case "ApproveWithoutComments"
                    FillInApproveOrDeny CaseNum, "CaptainApproved", ApproveDenyComments
            End Select
            SubjectName = Captain
        Case Else
            Stop '{DEV}
    End Select
    strsql = "SELECT [name],[email] FROM [Customs$] WHERE [name] ='" & emailTo & "'"
    strEmailBody = "<p>"
    strEmailBody = strEmailBody & "Hello " & abrev1 & " " & emailTo & ",<br><br>"
    strEmailBody = strEmailBody & "<i>&#42***This is an Automated Message***&#42</i>" & "<br><br>"
    strEmailBody = strEmailBody & "<i><b><u>There is no need to reply unless you feel you need to address an issue reported below.</u></b></i>" & "<br><br>"
    
    FormName = Mid(FormName, InStrRev(FormName, "_") + 1)
    
    If MasterFormName.FrameApprove.Visible = True Then
        Select Case ApproveDenyComments
            Case "DenyWithComments"
                strEmailBody = strEmailBody & "I have denied your " & FormName & " and provided the following comments:" & "<br><br>"
                strEmailBody = strEmailBody & "<i>" & Comments & "<i><br><br>"
                strEmailBody = strEmailBody & "Please check this form and resubmit for approval."
                strEmailBody = strEmailBody & "Thank You!" & "<br><br>"
            Case "DenyWithoutComments"
                strEmailBody = strEmailBody & "I have denied your " & FormName & "<br><br>"
                strEmailBody = strEmailBody & "Please check this form and resubmit for approval."
                strEmailBody = strEmailBody & "Thank You!" & "<br><br>"
            Case "ApproveWithComments"
                strEmailBody = strEmailBody & "I have appvored your " & FormName & " and it is being forwarded to " & abrev2 & " " & emailTo & " for approval." & "<br><br> """
                strEmailBody = strEmailBody & "I noted the following:" & "<br><br>"
                strEmailBody = strEmailBody & "<i>" & Comments & "<i><br><br>"
                strEmailBody = strEmailBody & "Thank You!" & "<br><br>"
            Case "ApproveWithoutComments"
                strEmailBody = strEmailBody & "I have appvored your " & FormName & " and it is being forwarded to " & abrev2 & " " & emailTo & " for approval." & "<br><br> """
                strEmailBody = strEmailBody & "Thank You!" & "<br><br>"
        End Select
        rst.Open strsql, cnn, 3, 1
        Do Until rst.EOF = True
            EmailToAddress = rst.fields(1).Value
            rst.movenext
        Loop
        rst.Close
        
        CreateMail EmailToAddress, abrev2 & " " & SubjectName & " responded to your " & FormName & " submission for Case Number: " & CaseNum, strEmailBody, CplEmail
    Else
        strEmailBody = strEmailBody & "Deputy " & ReqDeputy & " has submitted a " & FormName & " for your approval. Please check your queue and approve or deny as necessary." & "<br><br>"
        strEmailBody = strEmailBody & "Thank You!" & "<br><br>"
        CreateMail SgtEmail, "Deputy " & ReqDeputy & " submitted an " & FormName & " into Your Queue", strEmailBody, CplEmail
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


Sub CreateMail(To_ As String, Subject_ As String, Body_ As String, Optional CC_ As String, Optional From_ As String, Optional Attachments_ As Variant, Optional Template As String)

    Dim objOutlook As Object
    Dim objMail As Object
    Dim Attachment As Variant
    Dim signature As String

    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)

    objMail.Display
    signature = objMail.htmlbody
    objMail.Close 1

    Set objMail = objOutlook.CreateItem(0)
    With objMail
        .BodyFormat = 2
        If From_ <> vbNullString Then .FROM = From_
        .To = To_
        .subject = Subject_
        .cc = CC_
        .htmlbody = Body_ & signature
        If IsArray(Attachments_) Then
          For Each Attachment In Attachments_
              .attachments.Add Attachment
          Next Attachment
        End If
        .Save
        '.Display 'Instead of .Display, you can use .Send to send the email or .Save to save a copy in the drafts folder
    End With

End Sub


