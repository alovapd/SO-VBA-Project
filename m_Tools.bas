Attribute VB_Name = "m_Tools"
Option Explicit
Sub btn_Request()
    fm_Request.Show
End Sub


Public Function CheckForm(ReqFields As String) As Boolean
    
    Dim item As Variant
    Dim arrReqFields As Variant
    Dim v As Variant
    Dim status As Boolean
    Dim GoodCaseNum As Boolean
    Dim arrTermSelected As Variant
    
    status = True
    
    arrReqFields = Split(ReqFields, ",")
    
    For Each item In arrReqFields
        For Each v In fm_Request.Controls
            If item = v.Name Then
                If item = "txt_CaseNum" And v.Value <> "" Then
                    GoodCaseNum = CaseFormatIsCorrect(fm_Request.txt_CaseNum.Value)
                    If GoodCaseNum = False Then
                        v.BorderStyle = 1
                        v.BorderColor = vbRed
                        fm_Request.lbl_CaseNum.ForeColor = vbRed
                        fm_Request.lbl_CaseNum.Visible = True
                        status = False
                        Exit For
                    Else
                        v.BorderStyle = 0
                        v.SpecialEffect = 2
                        fm_Request.lbl_CaseNum.Visible = False
                        CheckForm = True
                    End If
                Else
                    If v.Value = "" Then
                        v.BorderStyle = 1
                        v.BorderColor = vbRed
                        status = False
                        Exit For
                    Else
                        v.BorderStyle = 0
                        v.SpecialEffect = 2
                        CheckForm = True
                    End If
                End If
            End If
        Next v
    Next item
    
    If status = False Then
        fm_Request.lbl_Blank.ForeColor = vbRed
        fm_Request.lbl_Blank.Visible = True
        CheckForm = False
    End If

End Function

Public Function CaseFormatIsCorrect(CaseNumber As String) As Boolean

   Dim regex As Object
   Dim matches As Object
   Set regex = CreateObject("vbscript.regexp")
   Const Pattern = "^\d{2}-[1-9]\d{0,4}$"
   With regex
       .MultiLine = False
       .Global = False
       .IgnoreCase = True
       .Pattern = Pattern
   End With
   
   Set matches = regex.Execute(CaseNumber)
   If matches.Count > 0 Then
       CaseFormatIsCorrect = True
   End If
   
End Function

Public Function CasePresent(CaseNum As String) As Boolean

    Dim cnn As Object: Set cnn = CreateObject("adodb.connection")
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    Dim strsql As String

    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=true"";"
    strsql = "SELECT [CaseNum] FROM [Data$] WHERE [CaseNum] IS NOT NULL AND [CaseNum] = '" & CaseNum & "'"
    
    rst.Open strsql, cnn, 3, 1
    
    If rst.EOF = True Or rst.bof = True Then
        CasePresent = False
    Else
        CasePresent = True
    End If
    
End Function

Public Sub FillDataCIDRequest()
    
    Dim workingrow As Long
    
    With ThisWorkbook.Sheets("Data").ListObjects(1)
        If .DataBodyRange(1, 1).Value = "" Then
            .DataBodyRange(1, 1).Value = fm_Request.cb_ReqDep.Value
            .DataBodyRange(1, 2).Value = fm_Request.txt_CaseNum.Value
            .DataBodyRange(1, 3).Value = fm_Request.cb_Crimes.Value
            .DataBodyRange(1, 4).Value = Date
            .DataBodyRange(1, 5).Value = fm_Request.cb_Sergeant.Value
            .DataBodyRange(1, 6).Value = fm_Request.cb_Corporal.Value
            .DataBodyRange(1, 11).Value = fm_Comments.txt_Notes.Value
        Else
            .ListRows.Add
            workingrow = .ListRows.Count
            .DataBodyRange(workingrow, 1).Value = fm_Request.cb_ReqDep.Value
            .DataBodyRange(workingrow, 2).Value = fm_Request.txt_CaseNum.Value
            .DataBodyRange(workingrow, 3).Value = fm_Request.cb_Crimes.Value
            .DataBodyRange(workingrow, 4).Value = Date
            .DataBodyRange(workingrow, 5).Value = fm_Request.cb_Sergeant.Value
            .DataBodyRange(workingrow, 6).Value = fm_Request.cb_Corporal.Value
            .DataBodyRange(workingrow, 11).Value = fm_Comments.txt_Notes.Value
        End If
    End With
    
    'Mailer fm_Request.cb_Sergeant.Value, fm_Request.cb_Corporal.Value, fm_Request.txt_CaseNum.Value, fm_Request.cb_ReqDep.Value
    
End Sub

Public Sub FillDataTable(SheetName As Variant, FormName As Object, Optional CommentForm As Boolean, Optional CommentFormName As Object)
    
    Dim v As Variant
    Dim item As Variant
    Dim CaseNum As String
    Dim i As Long
    Dim FirstRow As Boolean
    Dim x As Long
    Dim arrTermSelected As Variant
    Dim strTermSelected As String
    Dim thing As Variant
    
    CaseNum = FormName.txt_CaseNum.Value
    
    For v = 1 To ThisWorkbook.Sheets(SheetName).ListObjects(1).ListColumns.Count
        For Each item In FormName.Controls
            'Debug.Print Mid(item.Name, InStrRev(item.Name, "_") + 1, Len(item.Name)) & " - " & ThisWorkbook.Sheets(SheetName).ListObjects(1).HeaderRowRange(v).Value
                If InStrRev(item.Name, "_") <> 0 Then
                    If Mid(item.Name, InStrRev(item.Name, "_") + 1, Len(item.Name)) = ThisWorkbook.Sheets(SheetName).ListObjects(1).HeaderRowRange(v).Value Then
                        If ThisWorkbook.Sheets(SheetName).ListObjects(1).DataBodyRange(1, 1) = "" Then
                            ThisWorkbook.Sheets(SheetName).ListObjects(1).DataBodyRange(1, 1).Value = item.Value
                            FirstRow = True
                            Exit For
                        Else
                            If ThisWorkbook.Sheets(SheetName).ListObjects(1).DataBodyRange(ThisWorkbook.Sheets(SheetName).ListObjects(1).ListRows.Count, 1).Value <> CaseNum Then
                                ThisWorkbook.Sheets(SheetName).ListObjects(1).ListRows.Add
                            End If
                            If Left(item.Name, 2) <> "lb" And Left(item.Name, 2) <> "fm" Then
                                ThisWorkbook.Sheets(SheetName).ListObjects(1).DataBodyRange(ThisWorkbook.Sheets(SheetName).ListObjects(1).ListRows.Count, v).Value = item.Value
                            Else
                                If Left(item.Name, 2) = "lb" Then
                                    If item.ListCount > 0 Then
                                        ReDim arrTermSelected(item.ListCount - 1)
                                        For x = 0 To item.ListCount - 1
                                            arrTermSelected(x) = item.List(x)
                                            If strTermSelected = "" Then
                                                strTermSelected = item.List(x) & ","
                                            Else
                                                strTermSelected = strTermSelected & item.List(x) & ","
                                            End If
                                        Next x
                                        strTermSelected = Left(strTermSelected, Len(strTermSelected) - 1)
                                        ThisWorkbook.Sheets(SheetName).ListObjects(1).DataBodyRange(ThisWorkbook.Sheets(SheetName).ListObjects(1).ListRows.Count, v).Value = strTermSelected
                                    End If
                                Else
                                    For Each thing In FormName.fm_OtherUnits.Controls
                                        If thing <> "" Then
                                            Dim strUnits As String
                                            If strUnits = "" Then
                                                strUnits = UCase(thing) & ","
                                            Else
                                                strUnits = strUnits & UCase(thing) & ","
                                            End If
                                        End If
                                    Next thing
                                    If strUnits = "" Then Exit For
                                    strUnits = Left(strUnits, Len(strUnits) - 1)
                                    ThisWorkbook.Sheets(SheetName).ListObjects(1).DataBodyRange(ThisWorkbook.Sheets(SheetName).ListObjects(1).ListRows.Count, v).Value = strUnits
                                End If
                            End If
                            Exit For
                        End If
                    End If
                End If
        Next item
    Next v
    
    If CommentForm = True Then
        For i = 1 To ThisWorkbook.Sheets(SheetName).ListObjects(1).HeaderRowRange.Count
            If ThisWorkbook.Sheets(SheetName).ListObjects(1).HeaderRowRange(i).Value = "Comments" Then
                If FirstRow = True Then
                    ThisWorkbook.Sheets(SheetName).ListObjects(1).DataBodyRange(1, i).Value = CommentFormName.txt_Notes.Value
                    Else
                    ThisWorkbook.Sheets(SheetName).ListObjects(1).DataBodyRange(ThisWorkbook.Sheets(SheetName).ListObjects(1).ListRows.Count, v).Value = CommentFormName.txt_Notes.Value
                End If
            End If
        Next i
    End If
    
    With ThisWorkbook.Sheets(SheetName)
        .Cells.Select
        .Cells.EntireColumn.AutoFit
        .Cells(1, 1).Select
    End With
    
    Dim answer As Integer
    Dim FormType As String
    Dim DataSheet As String
    FormType = Mid(FormName.Name, InStrRev(FormName.Name, "_") + 1, Len(FormName.Name))
    DataSheet = "Data" & FormType
    answer = MsgBox("Is the form ready to be submitted to your supervisor?", vbYesNo, "Submit For Approval?")
    If answer = 6 Then
        With ThisWorkbook.Sheets(DataSheet).ListObjects(1)
            For x = 1 To .ListRows.Count
                Debug.Print .DataBodyRange(x, 1).Value & " <> " & FormName.txt_CaseNum.Value
                If .DataBodyRange(x, 1).Value = FormName.txt_CaseNum.Value Then
                    For i = 30 To .HeaderRowRange.Count
                        Debug.Print .HeaderRowRange(i).Value & " - " & "PatrolApproved"
                        If .HeaderRowRange(i).Value = "PatrolApproved" Then
                            .DataBodyRange(x, i).Value = Date
                            Exit For
                            Stop
                        End If
                    Next i
                Exit For
                End If
            Next x
        End With
        Mailer FormName.cb_Sergeant.Value, FormName.txt_CaseNum.Value, FormName.cb_Deputy.Value, FormType, FormName.cb_TeamNum.Value
    End If

    Stop
    
End Sub

Public Sub CreateFormforPrint(CaseNum As String, SheetName As String)

    Dim cnn As Object: Set cnn = CreateObject("adodb.connection")
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    Dim strAllHeaders As String
    Dim MySheetName As String
    Dim NewSheetName As String
    Dim RangeName As String
    Dim DataFieldNAme As String
    Dim CellRange As Range
    Dim arrItems As Variant
    Dim i As Long
    Dim v As Long
    Dim x As Long
    Dim j As Long
    Dim strsql As String
        
    MySheetName = "Data" & StrConv(Replace(SheetName, " ", ""), vbProperCase)
    NewSheetName = Replace(SheetName, " ", "")
    
    ThisWorkbook.Sheets(NewSheetName).Select
    
    For x = 1 To ThisWorkbook.Sheets(MySheetName).ListObjects(1).ListColumns.Count
        If strAllHeaders = "" Then
            strAllHeaders = "[" & ThisWorkbook.Sheets(MySheetName).ListObjects(1).HeaderRowRange(x).Value & "]"
        Else
            strAllHeaders = strAllHeaders & ",[" & ThisWorkbook.Sheets(MySheetName).ListObjects(1).HeaderRowRange(x).Value & "]"
        End If
    Next x

    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=true"";"
    strsql = "SELECT " & strAllHeaders & "FROM [" & MySheetName & "$] WHERE [CaseNum] IS NOT NULL AND [CaseNum] = '" & CaseNum & "'"
    
    rst.Open strsql, cnn, 3, 1
    
    For v = 1 To ThisWorkbook.Names.Count - 1
        RangeName = ThisWorkbook.Names(v).Name
        If Left(RangeName, Len(SheetName) - 1) = NewSheetName Then
            x = 0
            For x = 0 To rst.fields.Count - 1
                If Mid(RangeName, InStrRev(RangeName, "_") + 1, Len(RangeName)) = rst.fields(x).Name Then
                    If IsNull(rst.fields(x).Value) Then
                        ThisWorkbook.Sheets(NewSheetName).Range(RangeName).Value = ""
                    Else
                        If Mid(RangeName, InStrRev(RangeName, "_") + 1, Len(RangeName)) = "TermSelected" Then
                            ThisWorkbook.Sheets(NewSheetName).Range(RangeName).ClearContents
                            arrItems = Split(rst.fields(x).Value, ",")
                            For j = 0 To UBound(arrItems)
                                ThisWorkbook.Sheets(NewSheetName).Cells(j + 27, 4).Value = arrItems(j)
                            Next j
                        Else
                            ThisWorkbook.Sheets(NewSheetName).Range(RangeName).Value = rst.fields(x).Value
                        End If
                        Exit For
                    End If
                End If
            Next x
        End If
    Next v

End Sub

Public Sub FillFormFromDataBase(CaseNum As String, SheetName As String, FormName As Object, Optional Tag As String, Optional WhosQue As String)

    Dim cnn As Object: Set cnn = CreateObject("adodb.connection")
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    Dim x As Long
    Dim v As Long
    Dim y As Long
    Dim z As Long
    Dim arrUnits As Variant
    Dim filter As String
    Dim MySheetName As String
    Dim NewSheetName As String
    Dim item As Variant
    Dim strsql As String
    Dim thingy As Variant
    
    MySheetName = "Data" & StrConv(Replace(SheetName, " ", ""), vbProperCase)
    NewSheetName = Replace(SheetName, " ", "")

    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=true"";"
    
    For Each item In FormName.Controls
        If InStrRev(item.Name, "_") <> 0 Then
            If Left(Mid(item.Name, InStrRev(item.Name, "_") + 1, Len(item.Name)), 9) <> "OtherUnit" And Mid(item.Name, InStrRev(item.Name, "_") + 1, Len(item.Name)) <> "TermSelected" And Mid(item.Name, InStrRev(item.Name, "_") + 1, Len(item.Name)) <> "TermReson" Then
                strsql = "SELECT [" & Mid(item.Name, InStrRev(item.Name, "_") + 1, Len(item.Name)) & "] " & "FROM [" & MySheetName & "$] WHERE [CaseNum] IS NOT NULL AND [CaseNum] = '" & CaseNum & "'"
                'Debug.Print strsql
                rst.Open strsql, cnn, 3, 1
                item.Value = rst.fields(0).Value
                rst.Close
            End If
        End If
    Next item
    
    For Each item In FormName.Controls
        If Left(Mid(item.Name, InStrRev(item.Name, "_") + 1, Len(item.Name)), 9) = "OtherUnit" Then
            strsql = "SELECT [OtherUnits] FROM [" & MySheetName & "$] WHERE [CaseNum] IS NOT NULL AND [CaseNum] = '" & CaseNum & "'"
            rst.Open strsql, cnn, 3, 1
            If Not IsEmpty(arrUnits) Then
                arrUnits = Split(rst.fields(0).Value, ",")
                z = 0
                For Each thingy In FormName.fm_OtherUnits.Controls
                    If z = UBound(arrUnits) - LBound(arrUnits) + 1 Then Exit For
                    thingy.Value = arrUnits(z)
                    z = z + 1
                Next thingy
                rst.Close
                Exit For
            End If
            rst.Close
        End If
    Next item
    
    For Each item In FormName.Controls
        If Mid(item.Name, InStrRev(item.Name, "_") + 1, Len(item.Name)) = "TermSelected" Then
            strsql = "SELECT [TermSelected] FROM [" & MySheetName & "$] WHERE [CaseNum] IS NOT NULL AND [CaseNum] = '" & CaseNum & "'"
            rst.Open strsql, cnn, 3, 1
            arrUnits = Split(rst.fields(0).Value, ",")
            For z = 0 To UBound(arrUnits)
                item.AddItem arrUnits(z)
            Next z
            Exit For
        End If
    Next item
    
    If Tag = "Queue" Then
        FormName.FrameApprove.Visible = True
        FormName.lblApprovalAuthority.Caption = WhosQue
    End If
    FormName.Show
    
    'DEV TemNum didnt show on form

End Sub

Sub FrmUnloadAll()

    Dim frm As UserForm

    For Each frm In UserForms
        Unload frm
    Next frm

End Sub

Public Function OutlookOpen() As Boolean

    Dim oOutlook As Object

    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If oOutlook Is Nothing Then
        OutlookOpen = False
    Else
        OutlookOpen = True
    End If

End Function

Public Sub FillComments(WhatAction As String)

    Dim i As Long
    Dim x As Long
    Dim CaseNum As String
    Dim Notes As String
    Dim SheetName As String
    Dim FormName As Object
    
    Select Case WhatAction
        Case "DenyWithComments"
            fm_Comments.Show
        Case "DenyWithoutComments"
            Stop
        Case "ApproveWithComments"
            fm_Comments.Show
        Case "ApproveWithoutComments"
            Stop
        Case Else
            Stop
    End Select

    Notes = fm_Comments.txt_Notes.Value
    Set FormName = WhatFormSubmitted
    CaseNum = FormName.txt_CaseNum.Value
    SheetName = Replace("Data" & Mid(FormName.Name, InStrRev(FormName.Name, "_"), Len(Mid(FormName.Name, InStrRev(FormName.Name, "_")))), "_", "")
    
    With ThisWorkbook.Sheets(SheetName).ListObjects(1)
        For x = 1 To .ListRows.Count
            If .DataBodyRange(x, 1).Value = CaseNum Then
                For i = 32 To .HeaderRowRange.Count
                    If .HeaderRowRange(i).Value = "SupComments" Then
                        .DataBodyRange(x, i).Value = Notes
                        Exit For
                    End If
                Next i
                Exit For
            End If
        Next x
    End With
    
    Mailer "", CaseNum, FormName.cb_Deputy.Value, FormName.Name, btnPressed, ApproveDenyComments, fm_Password.cb_User.Value, FormName.cb_Lieutenant.Value, FormName.cb_Captain.Value
    
End Sub

Public Function WhatFormSubmitted() As Object

    Dim frm As Object
    Dim arrFormNames
    Dim strFormNames
    Dim x As Long

    For Each frm In UserForms
        If frm.Name Like "*#*" Then
            Set WhatFormSubmitted = frm
            Set MasterFormName = frm
        End If
    Next frm

End Function

Public Sub FillInApproveOrDeny(CaseNum As String, WhatLevel As String, Method As String)

    Dim x As Long
    Dim i As Long
    
    With ThisWorkbook.Sheets(SheetNameData).ListObjects(1)
        For i = 0 To .ListRows.Count
            If .DataBodyRange(i, 1).Value = CaseNum Then
                For x = 30 To .ListColumns.Count
                    If .HeaderRowRange(x).Value = WhatLevel Then
                        .DataBodyRange(i, x).Value = Date
                        Exit For
                    End If
                Next x
            Exit For
            End If
        Next i
    End With
End Sub

