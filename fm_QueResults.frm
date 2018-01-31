VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fm_QueResults 
   Caption         =   "Your Queue Results"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5565
   OleObjectBlob   =   "fm_QueResults.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fm_QueResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WhosQueue As String
Private Sub UserForm_Initialize()
    
    Dim filter As String
    Dim filter2 As String
    Dim filter3 As String
    Dim filter4 As String
    Dim filter5 As String
    Dim filter6 As String
    Dim arrDataBaseSheetNames As Variant
    Dim strDataBaseSheetNames As String
    Dim x As Long
    
    fm_Password.Hide
    
    strDataBaseSheetNames = "DataEvoc1$,DataEvoc2$,DataDT1$"
    arrDataBaseSheetNames = Split(strDataBaseSheetNames, ",")
        
    Select Case btnPressed
        Case "btn_CheckSgtQue"
            filter = "SergeantApproved"
            filter2 = "PatrolApproved"
            filter5 = "SergeantDenied"
            filter3 = fm_Password.cb_User.Value
            WhosQueue = fm_Password.cb_User.Value
            filter4 = "Sergeant"
        Case "btn_CheckLTQue"
            filter = "LieutenantApproved"
            filter2 = "SergeantApproved"
            filter3 = fm_Password.cb_User.Value
            WhosQueue = fm_Password.cb_User.Value
            filter4 = "Lieutenant"
            filter5 = "LieutenantDenied"
        Case "btn_CheckCptQue"
            filter = "CaptainApproved"
            filter2 = "LieutenantApproved"
            filter3 = fm_Password.cb_User.Value
            WhosQueue = fm_Password.cb_User.Value
            filter4 = "Captain"
            filter5 = "CaptainDenied"
        Case "btn_EditForm"
            filter = "PatrolApproved"
            filter2 = "SergeantDenied"
            filter3 = fm_Password.cb_User.Value
            WhosQueue = fm_Password.cb_User.Value
            filter4 = "Deputy"
        Case Else
            Stop '{DEV}
    End Select
    
    For x = 0 To UBound(arrDataBaseSheetNames) - 1
        If btnPressed <> "btn_EditForm" Then
            strsql = "SELECT [CaseNum],[" & filter4 & "] FROM [" & arrDataBaseSheetNames(x) & "] WHERE [" & filter2 & "] IS NOT NULL AND [" & filter & "] IS NULL AND [" & filter5 & "] IS NULL AND [" & filter4 & "] = '" & filter3 & "'"
        Else
            strsql = "SELECT [CaseNum],[" & filter4 & "] FROM [" & arrDataBaseSheetNames(x) & "] WHERE [" & filter2 & "] IS NOT NULL AND [" & filter4 & "] = '" & filter3 & "'"
        End If
        Debug.Print strsql
        rst.Open strsql, cnnThisWorkbook, 3, 1
        
        
        Do Until rst.EOF = True
            Me.ListBox_DetCrimes.AddItem rst.fields(0).Value & "   (" & Mid(arrDataBaseSheetNames(x), 5, Len(Mid(arrDataBaseSheetNames(x), 5)) - 1) & ")"
            rst.movenext
        Loop
        rst.Close
        Exit For
    Next x
         
End Sub

Private Sub frmbtn_OptionAdd_Click()

    ListBox_DetChoices.Clear
    
    For i = 0 To ListBox_DetCrimes.ListCount - 1
        If ListBox_DetCrimes.Selected(i) = True Then ListBox_DetChoices.AddItem ListBox_DetCrimes.List(i)
    Next i

End Sub

Private Sub frmbtn_OptionRemove_Click()

    Dim counter As Integer
    counter = 0
    
    For i = 0 To ListBox_DetChoices.ListCount - 1
        If ListBox_DetChoices.Selected(i - counter) Then
            ListBox_DetChoices.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
    
    frmcbox2_Toggle.Value = False

End Sub

Private Sub frmcbox1_Toggle_Click()

    If frmcbox1_Toggle.Value = True Then
        For i = 0 To ListBox_DetCrimes.ListCount - 1
            ListBox_DetCrimes.Selected(i) = True
        Next i
    End If
    
    If frmcbox1_Toggle.Value = False Then
        For i = 0 To ListBox_DetCrimes.ListCount - 1
            ListBox_DetCrimes.Selected(i) = False
        Next i
    End If

End Sub

Private Sub frmcbox2_Toggle_Click()

    If frmcbox2_Toggle.Value = True Then
        For i = 0 To ListBox_DetChoices.ListCount - 1
            ListBox_DetChoices.Selected(i) = True
        Next i
    End If
    
    If frmcbox2_Toggle.Value = False Then
        For i = 0 To ListBox_DetChoices.ListCount - 1
            ListBox_DetChoices.Selected(i) = False
        Next i
    End If

End Sub

Private Sub btn_View_Click()
    
    Dim rst As Object: Set rst = CreateObject("adodb.recordset")
    Dim strsql As String
    Dim SheetName As String
    Dim CaseNumber As String
    Dim FormName As String
    
    Me.Hide
    
    SheetName = Mid(Me.ListBox_DetChoices.List(0), InStrRev(Me.ListBox_DetChoices.List(0), "(") + 1, Len(Mid(Me.ListBox_DetChoices.List(0), InStrRev(Me.ListBox_DetChoices.List(0), "(") + 1)) - 1)
    CaseNumber = Trim(Left(Me.ListBox_DetChoices.List(0), InStrRev(Me.ListBox_DetChoices.List(0), "(") - 1))
    FormName = Mid(Me.ListBox_DetChoices.List(0), InStrRev(Me.ListBox_DetChoices.List(0), "(") + 1, Len(Mid(Me.ListBox_DetChoices.List(0), InStrRev(Me.ListBox_DetChoices.List(0), "(") + 1)) - 1)
    
    Select Case FormName
        Case "CIDReferal"
        
        Case "DT1"
        
        Case "Evoc1"
            FillFormFromDataBase CaseNumber, SheetName, fm_Evoc1, "Queue", WhosQueue
        Case "Evoc2"
    End Select
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     Unload Me
     End
End Sub


