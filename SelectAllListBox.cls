VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectAllListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ListBox As MSForms.ListBox
Attribute ListBox.VB_VarHelpID = -1
Public WithEvents CheckBox As MSForms.CheckBox
Attribute CheckBox.VB_VarHelpID = -1

Public Event Click()

Public Sub Initialize(ListBox As MSForms.ListBox, CheckBox As MSForms.CheckBox)

    Set Me.ListBox = ListBox
    Set Me.CheckBox = CheckBox

End Sub

Public Sub CheckBox_Click()
    
    Dim i As Long
    
    For i = 0 To Me.ListBox.ListCount - 1
        Me.ListBox.Selected(i) = Me.CheckBox.Value
    Next i

End Sub

Public Sub SendSelectedItems(target As MSForms.ListBox)

    Dim i As Long

    For i = 0 To Me.ListBox.ListCount - 1
        If Me.ListBox.Selected(i) = True Then
            target.AddItem Me.ListBox.List(i)
        End If
    Next i
    
    For i = Me.ListBox.ListCount - 1 To 0 Step -1
        If Me.ListBox.Selected(i) = True Then
            Me.ListBox.RemoveItem (i)
        End If
    Next i
    
    Me.CheckBox.Value = False

End Sub

Private Sub SortList()

    

End Sub


