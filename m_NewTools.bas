Attribute VB_Name = "m_NewTools"
Option Explicit
Private pcnnThisWorkbook As ADODB.Connection

Public Property Get cnnThisWorkbook() As ADODB.Connection
   If pcnnThisWorkbook Is Nothing Then
       Set pcnnThisWorkbook = New ADODB.Connection
       pcnnThisWorkbook.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
           "Data Source=" & ThisWorkbook.Path & "\" & ThisWorkbook.Name & ";" & _
           "Extended Properties=""Excel 8.0;HDR=Yes"";"
   End If
   Set cnnThisWorkbook = pcnnThisWorkbook
End Property
