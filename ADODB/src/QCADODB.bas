Attribute VB_Name = "QCADODB"
'@Folder "ipm-modules.QC.ADODB.src"
Option Explicit

Public Function NewConnection(Optional ByVal ConnectionString As String) As Object
    Set NewConnection = CreateObject("ADODB.Connection")
    NewConnection.ConnectionString = ConnectionString
End Function

Public Function NewRecordset() As Object
    Set NewRecordset = CreateObject("ADODB.Recordset")
End Function

Public Function NewCommand() As Object
    Set NewCommand = CreateObject("ADODB.Command")
End Function

Public Function NewParameter() As Object
    Set NewParameter = CreateObject("ADODB.Parameter")
End Function

Public Function NewRecord() As Object
    Set NewRecord = CreateObject("ADODB.Record")
End Function

Public Function NewStream() As Object
    Set NewStream = CreateObject("ADODB.Stream")
End Function
