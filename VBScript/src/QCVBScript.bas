Attribute VB_Name = "QCVBScript"
'@Folder "ipm-modules.QC.VBScript.src"
Option Explicit

Public Function NewRegExp(Optional ByVal g As Boolean, _
                          Optional ByVal i As Boolean, _
                          Optional ByVal m As Boolean, _
                          Optional ByVal Pattern As String) As Object
    Set NewRegExp = CreateObject("VBScript.RegExp")
    With NewRegExp
        .Global = g
        .IgnoreCase = i
        .MultiLine = m
        .Pattern = Pattern
    End With
End Function
