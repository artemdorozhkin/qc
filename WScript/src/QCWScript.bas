Attribute VB_Name = "QCWScript"
'@Folder "ipm-modules.QC.QCWScript.src"
Option Explicit

Public Function NewShell() As Object
    Set NewShell = CreateObject("WScript.Shell")
End Function

Public Function NewNetwork() As Object
    Set NewNetwork = CreateObject("WScript.Network")
End Function

Public Function NewShortCut(ByVal PathLink As String, Optional ByVal TargetPath As String, Optional ByVal StdIcon As qcStdIcons = EmptyFileIcon) As Object
    Set NewShortCut = NewShell().CreateShortcut(PathLink)
    If Len(TargetPath) > 0 Then NewShortCut.TargetPath = TargetPath
    NewShortCut.IconLocation = "%SystemRoot%\System32\shell32.dll," & StdIcon
End Function
