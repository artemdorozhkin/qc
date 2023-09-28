Attribute VB_Name = "QCMS"
'@Folder "ipm-modules.QC.MS.src"
Option Explicit

Public Function NewOutlook() As Object
    Set NewOutlook = CreateObject("Outlook.Application")
End Function

Public Function NewExcel() As Object
    Set NewExcel = CreateObject("Excel.Application")
End Function

Public Function NewWord() As Object
    Set NewWord = CreateObject("Word.Application")
End Function

Public Function NewAccess() As Object
    Set NewAccess = CreateObject("Access.Application")
End Function

Public Function NewPowerPoint() As Object
    Set NewPowerPoint = CreateObject("PowerPoint.Application")
End Function

Public Function NewMSProject() As Object
    Set NewMSProject = CreateObject("MSProject.Application")
End Function

Public Function NewPublisher() As Object
    Set NewPublisher = CreateObject("Publisher.Application")
End Function

Public Function NewVisio() As Object
    Set NewVisio = CreateObject("Visio.Application")
End Function

Public Function NewAdobeAcrobat() As Object
    Set NewAdobeAcrobat = CreateObject("AcroExch.App")
End Function
