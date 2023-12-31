Attribute VB_Name = "QCScripting"
'@Folder "ipm-modules.QC.Scripting.src"
Option Explicit

Public Function NewDictionary(Optional ByVal CompareMode As VbCompareMethod = vbBinaryCompare) As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
    With NewDictionary
        .CompareMode = CompareMode
    End With
End Function

Public Function NewFileSystemObject() As Object
    Set NewFileSystemObject = CreateObject("Scripting.FileSystemObject")
End Function

Public Function NewDrive(ByVal Path As String) As Object
    Dim DriveSpec As String: DriveSpec = Strings.Split(Path, Application.PathSeparator)(0)
    Set NewDrive = NewFileSystemObject().GetDrive(DriveSpec)
End Function

Public Function NewFolder(ByVal Path As String) As Object
    Dim FolderPath As String: FolderPath = CreateFoldersRecoursive(Path)
    Set NewFolder = NewFileSystemObject().GetFolder(FolderPath)
End Function

Public Function NewFile(ByVal Path As String) As Object
    Set NewFile = NewFileSystemObject().GetFile(Path)
End Function

Public Function NewTextStream(ByVal Path As String, Optional ByVal IOMode As qcIOMode, Optional ByVal Format As qcFormat) As Object
    Set NewTextStream = NewFileSystemObject().OpenTextFile(Path, IOMode, True, Format)
End Function

Public Function NewStandardStream(ByVal StandardStreamType As qcStandardStreamTypes, Optional ByVal Unicode As Boolean = False) As Object
    Set NewStandardStream = NewFileSystemObject().GetStandardStream(StandardStreamType, Unicode)
End Function
