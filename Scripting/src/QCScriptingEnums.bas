Attribute VB_Name = "QCScriptingEnums"
'@Folder "ipm-modules.QC.Scripting.src.Common"
Option Explicit

Public Enum qcIOMode
    qcForReading = 1
    qcForWriting = 2
    qcForAppend = 8
End Enum

Public Enum qcFormat
    qcDefault = -2
    qcUnicode
    qcASCII
End Enum

Public Enum qcStandardStreamTypes
    qcStdIn
    qcStdOut
    qcStdErr
End Enum
