Attribute VB_Name = "QCADODBEnums"
'@Folder "ipm-modules.QC.ADODB.src.Common"
Option Explicit

Public Enum qcCursorLocationEnum
    qcUseServer = 2
    qcUseClient
End Enum

Public Enum qcCursorTypeEnum
    qcOpenForwardOnly
    qcOpenKeyset
    qcOpenDynamic
    qcOpenStatic
End Enum

Public Enum qcLockTypeEnum
    qcLockReqcOnly = 1
    qcLockPessimistic
    qcLockOptimistic
    qcLockBatchOptimistic
End Enum

Public Enum qcMarshalOptionsEnum
    qcMarshalAll
    qcMarshalModifiedOnly
End Enum

Public Enum qcPositionEnum
    qcPosEOF = -3
    qcPosBOF
    qcPosUnknown
End Enum

Public Enum qcIsolationLevelEnum
    qcUnspecified = -1
    qcChaos = &H10
    qcBrowse = &H100
    qcReadUncommitted = &H100
    qcCursorStability = &H1000
    qcReadCommitted = &H1000
    qcRepeatableRead = &H10000
    qcIsolated = &H100000
    qcSerializable = &H100000
End Enum

Public Enum qcConnectModeEnum
    qcModeUnknown
    qcModeRead
    qcModeWrite
    qcModeReadWrite
    qcModeShareDenyRead
    qcModeShareDenyWrite = 8
    qcModeShareExclusive = 12
    qcModeShareDenyNone = &H10
    qcModeRecursive = &H400000
End Enum
