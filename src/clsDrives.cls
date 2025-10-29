VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[clsDrives.cls]

'
' Drives class by Dragokas
'
' Purpose: provides functionality to work with disk storage devices
'

Option Explicit

Public Enum STORAGE_TECHNOLOGY
    STORAGE_TECHNOLOGY_UNKNOWN = -1
    STORAGE_TECHNOLOGY_HDD
    STORAGE_TECHNOLOGY_SSD
End Enum

Public Enum PARTITION_STYLE
    PARTITION_STYLE_UNKNOWN = -1
    PARTITION_STYLE_MBR = 0
    PARTITION_STYLE_GPT = 1
    PARTITION_STYLE_RAW = 2
End Enum

Public Enum STORAGE_PROPERTY_ID
    StorageDeviceProperty = 0
    StorageAdapterProperty
    StorageDeviceIdProperty
    StorageDeviceUniqueIdProperty
    StorageDeviceWriteCacheProperty
    StorageMiniportProperty
    StorageAccessAlignmentProperty
    StorageDeviceSeekPenaltyProperty
    StorageDeviceTrimProperty
    StorageDeviceWriteAggregationProperty
    StorageDeviceDeviceTelemetryProperty
    StorageDeviceLBProvisioningProperty
    StorageDevicePowerProperty
    StorageDeviceCopyOffloadProperty
    StorageDeviceResiliencyProperty
    StorageDeviceMediumProductType
    StorageAdapterRpmbProperty
    StorageAdapterCryptoProperty
    StorageDeviceIoCapabilityProperty = 48
    StorageAdapterProtocolSpecificProperty
    StorageDeviceProtocolSpecificProperty
    StorageAdapterTemperatureProperty
    StorageDeviceTemperatureProperty
    StorageAdapterPhysicalTopologyProperty
    StorageDevicePhysicalTopologyProperty
    StorageDeviceAttributesProperty
    StorageDeviceManagementStatus
    StorageAdapterSerialNumberProperty
    StorageDeviceLocationProperty
    StorageDeviceNumaProperty
    StorageDeviceZonedDeviceProperty
    StorageDeviceUnsafeShutdownCount
    StorageDeviceEnduranceProperty
    StorageDeviceLedStateProperty
    StorageDeviceSelfEncryptionProperty = 64
    StorageFruIdProperty
    StorageStackProperty
    StorageAdapterProtocolSpecificPropertyEx
    StorageDeviceProtocolSpecificPropertyEx
End Enum

Public Enum STORAGE_QUERY_TYPE
    PropertyStandardQuery = 0
    PropertyExistsQuery
    PropertyMaskQuery
    PropertyQueryMaxDefined
End Enum

Private Type DISK_EXTENT
    DiskNumber As Long
    Align As Long
    StartingOffset As Currency
    ExtentLength As Currency
End Type

Private Type VOLUME_DISK_EXTENTS
    NumberOfDiskExtents As Long
    Padding             As Long
    Extents(0) As DISK_EXTENT
End Type

Private Type STORAGE_PROPERTY_QUERY
    PropertyId As STORAGE_PROPERTY_ID
    QueryType As STORAGE_QUERY_TYPE
    AdditionalParameters(0) As Byte
    Align(2) As Byte
End Type

Private Type DEVICE_SEEK_PENALTY_DESCRIPTOR
    Version As Long
    Size As Long
    IncursSeekPenalty As Byte
    Align(2) As Byte
End Type

Private Type PARTITION_INFORMATION_EX
    PartitionStyle As PARTITION_STYLE
    Align1 As Long
    StartingOffset As Currency
    PartitionLength As Currency
    PartitionNumber As Long
    RewritePartition As Byte
    IsServicePartition As Byte
    Align2 As Integer
    'union: PARTITION_INFORMATION_MBR or PARTITION_INFORMATION_GPT
    PartitionInfoSpecific(111) As Byte
End Type

Private Const IOCTL_DISK_GET_PARTITION_INFO_EX As Long = &H70048
Private Const IOCTL_STORAGE_QUERY_PROPERTY  As Long = &H2D1400
Private Const IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS  As Long = &H560000


Public Function GetPartitionStyle(sDrive As String) As PARTITION_STYLE
    On Error GoTo ErrorHandler

    GetPartitionStyle = PARTITION_STYLE_UNKNOWN

    Dim hVolume As Long
    hVolume = CreateFile(StrPtr("\\.\" & sDrive), _
                        FILE_READ_ATTRIBUTES, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, _
                        ByVal 0&, OPEN_EXISTING, 0&, 0&)
    
    If hVolume <> 0 Then
    
        Dim cbBytesReturned As Long
        Dim PartitionInfo As PARTITION_INFORMATION_EX
        
        If DeviceIoControl(hVolume, IOCTL_DISK_GET_PARTITION_INFO_EX, _
                            0&, 0&, VarPtr(PartitionInfo), LenB(PartitionInfo), _
                            cbBytesReturned, 0&) Then
            
            GetPartitionStyle = PartitionInfo.PartitionStyle
            
        Else
            Dbg "GetPartitionStyle is failed with: " & Err.LastDllError
        End If
        
        CloseHandle hVolume
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetPartitionStyle", "Drive:", sDrive
    If inIDE Then Stop: Resume Next
End Function

' Thanks to Raymond Chen
' https://devblogs.microsoft.com/oldnewthing/20201023-00/?p=104395
'
Public Function GetStorageTechnology(sDrive As String) As STORAGE_TECHNOLOGY
    On Error GoTo ErrorHandler
    
    GetStorageTechnology = STORAGE_TECHNOLOGY_UNKNOWN

    Dim hVolume As Long
    hVolume = CreateFile(StrPtr("\\.\" & sDrive), _
                        0, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, _
                        ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
    
    If hVolume = 0 Then
        Exit Function
    End If
    
    Dim cbBytesReturned As Long

    ' Get first random physical device for a volume
    Dim Extents As VOLUME_DISK_EXTENTS
    If (DeviceIoControl(hVolume, IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS, _
        0, 0, _
        VarPtr(Extents), LenB(Extents), _
        cbBytesReturned, _
         0) = 0 And Err.LastDllError <> ERROR_MORE_DATA) Then
    
        Dbg "GetStorageTechnology (Step1) is failed with: " & Err.LastDllError
        CloseHandle hVolume
        Exit Function
    End If
    
    CloseHandle hVolume
    
    Dim physicalDrivePath As String
    physicalDrivePath = "\\.\PhysicalDrive" & Extents.Extents(0).DiskNumber
    
    Dim hDevice As Long
    hDevice = CreateFile(StrPtr(physicalDrivePath), _
                        0, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, _
                        ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
    
    If hDevice = 0 Then
        Exit Function
    End If
    
    Dim result As DEVICE_SEEK_PENALTY_DESCRIPTOR
    Dim query As STORAGE_PROPERTY_QUERY
    
    query.PropertyId = StorageDeviceSeekPenaltyProperty
    query.QueryType = PropertyStandardQuery
    cbBytesReturned = 0
    
    If (DeviceIoControl(hDevice, IOCTL_STORAGE_QUERY_PROPERTY, _
        VarPtr(query), LenB(query), _
        VarPtr(result), LenB(result), _
        cbBytesReturned, 0)) Then
    
        If (result.IncursSeekPenalty) Then
            GetStorageTechnology = STORAGE_TECHNOLOGY_HDD
        Else
            GetStorageTechnology = STORAGE_TECHNOLOGY_SSD
        End If
    Else
        Dbg "GetStorageTechnology (Step2) is failed with: " & Err.LastDllError
    End If
    
    CloseHandle hDevice
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetStorageTechnology", "Drive:", sDrive
    If inIDE Then Stop: Resume Next
End Function


Public Function GetLogicalType(sDrive As String) As DRIVE_TYPE
    On Error GoTo ErrorHandler
    Dim lControlCode As Long
    Dim cbBytesReturned As Long
    Dim hVolume As Long
    
    GetLogicalType = DRIVE_UNKNOWN
    
    hVolume = CreateFile(StrPtr("\\.\" & sDrive), _
                             FILE_READ_ATTRIBUTES, _
                             FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, _
                             ByVal 0&, OPEN_EXISTING, 0&, 0&)
       
    If hVolume <> 0 Then
    
        If StrComp(sDrive, "A:", 1) = 0 Or StrComp(sDrive, "B:", 1) = 0 Then
            lControlCode = IOCTL_STORAGE_CHECK_VERIFY
        Else
            lControlCode = IOCTL_STORAGE_CHECK_VERIFY2
        End If
    
        If DeviceIoControl(hVolume, _
                    lControlCode, _
                    0&, 0&, _
                    0&, 0&, _
                    cbBytesReturned, _
                    0&) Then
            
            GetLogicalType = GetDriveType(StrPtr(sDrive))
       Else
           Dbg "GetLogicalType is failed with: " & Err.LastDllError
       End If
       
       CloseHandle hVolume
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetLogicalType", "Drive:", sDrive
    If inIDE Then Stop: Resume Next
End Function


Public Function GetFreeSpace(sDrive As String, bForCurrentUser As Boolean, Optional out_TotalSpace As Currency) As Currency ' result = Int64
    On Error GoTo ErrorHandler
    If IsProcedureAvail("GetDiskFreeSpaceExW", "kernel32.dll") Then
        If bForCurrentUser Then
            If GetDiskFreeSpaceEx(StrPtr(sDrive), VarPtr(GetFreeSpace), VarPtr(out_TotalSpace), 0&) = 0 Then Dbg "GetFreeSpace is failed with: " & Err.LastDllError
        Else
            If GetDiskFreeSpaceEx(StrPtr(sDrive), 0&, VarPtr(out_TotalSpace), VarPtr(GetFreeSpace)) = 0 Then Dbg "GetFreeSpace is failed with: " & Err.LastDllError
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFreeSpace", "Drive:", sDrive
    If inIDE Then Stop: Resume Next
End Function
