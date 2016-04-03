Option Explicit

'Environment
Dim SYSENV
Dim USRENV
Dim VOLENV
Dim PRCENV

'VBS: Visual Basic Script
Dim VBS

'System
Dim GLOBAL

Public Sub Sys_Init(ByRef ostr_ReturnCode,ByRef ostr_Message)
  Dim llng_Index
  
  Set SYSENV = New Sys_SystemEnvironment
  Set USRENV = New Sys_UserEnvironment
  Set VOLENV = New Sys_VolatileEnvironment
  Set PRCENV = New Sys_ProcessEnvironment
  Set VBS    = New Vbs_VBS
  Set GLOBAL = New Sys_Global
  
  LOGGER.WriteLine "<BIOS>"
  LOGGER.WriteLine "  Type          : " & GLOBAL.WMI.BIOS_Description      
  LOGGER.WriteLine "  Manufacturer  : " & GLOBAL.WMI.BIOS_Manufacturer     
  LOGGER.WriteLine "  Serial Number : " & GLOBAL.WMI.BIOS_SerialNumber     
  LOGGER.WriteLine "  Version       : " & GLOBAL.WMI.BIOS_SMBIOSBIOSVersion
  LOGGER.WriteLine "<CPU>"
  LOGGER.WriteLine "  Type                : " & GLOBAL.WMI.CPU_Description
  LOGGER.WriteLine "  Name                : " & GLOBAL.WMI.CPU_Name              
  LOGGER.WriteLine "  Manufacturer        : " & GLOBAL.WMI.CPU_Manufacturer      
  LOGGER.WriteLine "  Current Clock Speed : " & GLOBAL.WMI.CPU_CurrentClockSpeed 
  LOGGER.WriteLine "  Max Clock Speed     : " & GLOBAL.WMI.CPU_MaxClockSpeed     
  LOGGER.WriteLine "  L2 Cache Size       : " & GLOBAL.WMI.CPU_L2CacheSize       
  LOGGER.WriteLine "<Motherboard>"
  LOGGER.WriteLine "  Manufacturer : " & GLOBAL.WMI.MB_Manufacturer
  LOGGER.WriteLine "  Product      : " & GLOBAL.WMI.MB_Product
  LOGGER.WriteLine "  Version      : " & GLOBAL.WMI.MB_Version
  LOGGER.WriteLine "<Operating System>"
  LOGGER.WriteLine "  Boot Drive   : " & GLOBAL.WMI.OS_BootDevice              
  LOGGER.WriteLine "  Caption      : " & GLOBAL.WMI.OS_Caption                 
  LOGGER.WriteLine "  Version      : " & GLOBAL.WMI.OS_Version                 
  LOGGER.WriteLine "  Build Number : " & GLOBAL.WMI.OS_BuildNumber
  LOGGER.WriteLine "  Architecture : " & GLOBAL.WMI.OS_OSArchitecture
  LOGGER.WriteLine "  Install Date : " & GLOBAL.WMI.OS_InstallDate
  LOGGER.WriteLine "  Last Boot Up : " & GLOBAL.WMI.OS_LastBootUpTime
  LOGGER.WriteLine "<Paging File>"
  LOGGER.WriteLine "  Location     : " & GLOBAL.WMI.PageFileName       
  LOGGER.WriteLine "  Initial Size : " & GLOBAL.WMI.PageFileInitialSize
  LOGGER.WriteLine "  Max Size     : " & GLOBAL.WMI.PageFileMaximumSize
  LOGGER.WriteLine "<Physical Memory>"
  For llng_Index = 0 To GLOBAL.WMI.PhysicalMemoriesNumber - 1
    LOGGER.WriteLine "  " & CStr(llng_Index) & ". Bank Label : " & GLOBAL.WMI.PM_BankLabels(llng_Index)
    LOGGER.WriteLine "     Capacity   : " & GLOBAL.WMI.PM_Capacities(llng_Index)
  Next
  LOGGER.WriteLine "<Cache Memory>"
  For llng_Index = 0 To GLOBAL.WMI.CacheMemoriesNumber - 1
    LOGGER.WriteLine "  " & CStr(llng_Index) & ". Purpose        : " & GLOBAL.WMI.CM_Purposes(llng_Index)
    LOGGER.WriteLine "     Installed Size : " & GLOBAL.WMI.CM_InstalledSizes(llng_Index)
  Next
  LOGGER.WriteLine "<Computer System>"
  LOGGER.WriteLine "  User Name              : " & GLOBAL.WMI.COMSYS_UserName
  LOGGER.WriteLine "  Domain Role            : " & GLOBAL.WMI.COMSYS_DomainRole
  LOGGER.WriteLine "  Total Physical Memory  : " & GLOBAL.WMI.COMSYS_TotalPhysicalMemory
  LOGGER.WriteLine "<Computer System Product>"
  LOGGER.WriteLine "  Computer Name : " & GLOBAL.WMI.CSP_Name
  LOGGER.WriteLine "  Serial Number : " & GLOBAL.WMI.CSP_IdentifyingNumber
  LOGGER.WriteLine "  SKU Number    : " & GLOBAL.WMI.CSP_SKUNumber
  LOGGER.WriteLine "  Vendor        : " & GLOBAL.WMI.CSP_Vendor
  LOGGER.WriteLine "  Version       : " & GLOBAL.WMI.CSP_Version
  LOGGER.WriteLine "<On Board Device>"
  For llng_Index = 0 To GLOBAL.WMI.OnBoardDeviceNumber - 1
    LOGGER.WriteLine "  " & CStr(llng_Index) & ". Name : " & GLOBAL.WMI.OBD_Descriptions(llng_Index)
  Next
  LOGGER.WriteLine "<Network Login Profiles>"
  For llng_Index = 0 To GLOBAL.WMI.NetworkLoginProfilesNumber - 1
    LOGGER.WriteLine "  " & CStr(llng_Index) & ". Description      : " & GLOBAL.WMI.NLP_Names(llng_Index)
    LOGGER.WriteLine "     Last Logon       : " & GLOBAL.WMI.NLP_LastLogons    (llng_Index)
    LOGGER.WriteLine "     Number of Logons : " & GLOBAL.WMI.NLP_NumberOfLogons(llng_Index)
  Next
'  LOGGER.WriteLine "<Timestamp>"
'  LOGGER.WriteLine "  Recent BootUp   : " & GLOBAL.WMI.GetRecentBootUpDate()
'  LOGGER.WriteLine "  Recent Shutdown : " & GLOBAL.WMI.GetRecentShutdownDate()
  LOGGER.WriteLine "<Profile>"
  LOGGER.WriteLine "  Execute Date : " & GLOBAL.StartDate
  LOGGER.WriteLine "  Language     : " & GLOBAL.Language
  
  ostr_ReturnCode = "N"
  ostr_Message    = ""
End Sub
