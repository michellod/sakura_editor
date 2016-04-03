Option Explicit

Class Sys_WindowsManagementInstrumentation
  '=== ATTRIBUTE ===
  Private MSTR_CLASSNAME
  
  Private mobj_Locator
  Private mobj_Service
  Private mstr_BIOS_Description
  Private mstr_BIOS_Manufacturer
  Private mstr_BIOS_SerialNumber
  Private mstr_BIOS_SMBIOSBIOSVersion
  Private mstr_CPU_Description
  Private mstr_CPU_Name              
  Private mstr_CPU_Manufacturer      
  Private mstr_CPU_CurrentClockSpeed 
  Private mstr_CPU_MaxClockSpeed     
  Private mstr_CPU_L2CacheSize       
  Private mstr_MB_Manufacturer
  Private mstr_MB_Product
  Private mstr_MB_Version
  Private mstr_PageFileName        
  Private mstr_PageFileInitialSize 
  Private mstr_PageFileMaximumSize 
  Private mstr_COMSYS_UserName      
  Private mlng_COMSYS_DomainRoleCode
  Private mstr_COMSYS_DomainRole
  Private mlng_COMSYS_TotalPhysicalMemory
  Private mstr_WPA_ActivationRequired
  Private mstr_WPA_ProductID
  Private mlng_PhysicalMemoriesNumber
  Private mstr_PM_BankLabels()
  Private mstr_PM_Capacities()
  Private mstr_OS_ServicePackMajorVersion
  Private mstr_OS_CSName
  Private mstr_OS_Description
  Private mstr_OS_BootDevice              
  Private mstr_OS_Caption                 
  Private mstr_OS_Version                 
  Private mstr_OS_BuildNumber
  Private mstr_OS_OSArchitecture
  Private mstr_OS_InstallDate
  Private mstr_OS_LastBootUpTime
  Private mlng_CacheMemoriesNumber
  Private mstr_CM_Purposes()
  Private mstr_CM_InstalledSizes()
  Private mlng_NetworkLoginProfilesNumber
  Private mstr_NLP_Names         ()
  Private mstr_NLP_LastLogons    ()
  Private mstr_NLP_NumberOfLogons()
  Private mstr_CSP_Name
  Private mstr_CSP_IdentifyingNumber
  Private mstr_CSP_SKUNumber
  Private mstr_CSP_Vendor
  Private mstr_CSP_Version
  Private mlng_OnBoardDeviceNumber
  Private mstr_OBD_Descriptions()
  
  '=== PROPERTY ===
  Public Property Get CLASSNAME                     (): CLASSNAME                      = "WindowsManagementInstrument": End Property
  Public Property Get BIOS_Description              (): BIOS_Description               = mstr_BIOS_Description        : End Property
  Public Property Get BIOS_Manufacturer             (): BIOS_Manufacturer              = mstr_BIOS_Manufacturer       : End Property
  Public Property Get BIOS_SerialNumber             (): BIOS_SerialNumber              = mstr_BIOS_SerialNumber       : End Property
  Public Property Get BIOS_SMBIOSBIOSVersion        (): BIOS_SMBIOSBIOSVersion         = mstr_BIOS_SMBIOSBIOSVersion  : End Property
  Public Property Get CPU_Description               (): CPU_Description                = mstr_CPU_Description         : End Property
  Public Property Get CPU_Name                      (): CPU_Name                       = mstr_CPU_Name                : End Property
  Public Property Get CPU_Manufacturer              (): CPU_Manufacturer               = mstr_CPU_Manufacturer        : End Property
  Public Property Get CPU_CurrentClockSpeed         (): CPU_CurrentClockSpeed          = mstr_CPU_CurrentClockSpeed   : End Property
  Public Property Get CPU_MaxClockSpeed             (): CPU_MaxClockSpeed              = mstr_CPU_MaxClockSpeed       : End Property
  Public Property Get CPU_L2CacheSize               (): CPU_L2CacheSize                = mstr_CPU_L2CacheSize         : End Property
  Public Property Get MB_Manufacturer               (): MB_Manufacturer                = mstr_MB_Manufacturer         : End Property
  Public Property Get MB_Product                    (): MB_Product                     = mstr_MB_Product              : End Property
  Public Property Get MB_Version                    (): MB_Version                     = mstr_MB_Version              : End Property
  Public Property Get PageFileName                  (): PageFileName                   = mstr_PageFileName            : End Property
  Public Property Get PageFileInitialSize           (): PageFileInitialSize            = mstr_PageFileInitialSize     : End Property
  Public Property Get PageFileMaximumSize           (): PageFileMaximumSize            = mstr_PageFileMaximumSize     : End Property
  Public Property Get COMSYS_UserName               (): COMSYS_UserName                = mstr_COMSYS_UserName            : End Property
  Public Property Get COMSYS_DomainRoleCode         (): COMSYS_DomainRoleCode          = mlng_COMSYS_DomainRoleCode      : End Property 
  Public Property Get COMSYS_DomainRole             (): COMSYS_DomainRole              = mstr_COMSYS_DomainRole          : End Property 
  Public Property Get COMSYS_TotalPhysicalMemory    (): COMSYS_TotalPhysicalMemory     = mlng_COMSYS_TotalPhysicalMemory : End Property 
  Public Property Get WPA_ActivationRequired        (): WPA_ActivationRequired         = mstr_WPA_ActivationRequired     : End Property 
  Public Property Get WPA_ProductID                 (): WPA_ProductID                  = mstr_WPA_ProductID              : End Property 
  Public Property Get PhysicalMemoriesNumber        (): PhysicalMemoriesNumber         = mlng_PhysicalMemoriesNumber     : End Property 
  Public Property Get CacheMemoriesNumber           (): CacheMemoriesNumber            = mlng_CacheMemoriesNumber        : End Property 
  Public Property Get OS_ServicePackMajorVersion    (): OS_ServicePackMajorVersion     = mstr_OS_ServicePackMajorVersion : End Property 
  Public Property Get OS_CSName                     (): OS_CSName                      = mstr_OS_CSName                  : End Property 
  Public Property Get OS_Description                (): OS_Description                 = mstr_OS_Description             : End Property 
  Public Property Get OS_BootDevice                 (): OS_BootDevice                  = mstr_OS_BootDevice              : End Property 
  Public Property Get OS_Caption                    (): OS_Caption                     = mstr_OS_Caption                 : End Property 
  Public Property Get OS_Version                    (): OS_Version                     = mstr_OS_Version                 : End Property 
  Public Property Get OS_BuildNumber                (): OS_BuildNumber                 = mstr_OS_BuildNumber             : End Property 
  Public Property Get OS_OSArchitecture             (): OS_OSArchitecture              = mstr_OS_OSArchitecture          : End Property 
  Public Property Get OS_InstallDate                (): OS_InstallDate                 = mstr_OS_InstallDate             : End Property 
  Public Property Get OS_LastBootUpTime             (): OS_LastBootUpTime              = mstr_OS_LastBootUpTime          : End Property 
  Public Property Get NetworkLoginProfilesNumber    (): NetworkLoginProfilesNumber     = mlng_NetworkLoginProfilesNumber : End Property 
  Public Property Get CSP_Name                      (): CSP_Name                       = mstr_CSP_Name                   : End Property
  Public Property Get CSP_IdentifyingNumber         (): CSP_IdentifyingNumber          = mstr_CSP_IdentifyingNumber      : End Property
  Public Property Get CSP_SKUNumber                 (): CSP_SKUNumber                  = mstr_CSP_SKUNumber              : End Property
  Public Property Get CSP_Vendor                    (): CSP_Vendor                     = mstr_CSP_Vendor                 : End Property
  Public Property Get CSP_Version                   (): CSP_Version                    = mstr_CSP_Version                : End Property
  Public Property Get OnBoardDeviceNumber           (): OnBoardDeviceNumber            = mlng_OnBoardDeviceNumber        : End Property
  
  Public Function PM_BankLabels     (ByVal ilng_Index): PM_BankLabels      = mstr_PM_BankLabels    (ilng_Index): End Function
  Public Function PM_Capacities     (ByVal ilng_Index): PM_Capacities      = mstr_PM_Capacities    (ilng_Index): End Function
  Public Function CM_Purposes       (ByVal ilng_Index): CM_Purposes        = mstr_CM_Purposes      (ilng_Index): End Function
  Public Function CM_InstalledSizes (ByVal ilng_Index): CM_InstalledSizes  = mstr_CM_InstalledSizes(ilng_Index): End Function
  Public Function OBD_Descriptions  (ByVal ilng_Index): OBD_Descriptions   = mstr_OBD_Descriptions (ilng_Index): End Function
  Public Function NLP_Names         (ByVal ilng_Index): NLP_Names          = mstr_NLP_Names         (ilng_Index): End Function
  Public Function NLP_LastLogons    (ByVal ilng_Index): NLP_LastLogons     = mstr_NLP_LastLogons    (ilng_Index): End Function
  Public Function NLP_NumberOfLogons(ByVal ilng_Index): NLP_NumberOfLogons = mstr_NLP_NumberOfLogons(ilng_Index): End Function
  
  '=== PROCEDURE ===
  Public Sub Class_Initialize()
    Dim lobj_ClassSet
    Dim lobj_Class
    
    Set mobj_Locator  = CreateObject("WbemScripting.SWbemLocator")
    Set mobj_Service  = mobj_Locator.ConnectServer
    
    'BIOS
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_BIOS")
    For Each lobj_Class In lobj_ClassSet
      mstr_BIOS_Description       = lobj_Class.Description      
      mstr_BIOS_Manufacturer      = lobj_Class.Manufacturer     
      mstr_BIOS_SerialNumber      = lobj_Class.SerialNumber     
      mstr_BIOS_SMBIOSBIOSVersion = lobj_Class.SMBIOSBIOSVersion
    Next
    
    'CPU
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_Processor")
    For Each lobj_Class In lobj_ClassSet
      mstr_CPU_Description       = lobj_Class.Description
      mstr_CPU_Name              = lobj_Class.Name             
      mstr_CPU_Manufacturer      = lobj_Class.Manufacturer     
      mstr_CPU_CurrentClockSpeed = lobj_Class.CurrentClockSpeed
      mstr_CPU_MaxClockSpeed     = lobj_Class.MaxClockSpeed    
      mstr_CPU_L2CacheSize       = lobj_Class.L2CacheSize      
    Next
    
    'Motherboard
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_BaseBoard")
    For Each lobj_Class In lobj_ClassSet
      mstr_MB_Manufacturer = lobj_Class.Manufacturer
      mstr_MB_Product      = lobj_Class.Product
      mstr_MB_Version      = lobj_Class.Version
    Next
    
    'Page File
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_PageFileSetting")
    For Each lobj_Class In lobj_ClassSet
      mstr_PageFileName        = lobj_Class.FileName       
      mstr_PageFileInitialSize = CStr(lobj_Class.FileInitialSize)
      mstr_PageFileMaximumSize = CStr(lobj_Class.FileMaximumSize)
    Next
    
    'Logon User
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_ComputerSystem")
    For Each lobj_Class In lobj_ClassSet
      mstr_COMSYS_UserName            = lobj_Class.UserName
      mlng_COMSYS_DomainRoleCode      = lobj_Class.DomainRole
      mlng_COMSYS_TotalPhysicalMemory = lobj_Class.TotalPhysicalMemory
    Next
    Select Case mlng_COMSYS_DomainRoleCode
      Case 0: mstr_COMSYS_DomainRole = "Standalone Workstation"
      Case 1: mstr_COMSYS_DomainRole = "Member Workstation"
      Case 2: mstr_COMSYS_DomainRole = "Standalone Server"
      Case 3: mstr_COMSYS_DomainRole = "Member Server"
      Case 4: mstr_COMSYS_DomainRole = "Backup Domain Controller"
      Case 5: mstr_COMSYS_DomainRole = "Primary Domain Controller"
    End Select
    
    'Windows Product Activation
'    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_WindowsProductActivation")
'    For Each lobj_Class In lobj_ClassSet
'      mstr_WPA_ActivationRequired = lobj_Class.ActivationRequired
'      mstr_WPA_ProductID          = lobj_Class.ProductID         
'    Next
    
    'Physical Memory
    mlng_PhysicalMemoriesNumber = 0
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_PhysicalMemory")
    For Each lobj_Class In lobj_ClassSet
      ReDim Preserve mstr_PM_BankLabels(mlng_PhysicalMemoriesNumber): mstr_PM_BankLabels(mlng_PhysicalMemoriesNumber) = lobj_Class.BankLabel
      ReDim Preserve mstr_PM_Capacities(mlng_PhysicalMemoriesNumber): mstr_PM_Capacities(mlng_PhysicalMemoriesNumber) = lobj_Class.Capacity
      mlng_PhysicalMemoriesNumber = mlng_PhysicalMemoriesNumber + 1
    Next
    
    'Operating System
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_OperatingSystem")
    For Each lobj_Class In lobj_ClassSet
      mstr_OS_ServicePackMajorVersion = lobj_Class.ServicePackMajorVersion
      mstr_OS_CSName                  = lobj_Class.CSName     
      mstr_OS_Description             = lobj_Class.Description
      mstr_OS_BootDevice              = lobj_Class.BootDevice
      mstr_OS_Caption                 = lobj_Class.Caption   
      mstr_OS_Version                 = lobj_Class.Version   
      mstr_OS_BuildNumber             = lobj_Class.BuildNumber
      mstr_OS_OSArchitecture          = lobj_Class.OSArchitecture
      mstr_OS_InstallDate             = lobj_Class.InstallDate
      mstr_OS_LastBootUpTime          = lobj_Class.LastBootUpTime
    Next
    
    'Cache Memory 
    mlng_CacheMemoriesNumber = 0
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_CacheMemory")
    For Each lobj_Class In lobj_ClassSet
      ReDim Preserve mstr_CM_Purposes(mlng_CacheMemoriesNumber)      : mstr_CM_Purposes(mlng_CacheMemoriesNumber)       = lobj_Class.Purpose
      ReDim Preserve mstr_CM_InstalledSizes(mlng_CacheMemoriesNumber): mstr_CM_InstalledSizes(mlng_CacheMemoriesNumber) = lobj_Class.InstalledSize
      mlng_CacheMemoriesNumber = mlng_CacheMemoriesNumber + 1
    Next
    
    'Network Login Profile
    mlng_NetworkLoginProfilesNumber = 0
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_NetworkLoginProfile")
    For Each lobj_Class In lobj_ClassSet
      ReDim Preserve mstr_NLP_Names         (mlng_NetworkLoginProfilesNumber): mstr_NLP_Names         (mlng_NetworkLoginProfilesNumber) = lobj_Class.Name         
      ReDim Preserve mstr_NLP_LastLogons    (mlng_NetworkLoginProfilesNumber): mstr_NLP_LastLogons    (mlng_NetworkLoginProfilesNumber) = lobj_Class.LastLogon    
      ReDim Preserve mstr_NLP_NumberOfLogons(mlng_NetworkLoginProfilesNumber): mstr_NLP_NumberOfLogons(mlng_NetworkLoginProfilesNumber) = lobj_Class.NumberOfLogons
      mlng_NetworkLoginProfilesNumber = mlng_NetworkLoginProfilesNumber + 1
    Next
    
    'Computer System Product
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_ComputerSystemProduct")
    For Each lobj_Class In lobj_ClassSet
      mstr_CSP_Name              = lobj_Class.Name             
      mstr_CSP_IdentifyingNumber = lobj_Class.IdentifyingNumber
      mstr_CSP_SKUNumber         = lobj_Class.SKUNumber        
      mstr_CSP_Vendor            = lobj_Class.Vendor           
      mstr_CSP_Version           = lobj_Class.Version          
    Next
    
    'On Board Device
    mlng_OnBoardDeviceNumber = 0
    Set lobj_ClassSet = mobj_Service.ExecQuery("Select * From Win32_OnBoardDevice")
    For Each lobj_Class In lobj_ClassSet
      ReDim Preserve mstr_OBD_Descriptions(mlng_OnBoardDeviceNumber): mstr_OBD_Descriptions(mlng_OnBoardDeviceNumber) = lobj_Class.Description
      mlng_OnBoardDeviceNumber = mlng_OnBoardDeviceNumber + 1
    Next
    
    Set lobj_ClassSet = Nothing
  End Sub
  
  Public Sub Class_Terminate()
    Set mobj_Service  = Nothing
    Set mobj_Locator  = Nothing
  End Sub
  
  Public Function GetBootUpDate(ByVal ilng_Days)
    GetBootUpDate = ""
    
    Dim lobj_ELStartDate
    Dim lobj_ELEndDate
    Dim lobj_EventDate
    Dim ldat_Today
    Dim lobj_WMIService
    Dim lobj_ColumnEvents
    Dim lobj_Event
    
    Set lobj_ELStartDate = CreateObject("WbemScripting.SWbemDateTime")
    Set lobj_ELEndDate   = CreateObject("WbemScripting.SWbemDateTime")
    Set lobj_EventDate   = CreateObject("WbemScripting.SWbemDateTime")
    ldat_Today           = Date
    
    Call lobj_ELStartDate.SetVarDate(ldat_Today - ilng_Days, True)
    Call lobj_ELEndDate.SetVarDate  (ldat_Today - ilng_Days + 1, True)
    
    Set lobj_WMIService   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set lobj_ColumnEvents = lobj_WMIService.ExecQuery("Select * from Win32_NTLogEvent Where TimeWritten >= '" & lobj_ELStartDate & "' and TimeWritten < '" & lobj_ELEndDate & "' AND EventCode = 6005")
    
    For Each lobj_Event In lobj_ColumnEvents
      lobj_EventDate.Value = lobj_Event.TimeWritten
      GetBootUpDate = lobj_EventDate.GetVarDate(True)
    Next
    
    Set lobj_ColumnEvents = Nothing
    Set lobj_WMIService   = Nothing
    Set lobj_ELStartDate = Nothing
    Set lobj_ELEndDate   = Nothing
  End Function
  
  Public Function GetShutdownDate(ByVal ilng_Days)
    GetShutdownDate = ""
    
    Dim lobj_ELStartDate
    Dim lobj_ELEndDate
    Dim lobj_EventDate
    Dim ldat_Today
    Dim lobj_WMIService
    Dim lobj_ColumnEvents
    Dim lobj_Event
    
    Set lobj_ELStartDate = CreateObject("WbemScripting.SWbemDateTime")
    Set lobj_ELEndDate   = CreateObject("WbemScripting.SWbemDateTime")
    Set lobj_EventDate   = CreateObject("WbemScripting.SWbemDateTime")
    ldat_Today           = Date
    
    Call lobj_ELStartDate.SetVarDate(ldat_Today - ilng_Days, True)
    Call lobj_ELEndDate.SetVarDate  (ldat_Today - ilng_Days + 1, True)
    
    Set lobj_WMIService   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set lobj_ColumnEvents = lobj_WMIService.ExecQuery("Select * from Win32_NTLogEvent Where TimeWritten >= '" & lobj_ELStartDate & "' and TimeWritten < '" & lobj_ELEndDate & "' AND EventCode = 6006")
    
    For Each lobj_Event In lobj_ColumnEvents
      lobj_EventDate.Value = lobj_Event.TimeWritten
      GetShutdownDate = lobj_EventDate.GetVarDate(True)
    Next
    
    Set lobj_ColumnEvents = Nothing
    Set lobj_WMIService   = Nothing
    Set lobj_ELStartDate = Nothing
    Set lobj_ELEndDate   = Nothing
  End Function
  
  Public Function GetRecentBootUpDate()
    GetRecentBootUpDate = ""
    
    Dim lobj_ELStartDate
    Dim lobj_ELEndDate
    Dim lobj_EventDate
    Dim ldat_Today
    Dim lobj_WMIService
    Dim lobj_ColumnEvents
    Dim lobj_Event
    Dim llng_Day
    
    Set lobj_ELStartDate = CreateObject("WbemScripting.SWbemDateTime")
    Set lobj_ELEndDate   = CreateObject("WbemScripting.SWbemDateTime")
    Set lobj_EventDate   = CreateObject("WbemScripting.SWbemDateTime")
    Set lobj_WMIService   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    ldat_Today           = Date
    
    llng_Day = 1
    Do While GetRecentBootUpDate = ""
      Call lobj_ELStartDate.SetVarDate(ldat_Today - llng_Day, True)
      Call lobj_ELEndDate.SetVarDate  (ldat_Today - llng_Day + 1, True)
      
      Set lobj_ColumnEvents = lobj_WMIService.ExecQuery("Select * from Win32_NTLogEvent Where TimeWritten >= '" & lobj_ELStartDate & "' and TimeWritten < '" & lobj_ELEndDate & "' AND EventCode = 6005")
      
      For Each lobj_Event In lobj_ColumnEvents
        lobj_EventDate.Value  = lobj_Event.TimeWritten
        GetRecentBootUpDate = lobj_EventDate.GetVarDate(True)
      Next
      
      llng_Day = llng_Day + 1
    Loop
    
    Set lobj_ColumnEvents = Nothing
    Set lobj_WMIService   = Nothing
    Set lobj_ELStartDate  = Nothing
    Set lobj_ELEndDate    = Nothing
  End Function
  
  Public Function GetRecentShutdownDate()
    GetRecentShutdownDate = ""
    
    Dim lobj_ELStartDate
    Dim lobj_ELEndDate
    Dim lobj_EventDate
    Dim ldat_Today
    Dim lobj_WMIService
    Dim lobj_ColumnEvents
    Dim lobj_Event
    Dim llng_Day
    
    Set lobj_ELStartDate = CreateObject("WbemScripting.SWbemDateTime")
    Set lobj_ELEndDate   = CreateObject("WbemScripting.SWbemDateTime")
    Set lobj_EventDate   = CreateObject("WbemScripting.SWbemDateTime")
    Set lobj_WMIService   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    ldat_Today           = Date
    
    llng_Day = 1
    Do While GetRecentShutdownDate = ""
      Call lobj_ELStartDate.SetVarDate(ldat_Today - llng_Day, True)
      Call lobj_ELEndDate.SetVarDate  (ldat_Today - llng_Day + 1, True)
      
      Set lobj_ColumnEvents = lobj_WMIService.ExecQuery("Select * from Win32_NTLogEvent Where TimeWritten >= '" & lobj_ELStartDate & "' and TimeWritten < '" & lobj_ELEndDate & "' AND EventCode = 6006")
      
      For Each lobj_Event In lobj_ColumnEvents
        lobj_EventDate.Value  = lobj_Event.TimeWritten
        GetRecentShutdownDate = lobj_EventDate.GetVarDate(True)
      Next
      
      llng_Day = llng_Day + 1
    Loop
    
    Set lobj_ColumnEvents = Nothing
    Set lobj_WMIService   = Nothing
    Set lobj_ELStartDate = Nothing
    Set lobj_ELEndDate   = Nothing
  End Function
End Class
