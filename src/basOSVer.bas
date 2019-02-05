Attribute VB_Name = "basOSVer"
Private Type OSVERSIONINFO
     dwOSVersionInfoSize As Long
     dwMajorVersion As Long
     dwMinorVersion As Long
     dwBuildNumber As Long
     dwPlatformId As Long
     szCSDVersion As String * 128 ' // Service Pack
End Type

Private Type OSVERSIONINFOEX
     dwOSVersionInfoSize As Long
     dwMajorVersion As Long
     dwMinorVersion As Long
     dwBuildNumber As Long
     dwPlatformId As Long
     szCSDVersion As String * 128 ' // Service Pack
     wServicePackMajor As Integer
     wServicePackMinor As Integer
     wSuiteMask As Integer
     wProductType As Byte
     wReserved As Byte
End Type

Private Type SYSTEM_INFO
     ' //dwOemID As Long
     wProcessorArchitecture As Integer
     wReserved As Integer
     dwPageSize As Long
     lpMinimumApplicationAddress As Long
     lpMaximumApplicationAddress As Long
     dwActiveProcessorMask As Long
     dwNumberOfProcessors As Long
     dwProcessorType As Long
     dwAllocationGranularity As Long
     wProcessorLevel As Integer
     wProcessorRevision As Integer
End Type
' Enumerationen
Public Enum WindowsVersion
         WIN_OLD
         WIN_31
         WIN_NT_3x
         WIN_NT_4x
         WIN_95
         WIN_98
         WIN_98_ME
         WIN_2K
         WIN_CE
         WIN_XP
         WIN_2003
         WIN_2003_R2
         WIN_VISTA
         WIN_2008
         WIN_2008_R2
         WIN_7
End Enum

Public Declare Function GetTickCount Lib "kernel32" () As Long
' // API: min. Windows 95
Private Declare Function GetVersionEx1 Lib "kernel32.dll" _
                          Alias "GetVersionExA" ( _
                          ByRef lpVersionInformation As OSVERSIONINFO) As Long
                          
' // API: min. Windows 2000
Private Declare Function GetVersionEx2 Lib "kernel32.dll" _
                          Alias "GetVersionExA" ( _
                          ByRef lpVersionInformation As OSVERSIONINFOEX) As Long
                          
' // API: min. Windows XP
Private Declare Function IsWow64Process Lib "kernel32" ( _
                          ByVal hProcess As Long, _
                          ByRef Wow64Process As Long) As Long
                          
' // API: min Windows 95
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
' // API: min. Windows 2000
Private Declare Sub GetSystemInfo Lib "kernel32.dll" ( _
                     ByRef lpSystemInfo As SYSTEM_INFO)
                     
' // API: min. Windows 2003 (WOW64)
Private Declare Sub GetNativeSystemInfo Lib "kernel32.dll" ( _
                     ByRef lpSystemInfo As SYSTEM_INFO)
                     
' // API: min. Windows 2000
Private Declare Function GetSystemMetrics Lib "user32.dll" ( _
                          ByVal nIndex As Long) As Long
                          
' // API: only Windows Vista or Windows Server 2008
Private Declare Function GetProductInfo Lib "kernel32.dll" ( _
                          ByVal dwOSMajorVersion As Long, _
                          ByVal dwOSMinorVersion As Long, _
                          ByVal dwSpMajorVersion As Long, _
                          ByVal dwSpMinorVersion As Long, _
                          ByRef pdwReturnedProductType As Long) As Boolean
                          
' // Const: GetSystemMetrics
' // Windows Server 2003 R2
Private Const SM_SERVERR2 As Long = 89&
' // Windows XP Media Center Edition
Private Const SM_MEDIACENTER As Long = 87&
' // Windows XP Starter Edition
Private Const SM_STARTER As Long = 88&
' // Windows XP Tablet PC Edition
Private Const SM_TABLETPC As Long = 86&
' // Const: GetVersionEx.wProcessorArchitecture
' // x64 (AMD Or Intel)
Private Const PROCESSOR_ARCHITECTURE_AMD64 As Long = &H9&
' // Intel Itanium Processor Family (IPF)
Private Const PROCESSOR_ARCHITECTURE_IA64 As Long = &H6&
' // x86
Private Const PROCESSOR_ARCHITECTURE_INTEL As Long = &H0&
' // Unknown architecture.
Private Const PROCESSOR_ARCHITECTURE_UNKNOWN As Long = &HFFFF&
' // Const: GetVersionEx.wProductType
' // The system is a domain controller and the operating system is Windows
' Server 2008, Windows Server 2003, or Windows 2000 Server.
Private Const VER_NT_DOMAIN_CONTROLLER As Long = &H2&
' // The operating system is Windows Server 2008, Windows Server 2003, or
' Windows 2000 Server. Note that a server that is also a domain controller
' is reported as VER_NT_DOMAIN_CONTROLLER, not VER_NT_SERVER.
Private Const VER_NT_SERVER As Long = &H3&
' // The operating system is Windows Vista, Windows XP Professional,
' Windows XP Home Edition, or Windows 2000 Professional.
Private Const VER_NT_WORKSTATION As Long = &H1&
' // Const: GetVersionEx.wSuiteMask
' // Microsoft BackOffice components are installed.
Private Const VER_SUITE_BACKOFFICE As Long = &H4&
' // Windows Server 2003, Web Edition is installed.
Private Const VER_SUITE_BLADE As Long = &H400&
' // Windows Server 2003, Compute Cluster Edition is installed.
Private Const VER_SUITE_COMPUTE_SERVER As Long = &H4000&
' // Windows Server 2008 Datacenter, Windows Server 2003, Datacenter
' Edition, or Windows 2000 Datacenter Server is installed.
Private Const VER_SUITE_DATACENTER As Long = &H80&
' // Windows Server 2008 Enterprise, Windows Server 2003, Enterprise
' Edition, or Windows 2000 Advanced Server is installed. Refer to the
' Remarks section for more information about this bit flag.
Private Const VER_SUITE_ENTERPRISE As Long = &H2&
' // Windows XP Embedded is installed.
Private Const VER_SUITE_EMBEDDEDNT As Long = &H40&
' // Windows Vista Home Premium, Windows Vista Home Basic, or Windows XP
' Home Edition is installed.
Private Const VER_SUITE_PERSONAL As Long = &H200&
' // Remote Desktop is supported, but only one interactive session is
' supported. This value is set unless the system is running in application
' server mode.
Private Const VER_SUITE_SINGLEUSERTS As Long = &H100&
' // Microsoft Small Business Server was once installed on the system, but
' may have been upgraded to another version of Windows. Refer to the
' Remarks section for more information about this bit flag.
Private Const VER_SUITE_SMALLBUSINESS As Long = &H1&
' // Microsoft Small Business Server is installed with the restrictive
' client license in force. Refer to the Remarks section for more
' information about this bit flag.
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Long = &H20&
' // Windows Storage Server 2003 R2 or Windows Storage Server 2003is
' installed.
Private Const VER_SUITE_STORAGE_SERVER As Long = &H2000&
' // Terminal Services is installed. This value is always set. If
' VER_SUITE_TERMINAL is set but VER_SUITE_SINGLEUSERTS is not set, the
' system is running in application server mode.
Private Const VER_SUITE_TERMINAL As Long = &H10&
' // Windows Home Server is installed.
Private Const VER_SUITE_WH_SERVER As Long = &H8000&
' // Const: GetVersionEx.dwPlatformId
' // Specifies the Windows 3.1 OS.
Private Const VER_PLATFORM_WIN32s As Long = &H0&
' // Specifies the Windows 95 or Windows 98 OS.
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = &H1&
' // Specifies the Windows NT OS.
Private Const VER_PLATFORM_WIN32_NT As Long = &H2&
' // Specifies the Windows CE OS.
Private Const VER_PLATFORM_WIN32_CE As Long = &H3&
' // Const: GetProductInfo(pdwReturnedProductType)
' // Business Edition
Private Const PRODUCT_BUSINESS As Long = &H6&
' // Business Edition
Private Const PRODUCT_BUSINESS_N As Long = &H10&
' // Cluster Server Edition
Private Const PRODUCT_CLUSTER_SERVER As Long = &H12&
' // Server Datacenter Edition (full installation)
Private Const PRODUCT_DATACENTER_SERVER As Long = &H8&
' // Server Datacenter Edition (core installation)
Private Const PRODUCT_DATACENTER_SERVER_CORE As Long = &HC&
' // Server Datacenter Edition without Hyper-V (core installation)
Private Const PRODUCT_DATACENTER_SERVER_CORE_V As Long = &H27&
' // Server Datacenter Edition without Hyper-V (full installation)
Private Const PRODUCT_DATACENTER_SERVER_V As Long = &H25&
' // Enterprise Edition
Private Const PRODUCT_ENTERPRISE As Long = &H4&
' // Enterprise Edition
Private Const PRODUCT_ENTERPRISE_N As Long = &H1B&
' // Server Enterprise Edition (full installation)
Private Const PRODUCT_ENTERPRISE_SERVER As Long = &HA&
' // Server Enterprise Edition (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_CORE As Long = &HE&
' // Server Enterprise Edition without Hyper-V (core installation)
Private Const PRODUCT_ENTERPRISE_SERVER_CORE_V As Long = &H29&
' // Server Enterprise Edition for Itanium-based Systems
Private Const PRODUCT_ENTERPRISE_SERVER_IA64 As Long = &HF&
' // Server Enterprise Edition without Hyper-V (full installation)
Private Const PRODUCT_ENTERPRISE_SERVER_V As Long = &H26&
' // Home Basic Edition
Private Const PRODUCT_HOME_BASIC As Long = &H2&
' // Home Basic Edition
Private Const PRODUCT_HOME_BASIC_N As Long = &H5&
' // Home Premium Edition
Private Const PRODUCT_HOME_PREMIUM As Long = &H3&
' // Home Premium Edition
Private Const PRODUCT_HOME_PREMIUM_N As Long = &H1A&
' // Home Server Edition
Private Const PRODUCT_HOME_SERVER As Long = &H13&
' // Windows Essential Business Server Management Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MANAGEMENT As Long = &H1E&
' // Windows Essential Business Server Messaging Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_MESSAGING As Long = &H20&
' // Windows Essential Business Server Security Server
Private Const PRODUCT_MEDIUMBUSINESS_SERVER_SECURITY As Long = &H1F&
' // Server for Small Business Edition
Private Const PRODUCT_SERVER_FOR_SMALLBUSINESS As Long = &H18&
' // Small Business Server
Private Const PRODUCT_SMALLBUSINESS_SERVER As Long = &H9&
' // Small Business Server Premium Edition
Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM As Long = &H19&
' // Server Standard Edition (full installation)
Private Const PRODUCT_STANDARD_SERVER As Long = &H7&
' // Server Standard Edition (core installation)
Private Const PRODUCT_STANDARD_SERVER_CORE As Long = &HD&
' // Server Standard Edition without Hyper-V (core installation)
Private Const PRODUCT_STANDARD_SERVER_CORE_V As Long = &H28&
' // Server Standard Edition without Hyper-V (full installation)
Private Const PRODUCT_STANDARD_SERVER_V As Long = &H24&
' // Starter Edition
Private Const PRODUCT_STARTER As Long = &HB&
' // Storage Server Enterprise Edition
Private Const PRODUCT_STORAGE_ENTERPRISE_SERVER As Long = &H17&
' // Storage Server Express Edition
Private Const PRODUCT_STORAGE_EXPRESS_SERVER As Long = &H14&
' // Storage Server Standard Edition
Private Const PRODUCT_STORAGE_STANDARD_SERVER As Long = &H15&
' // Storage Server Workgroup Edition
Private Const PRODUCT_STORAGE_WORKGROUP_SERVER As Long = &H16&
' // An unknown product
Private Const PRODUCT_UNDEFINED As Long = &H0&
' // Ultimate Edition
Private Const PRODUCT_ULTIMATE As Long = &H1&
' // Ultimate Edition
Private Const PRODUCT_ULTIMATE_N As Long = &H1C&
' // Web Server Edition (full installation)
Private Const PRODUCT_WEB_SERVER As Long = &H11&
' // Web Server Edition (core installation)
Private Const PRODUCT_WEB_SERVER_CORE As Long = &H1D&

Public Function GetOSVersion(ByRef OSType As WindowsVersion) As String

Dim OsVersInfoEx As OSVERSIONINFOEX
Dim OsVersInfo As OSVERSIONINFO
Dim OsSystemInfo As SYSTEM_INFO
Dim Ret As Long
Dim mOSVersion As WindowsVersion
Dim sVersionName As String

OsVersInfo.dwOSVersionInfoSize = Len(OsVersInfo)
Call GetVersionEx1(OsVersInfo)

Select Case OsVersInfo.dwPlatformId
     Case VER_PLATFORM_WIN32s ' //Specifies the Windows 3.1 OS.
         mOSVersion = WIN_31
         sVersionName = "Windows 3.1"
     Case VER_PLATFORM_WIN32_WINDOWS ' // Specifies the Windows 95 or Windows 98 or Windows ME OS.
         Select Case OsVersInfo.dwMinorVersion
         Case 0
             mOSVersion = WIN_95
             OSType = WIN_95
             If OsVersInfo.dwBuildNumber = 950 Then
                 sVersionName = "Windows 95"
             ElseIf (OsVersInfo.dwBuildNumber > 950) And ( _
                 OsVersInfo.dwBuildNumber <= 1080) Then
                 sVersionName = "Windows 95 SP1"
             ElseIf OsVersInfo.dwBuildNumber > 1080 Then
                 sVersionName = "Windows 95 OSR2"
             End If
         Case 10
             mOSVersion = WIN_98
             If OsVersInfo.dwBuildNumber = 1998 Then
                 sVersionName = "Windows 98"
             ElseIf (OsVersInfo.dwBuildNumber > 1998) And (OsVersInfo.dwBuildNumber < 2183) Then
                 sVersionName = "Windows 98 SP1"
             ElseIf OsVersInfo.dwBuildNumber >= 2183 Then
                 sVersionName = "Windows 98 SE"
             End If
         Case 90
             mOSVersion = WIN_98_ME
             sVersionName = "Windows 98 ME"
         End Select
     Case VER_PLATFORM_WIN32_NT ' // Specifies the Windows NT OS.
         Select Case OsVersInfo.dwMajorVersion
         Case 3
             mOSVersion = WIN_NT_3x
             sVersionName = "Windows NT 3.x"
         Case 4
             mOSVersion = WIN_NT_4x
             sVersionName = "Windows NT 4.x"
         Case 5 ' // The operating system is Windows
                                    ' Server 2003 R2, Windows Server 2003,
                                    ' Windows XP, or Windows 2000.
             OsVersInfoEx.dwOSVersionInfoSize = Len(OsVersInfoEx)
             If GetVersionEx2(OsVersInfoEx) = 0 Then Err.Raise -1
            Select Case OsVersInfoEx.dwMinorVersion
             Case 0 ' // The operating system is Windows 2000
                 mOSVersion = WIN_2K
                 If OsVersInfoEx.wProductType = VER_NT_WORKSTATION Then
                     sVersionName = "Windows 2000 Professional"
                 Else
                     If OsVersInfoEx.wSuiteMask And VER_SUITE_DATACENTER Then
                         sVersionName = "Windows 2000 Datacenter Server"
                     ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_ENTERPRISE Then
                         sVersionName = "Windows 2000 Advanced Server"
                     Else
                         sVersionName = "Windows 2000 Server"
                     End If
                 End If
             Case 1 ' // The operating system is Windows XP.
                 mOSVersion = WIN_XP
                 OSType = WIN_XP
                 If GetSystemMetrics(SM_MEDIACENTER) Then
                     sVersionName = "Windows XP Media Center Edition"
                 ElseIf GetSystemMetrics(SM_STARTER) Then
                     sVersionName = "Windows XP Starter Edition"
                 ElseIf GetSystemMetrics(SM_TABLETPC) Then
                     sVersionName = "Windows XP Tablet PC Edition"
                 ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_PERSONAL Then
                     sVersionName = "Windows XP Home Edition"
                 Else
                     sVersionName = "Windows XP Professional"
                 End If
             Case 2 ' // The operating system is Windows
                                    ' Server 2003 R2, Windows Server 2003,
                                    ' or Windows XP Professional x64
                                    ' Edition.
                 If IsWow64Process(GetCurrentProcess(), Ret) = 0 Then Err.Raise -1
                 If Ret <> 0 Then
                     Call GetNativeSystemInfo(OsSystemInfo)
                 Else
                     Call GetSystemInfo(OsSystemInfo)
                 End If
                 
                If GetSystemMetrics(SM_SERVERR2) Then
                     mOSVersion = WIN_2003_R2
                     sVersionName = "Windows Server 2003 R2, "
                 ElseIf OsVersInfoEx.wSuiteMask = VER_SUITE_STORAGE_SERVER Then
                     mOSVersion = WIN_2003
                     sVersionName = "Windows Storage Server 2003"
                 ElseIf (OsVersInfoEx.wProductType = VER_NT_WORKSTATION) And _
                     OsSystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
                     mOSVersion = WIN_XP
                     sVersionName = "Windows XP Professional x64 Edition"
                 Else
                     mOSVersion = WIN_2003
                     sVersionName = "Windows Server 2003, "
                 End If
                 ' // Test for the server type.
                 If OsVersInfoEx.wProductType <> VER_NT_WORKSTATION Then
                     If OsSystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_IA64 Then
                         If OsVersInfoEx.wSuiteMask And VER_SUITE_DATACENTER Then
                             sVersionName = sVersionName & "Datacenter Edition for Itanium-based Systems"
                         ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_ENTERPRISE Then
                             sVersionName = sVersionName & "Enterprise Edition for Itanium-based Systems"
                         End If
                     End If
                 ElseIf OsSystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
                     If OsVersInfoEx.wSuiteMask And VER_SUITE_DATACENTER Then
                         sVersionName = sVersionName & "Datacenter x64 Edition"
                     ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_ENTERPRISE Then
                         sVersionName = sVersionName & "Enterprise x64 Edition"
                     Else
                         sVersionName = sVersionName & "Standard x64 Edition"
                     End If
                 Else
                     If OsVersInfoEx.wSuiteMask And VER_SUITE_COMPUTE_SERVER Then
                         sVersionName = sVersionName & "Compute Cluster Edition"
                     ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_DATACENTER Then
                         sVersionName = sVersionName & "Datacenter Edition"
                     ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_ENTERPRISE Then
                         sVersionName = sVersionName & "Enterprise Edition"
                     ElseIf OsVersInfoEx.wSuiteMask And VER_SUITE_BLADE Then
                         sVersionName = sVersionName & "Web Edition"
                     Else
                         sVersionName = sVersionName & "Standard Edition"
                     End If
                 End If
             End Select
         Case 6 '// The operating system is Windows Vista or _
                     Windows Server 2008 or Windows 7 or Windows _ Server 2008 R2
             OsVersInfoEx.dwOSVersionInfoSize = Len(OsVersInfoEx)
             If GetVersionEx2(OsVersInfoEx) = 0 Then Err.Raise -1
             Select Case OsVersInfoEx.dwMinorVersion
             Case 0
                 If OsVersInfoEx.wProductType = VER_NT_WORKSTATION Then
                     mOSVersion = WIN_VISTA
                     sVersionName = "Windows Vista "
                 Else
                     mOSVersion = WIN_2008
                     sVersionName = "Windows Server 2008 "
                 End If
             Case 1
                 If OsVersInfoEx.wProductType = VER_NT_WORKSTATION Then
                     mOSVersion = WIN_7
                     sVersionName = "Windows 7 "
                     OSType = WIN_XP
                 Else
                     mOSVersion = WIN_2008_R2
                     sVersionName = "Windows Server 2008 R2 "
                 End If
             End Select
             Dim dwType As Long
             Call GetProductInfo(6, 0, 0, 0, dwType)
             Select Case dwType
             Case PRODUCT_ULTIMATE
                 sVersionName = sVersionName & "Ultimate Edition"
             Case PRODUCT_HOME_PREMIUM
                 sVersionName = sVersionName & "Home Premium Edition"
             Case PRODUCT_HOME_BASIC
                 sVersionName = sVersionName & "Home Basic Edition"
             Case PRODUCT_ENTERPRISE
                 sVersionName = sVersionName & "Enterprise Edition"
             Case PRODUCT_BUSINESS
                 sVersionName = sVersionName & "Business Edition"
             Case PRODUCT_STARTER
                 sVersionName = sVersionName & "Starter Edition"
             Case PRODUCT_CLUSTER_SERVER
                 sVersionName = sVersionName & "Cluster Server Edition"
             Case PRODUCT_DATACENTER_SERVER
                 sVersionName = sVersionName & "Datacenter Edition"
             Case PRODUCT_DATACENTER_SERVER_CORE
                 sVersionName = sVersionName & "Datacenter Edition (core installation)"
             Case PRODUCT_ENTERPRISE_SERVER
                 sVersionName = sVersionName & "Enterprise Edition"
             Case PRODUCT_ENTERPRISE_SERVER_CORE
                 sVersionName = sVersionName & "Enterprise Edition (core installation)"
             Case PRODUCT_ENTERPRISE_SERVER_IA64
                 sVersionName = sVersionName & "Enterprise Edition for Itanium-based Systems"
             Case PRODUCT_SMALLBUSINESS_SERVER
                 sVersionName = sVersionName & "Small Business Server"
             Case PRODUCT_SMALLBUSINESS_SERVER_PREMIUM
                 sVersionName = sVersionName & "Small Business Server Premium Edition"
             Case PRODUCT_STANDARD_SERVER
                 sVersionName = sVersionName & "Standard Edition"
             Case PRODUCT_STANDARD_SERVER_CORE
                 sVersionName = sVersionName & "Standard Edition (core installation)"
             Case PRODUCT_WEB_SERVER
                 sVersionName = sVersionName & "Web Server Edition"
             End Select
             If IsWow64Process(GetCurrentProcess(), Ret) = 0 Then Err.Raise -1
             If Ret <> 0 Then
                 Call GetNativeSystemInfo(OsSystemInfo)
             Else
                 Call GetSystemInfo(OsSystemInfo)
             End If
             If OsSystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
                 sVersionName = sVersionName & ", 64-bit"
             ElseIf OsSystemInfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_INTEL Then
                 sVersionName = sVersionName & ", 32-bit"
             End If
         End Select
     Case VER_PLATFORM_WIN32_CE ' // Specifies the Windows CE OS.
         mOSVersion = WIN_CE
         sVersionName = "Windows CE"
     Case Else
         mOSVersion = WIN_OLD
         sVersionName = "Windows Version unknown"
     End Select
     
     GetOSVersion = sVersionName

End Function

