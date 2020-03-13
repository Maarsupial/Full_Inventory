
' Objetos '
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWSN = CreateObject("WScript.Network")
Set htmlFile = objFSO.CreateTextFile(objWSN.ComputerName & ".html", True)

htmlFile.writeLine("<!DOCTYPE html>" & vbNewLine)
htmlFile.writeLine("<html>")
htmlFile.writeLine("<head>")
htmlFile.writeLine("<meta charset=""UTF-8"">")
htmlFile.writeLine("<font face = ""DejaVu Sans Mono"">")
htmlFile.writeLine("<title> " & objWSN.ComputerName & "</title>")

htmlFile.writeLine("<link rel=""stylesheet"" type=""text/css"" href=""content.css"">")

htmlFile.writeLine("</head>")
htmlFile.writeLine("<body>")
htmlFile.writeLine("<center><h1><b> Inventário de Estações de Trabalho PMMC </b></h1></center>")
htmlFile.writeLine("<center><h1><b> " & objWSN.ComputerName & " </b></h1></center>")

htmlFile.writeLine("<li class=""collapsible""> Sistema Operacional </li>")
GetSistemaOperacional()
htmlFile.writeLine("<li class=""collapsible""> Processadores </li>")
GetProcessador()
htmlFile.writeLine("<li class=""collapsible""> BIOS </li>")
GetBios()
htmlFile.writeLine("<li class=""collapsible""> Memória </li>")
GetMemoria()
htmlFile.writeLine("<li class=""collapsible""> Arquivo de Paginação </li>")
GetPageFile()
htmlFile.writeLine("<li class=""collapsible""> Dispositivos de Rede </li>")
GetNetworkDevices()
htmlFile.writeLine("<li class=""collapsible""> Configurações TCP/IP </li>")
GetTcpIp()
htmlFile.writeLine("<li class=""collapsible""> Configurações de Discos </li>")
GetDisks()
htmlFile.writeLine("<li class=""collapsible""> Placas Controladoras </li>")
htmlFile.writeLine("<li class=""collapsible""> Unidade de Backup </li>")
htmlFile.writeLine("<li class=""collapsible""> Usuários </li>")
htmlFile.writeLine("<li class=""collapsible""> Softwares Instalados </li>")
htmlFile.writeLine("<li class=""collapsible""> Status dos Serviços </li>")
htmlFile.writeLine("<li class=""collapsible""> Compartilhamentos </li>")
htmlFile.writeLine("<li class=""collapsible""> Impressoras </li>")
htmlFile.writeLine("<li class=""collapsible""> Portas de Impressoras </li>")
htmlFile.writeLine("<li class=""collapsible""> Visualizador de Eventos </li>")
htmlFile.writeLine("<script type=""text/javascript"" src=""content.js""> </script>")
htmlFile.writeLine("</ol>")

' Sistema Operacional '
Function GetSistemaOperacional()
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer(".", "root\cimv2")
    objService.Security_.ImpersonationLevel = 3
    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_OperatingSystem")

    htmlFile.writeLine("<div class=""content"">")
    For Each objItem In Jobs
        installDate = GetDate(objItem.InstallDate)
        localDateTime = GetDate(objItem.LocalDateTime)
        htmlFile.writeLine("<ul>")
        htmlFile.writeLine("<li> Nome______________________: " & objItem.Caption & "</li>")
        htmlFile.writeLine("<li> Versão____________________: " & objItem.Version & "</li>")
        htmlFile.writeLine("<li> Build_____________________: " & objItem.BuildNumber & "</li>")
        htmlFile.writeLine("<li> Fabricante________________: " & objItem.Manufacturer & "</li>")
        htmlFile.writeLine("<li> Arquitetura_______________: " & objItem.OSArchitecture & "</li>")
        htmlFile.writeLine("<li> Tipo______________________: " & GetProductType(CInt(objItem.ProductType)) & "</li>")
        htmlFile.writeLine("<li> Usuário Registrado________: " & objItem.RegisteredUser & "</li>")
        htmlFile.writeLine("<li> Organização_______________: " & objItem.Organization & "</li>")
        htmlFile.writeLine("<li> Serial Key________________: " & GetSerialKey() & "</li>")
        htmlFile.writeLine("<li> Serial Number_____________: " & objItem.SerialNumber & "</li>")
        htmlFile.writeLine("<li> Status____________________: " & objItem.Status & "</li>")
        htmlFile.writeLine("<li> Fuso Horário______________: " & GetFormatedTimeZone(objItem.CurrentTimeZone) & "</li>")
        htmlFile.writeLine("<li> Data de Instalação________: " & installDate & "</li>")
        htmlFile.writeLine("<li> Data do Computador________: " & localDateTime & "</li>")
        htmlFile.writeLine("<li> Linguagem_________________: " & objItem.MUILanguages(0) & "</li>")
        htmlFile.writeLine("<li> País______________________: " & Country(CInt(objItem.CountryCode)) & "</li>")
        htmlFile.writeLine("<li> Código de Página__________: " & objItem.CodeSet & "</li>")
        htmlFile.writeLine("<li> Dispositivo de Boot_______: " & objItem.BootDevice & "</li>")
        htmlFile.writeLine("<li> Partição do Sistema_______: " & objItem.SystemDrive & "</li>")
        htmlFile.writeLine("<li> Caminho da Instalação_____: " & objItem.WindowsDirectory & "</li>")
        htmlFile.writeLine("</ul>")
    Next
    htmlFile.writeLine("</div>")
End Function

Function GetSerialKey()
    Set WshShell = CreateObject("WScript.Shell")
    GetSerialKey = ConvertToKey(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))
End Function

Function ConvertToKey(Key)
    Const KeyOffset = 52
    i = 28
    Chars = "BCDFGHJKMPQRTVWXY2346789"
    Do
        Cur = 0
        x = 14
        Do
            Cur = Cur * 256
            Cur = Key(x + KeyOffset) + Cur
            Key(x + KeyOffset) = (Cur \ 24) And 255
            Cur = Cur Mod 24
            x = x - 1
        Loop While x >= 0
        i = i - 1
        KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
        If (((29 - i) Mod 6) = 0) And (i <> -1) Then
            i = i -1
            KeyOutput = "-" & KeyOutput
        End If
    Loop While i >= 0
    ConvertToKey = KeyOutput
End Function

' Processadores '
Function GetProcessador()
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer(".", "root\cimv2")
    objService.Security_.ImpersonationLevel = 3
    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_Processor")

    htmlFile.writeLine("<div class=""content"">")
    For Each objItem In Jobs
        htmlFile.writeLine("<ul>")
        htmlFile.writeLine("<li> Nome______________________: " & objItem.Name & "</li>")
        htmlFile.writeLine("<li> Disponibilidade___________: " & GetAvailability(CInt(objItem.Availability)) & "</li>")
        htmlFile.writeLine("<li> Status de Uso_____________: " & GetCpuStatus(CInt(objItem.CpuStatus)) & "</li>")
        htmlFile.writeLine("<li> Estado Atual______________: " & objItem.Status & "</li>")
        htmlFile.writeLine("<li> Fabricante________________: " & objItem.Manufacturer & "</li>")
        htmlFile.writeLine("<li> Número de Núcleos_________: " & objItem.NumberOfCores & "</li>")
        htmlFile.writeLine("<li> Núcleos Lógicos___________: " & objItem.NumberOfLogicalProcessors & "</li>")
        htmlFile.writeLine("<li> Porcentagem de Uso________: " & objItem.LoadPercentage & " %</li>")
        htmlFile.writeLine("<li> Arquitetura_______________: " & GetArchitecture(CInt(objItem.Architecture)) & "</li>")
        htmlFile.writeLine("<li> Barramento________________: " & objItem.AddressWidth & " bits</li>")
        htmlFile.writeLine("<li> Largura de Dados__________: " & objItem.DataWidth & " bits</li>")
        htmlFile.writeLine("<li> Clock Máximo______________: " & GetClock(CInt(objItem.MaxClockSpeed)) & "</li>")
        htmlFile.writeLine("<li> Clock Atual_______________: " & GetClock(CInt(objItem.CurrentClockSpeed)) & "</li>")
        htmlFile.writeLine("<li> Clock Externo_____________: " & GetClock(CInt(objItem.ExtClock)) & "</li>")
        htmlFile.writeLine("<li> Cache L2__________________: " & CInt(objItem.L2CacheSize) / 1024 & " MB</li>")
        htmlFile.writeLine("<li> Cache L3__________________: " & CInt(objItem.L3CacheSize) / 1024 & " MB</li>")
        htmlFile.writeLine("<li> Tensão Atual______________: " & objItem.CurrentVoltage & "V</li>")
        htmlFile.writeLine("<li> ID do Dispositivo_________: " & objItem.DeviceID & "</li>")
        htmlFile.writeLine("<li> ID do Processador_________: " & objItem.ProcessorId & "</li>")
        htmlFile.writeLine("<li> Revisão / Versão__________: " & objItem.Revision & "</li>")
        htmlFile.writeLine("<li> Socket____________________: " & objItem.SocketDesignation & "</li>")
        htmlFile.writeLine("</ul>")
    Next
    htmlFile.writeLine("</div>")
End Function

' BIOS '
Function GetBIOS()
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer(".", "root\cimv2")
    objService.Security_.ImpersonationLevel = 3
    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_BIOS")

    htmlFile.writeLine("<div class=""content"">")
    For Each objItem In Jobs
        htmlFile.writeLine("<ul>")
        htmlFile.writeLine("<li> Nome______________________: " & objItem.Name & "</li>")
        htmlFile.writeLine("<li> Fabricante________________: " & objItem.Manufacturer & "</li>")
        htmlFile.writeLine("<li> Número Serial_____________: " & objItem.SerialNumber & "</li>")
        htmlFile.writeLine("<li> Versão____________________: " & objItem.Version & " - " & objItem.SMBIOSBIOSVersion &"</li>")
        htmlFile.writeLine("<li> Linguagem Atual___________: " & objItem.CurrentLanguage & "</li>")
        htmlFile.writeLine("<li> Linguagens Instaladas_____: " & objItem.InstallableLanguages & "</li>")
        htmlFile.writeLine("<li> Data de Lançamento________: " & GetDate(objItem.ReleaseDate) & "</li>")
        htmlFile.writeLine("<li> Versão SMBIOS_____________: " & objItem.SMBIOSMajorVersion & "." & objItem.SMBIOSMinorVersion & "</li>")
        htmlFile.writeLine("<li> Status____________________: " & objItem.Status & "</li>")
        htmlFile.writeLine("<li> SO Alvo___________________: " & GetTargetSO(CInt(objItem.TargetOperatingSystem)) & "</li>")
        htmlFile.writeLine("</ul>")
    Next
    htmlFile.writeLine("</div>")
End Function

' Memória '
Function GetMemoria()
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer(".", "root\cimv2")
    objService.Security_.ImpersonationLevel = 3
    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
    memoryDevices = 0
    totalMemory = 0

    htmlFile.writeLine("<div class=""content"">")
    For Each objItem In Jobs
        memoryDevices = memoryDevices + 1
        totalMemory = totalMemory + (objItem.Capacity / 1048576)
    Next

    htmlFile.writeLine("<ul>")
    htmlFile.writeLine("<li> Pentes de Memória_________: " & memoryDevices & "</li>")
    htmlFile.writeLine("<li> Capacitade Total__________: " & totalMemory & " MB</li>")
    htmlFile.writeLine("<br>")

    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_PhysicalMemory")

    For Each objItem In Jobs
        htmlFile.writeLine("<li><b>" & objItem.Tag & "</b></li>")
        htmlFile.writeLine("<ul>")
        htmlFile.writeLine("<li> Capacidade____________: " & objItem.Capacity / 1048576 & " MB</li>")
        htmlFile.writeLine("<li> Largura de Dados______: " & objItem.DataWidth & " bits</li>")
        htmlFile.writeLine("<li> Socket________________: " & objItem.DeviceLocator & "</li>")
        htmlFile.writeLine("<li> Fabricante____________: " & objItem.Manufacturer & "</li>")
        htmlFile.writeLine("<li> Tipo__________________: " & GetRAMType(CInt(objItem.MemoryType)) & "</li>")
        htmlFile.writeLine("<li> Part Number___________: " & objItem.PartNumber & "</li>")
        htmlFile.writeLine("<li> Serial Number_________: " & objItem.SerialNumber & "</li>")
        htmlFile.writeLine("<li> Velocidade____________: " & objItem.Speed & "</li>")
        htmlFile.writeLine("<li> Largura Total_________: " & objItem.TotalWidth & "</li>")
        htmlFile.writeLine("<li> Largura Total_________: " & objItem.TotalWidth & "</li>")
        htmlFile.writeLine("</ul>")
        htmlFile.writeLine("<br>")
    Next
    htmlFile.writeLine("</ul>")
    htmlFile.writeLine("</div>")
End Function

' Arquivo de Paginação '
Function GetPageFile()
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer(".", "root\cimv2")
    objService.Security_.ImpersonationLevel = 3
    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_PageFileUsage")

    htmlFile.writeLine("<div class=""content"">")
    For Each objItem In Jobs
        installDate = GetDate(objItem.InstallDate)
        htmlFile.writeLine("<ul>")
        htmlFile.writeLine("<li> Localização_______________: " & objItem.Caption & "</li>")
        htmlFile.writeLine("<li> Data de Instalação________: " & installDate & "</li>")
        htmlFile.writeLine("<li> Espaço Padrão_____________: " & objItem.AllocatedBaseSize & " MB</li>")
        htmlFile.writeLine("<li> Uso Atual_________________: " & objItem.CurrentUsage & " MB</li>")
        htmlFile.writeLine("<li> Pico de Uso_______________: " & objItem.PeakUsage & " MB</li>")
        htmlFile.writeLine("</ul>")
    Next

    htmlFile.writeLine("</div>")
End Function

' Dispositivos de Rede '
Function GetNetworkDevices()
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer(".", "root\cimv2")
    objService.Security_.ImpersonationLevel = 3
    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_NetworkAdapter")
    networkDevices = 0

    htmlFile.writeLine("<div class=""content"">")
    For Each objItem In Jobs
        networkDevices = networkDevices + 1
    Next

    htmlFile.writeLine("<ul>")
    htmlFile.writeLine("<li> Dispositivos de Rede______: " & networkDevices & "</li>")
    htmlFile.writeLine("<br>")

    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_NetworkAdapter")

    For Each objItem In Jobs
        lastResetDate = GetDate(objItem.TimeOfLastReset)
        htmlFile.writeLine("<li><b>" & objItem.Description & "</b></li>")
        htmlFile.writeLine("<ul>")
        htmlFile.writeLine("<li> Tipo do Adaptador_____: " & objItem.AdapterType & "</li>")
        htmlFile.writeLine("<li> Disponibilidade_______: " & GetNetDeviceAvailability(objItem.Availability) & "</li>")
        htmlFile.writeLine("<li> Estado do Dispositivo_: " & GetConfigManagerErrorCode(objItem.ConfigManagerErrorCode) & "</li>")
        htmlFile.writeLine("<li> GUID__________________: " & objItem.GUID & "</li>")
        htmlFile.writeLine("<li> Instalado_____________: " & objItem.Installed & "</li>")
        htmlFile.writeLine("<li> Endereço MAC__________: " & objItem.MACAddress & "</li>")
        htmlFile.writeLine("<li> Fabricante____________: " & objItem.Manufacturer & "</li>")
        htmlFile.writeLine("<li> ID da Conexão_________: " & objItem.NetConnectionID & "</li>")
        htmlFile.writeLine("<li> Estado da Conexão_____: " & GetNetConnectionStatus(objItem.NetConnectionStatus) & "</li>")
        htmlFile.writeLine("<li> Conexão Ativada_______: " & objItem.NetEnabled & "</li>")
        htmlFile.writeLine("<li> Adaptador Físico______: " & objItem.PhysicalAdapter & "</li>")
        htmlFile.writeLine("<li> Nome do Serviço_______: " & objItem.ServiceName & "</li>")
        htmlFile.writeLine("<li> Velocidade da Conexão_: " & GetNetDeviceSpeed(objItem.Speed) & "</li>")
        htmlFile.writeLine("<li> Ultimo Reset__________: " & lastResetDate & "</li>")
        htmlFile.writeLine("</ul>")
        htmlFile.writeLine("<br>")
    Next
    htmlFile.writeLine("</ul>")
    htmlFile.writeLine("</div>")
End Function

' Configurações TCP/IP '
Function GetTcpIp()
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer(".", "root\cimv2")
    objService.Security_.ImpersonationLevel = 3
    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

    htmlFile.writeLine("<div class=""content"">")
    For Each objItem In Jobs
        htmlFile.writeLine("<ul>")
        htmlFile.writeLine("<li><b>" & objItem.Description & "</b></li>")
        htmlFile.writeLine("<ul>")
        If (IsCollection(objItem.DefaultIPGateway)) Then
            htmlFile.writeLine("<li> IP de Gateway_________: " & objItem.DefaultIPGateway(0) & "</li>")
        End If
        htmlFile.writeLine("<li> DHCP Ativado__________: " & objItem.DHCPEnabled & "</li>")
        If (IsCollection(objItem.DNSServerSearchOrder)) Then
            htmlFile.writeLine("<li> DNS Primário__________: " & UCase(objItem.DNSServerSearchOrder(0)) & "</li>")
            htmlFile.writeLine("<li> DNS Secundário________: " & UCase(objItem.DNSServerSearchOrder(1)) & "</li>")
        End If
        PrintIpv4v6(objItem.IPAddress)
        htmlFile.writeLine("<li> Endereço MAC__________: " & objItem.MACAddress & "</li>")
        htmlFile.writeLine("<li> Máscara de Subrede____: " & objItem.IPSubnet(0) & "</li>")
        htmlFile.writeLine("<li> Nome do Serviço_______: " & objItem.ServiceName & "</li>")
        htmlFile.writeLine("</ul>")
        htmlFile.writeLine("<br>")
        htmlFile.writeLine("</ul>")
    Next
    htmlFile.writeLine("</div>")
End Function

Function PrintIpv4v6(Jobs)
    On Error Resume Next
    For Each objItem In Jobs
        If InStr(objItem, ":") <> 0 Then
            htmlFile.writeLine("<li> Endereço IPv6_________: " & UCase(objItem) & "</li>")
        Else
            htmlFile.writeLine("<li> Endereço IPv4_________: " & UCase(objItem) & "</li>")
        End If
    Next
End Function

Function PrintDnsArray(Jobs)
    On Error Resume Next
    htmlFile.writeLine("<li> DNS Primário__________: " & UCase(Jobs(0)) & "</li>")
    htmlFile.writeLine("<li> DNS Secundário________: " & UCase(Jobs(1)) & "</li>")
End Function


Function GetPageFile()
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer(".", "root\cimv2")
    objService.Security_.ImpersonationLevel = 3
    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_PageFileUsage")

    htmlFile.writeLine("<div class=""content"">")
    For Each objItem In Jobs
        installDate = GetDate(objItem.InstallDate)
        htmlFile.writeLine("<ul>")
        htmlFile.writeLine("<li> Localização_______________: " & objItem.Caption & "</li>")
        htmlFile.writeLine("<li> Data de Instalação________: " & installDate & "</li>")
        htmlFile.writeLine("<li> Espaço Padrão_____________: " & objItem.AllocatedBaseSize & " MB</li>")
        htmlFile.writeLine("<li> Uso Atual_________________: " & objItem.CurrentUsage & " MB</li>")
        htmlFile.writeLine("<li> Pico de Uso_______________: " & objItem.PeakUsage & " MB</li>")
        htmlFile.writeLine("</ul>")
    Next

    htmlFile.writeLine("</div>")
End Function

' Configurações de Discos '
Function GetDisks()
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer(".", "root\cimv2")
    objService.Security_.ImpersonationLevel = 3
    Set Jobs = objService.ExecQuery("SELECT * FROM Win32_LogicalDisk")

    htmlFile.writeLine("<div class=""content"">")
    For Each objItem In Jobs
        htmlFile.writeLine("<li> BlockSize: " & objItem.BlockSize & "</li>")
        htmlFile.writeLine("<li> Caption: " & objItem.Caption & "</li>")
        htmlFile.writeLine("<li> Compressed: " & objItem.Compressed & "</li>")
        htmlFile.writeLine("<li> ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode & "</li>")
        htmlFile.writeLine("<li> ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig & "</li>")
        htmlFile.writeLine("<li> CreationClassName: " & objItem.CreationClassName & "</li>")
        htmlFile.writeLine("<li> Description: " & objItem.Description & "</li>")
        htmlFile.writeLine("<li> DeviceID: " & objItem.DeviceID & "</li>")
        htmlFile.writeLine("<li> DriveType: " & objItem.DriveType & "</li>")
        htmlFile.writeLine("<li> ErrorCleared: " & objItem.ErrorCleared & "</li>")
        htmlFile.writeLine("<li> ErrorDescription: " & objItem.ErrorDescription & "</li>")
        htmlFile.writeLine("<li> ErrorMethodology: " & objItem.ErrorMethodology & "</li>")
        htmlFile.writeLine("<li> FileSystem: " & objItem.FileSystem & "</li>")
        htmlFile.writeLine("<li> FreeSpace: " & objItem.FreeSpace & "</li>")
        htmlFile.writeLine("<li> InstallDate: " & objItem.InstallDate & "</li>")
        htmlFile.writeLine("<li> LastErrorCode: " & objItem.LastErrorCode & "</li>")
        htmlFile.writeLine("<li> MaximumComponentLength: " & objItem.MaximumComponentLength & "</li>")
        htmlFile.writeLine("<li> MediaType: " & objItem.MediaType & "</li>")
        htmlFile.writeLine("<li> Name: " & objItem.Name & "</li>")
        htmlFile.writeLine("<li> NumberOfBlocks: " & objItem.NumberOfBlocks & "</li>")
        htmlFile.writeLine("<li> PNPDeviceID: " & objItem.PNPDeviceID & "</li>")
        'htmlFile.writeLine("<li> PowerManagementCapabilities[]: " & PrintArray(objItem.PowerManagementCapabilities) & "</li>")
        htmlFile.writeLine("<li> PowerManagementSupported: " & objItem.PowerManagementSupported & "</li>")
        htmlFile.writeLine("<li> ProviderName: " & objItem.ProviderName & "</li>")
        htmlFile.writeLine("<li> Purpose: " & objItem.Purpose & "</li>")
        htmlFile.writeLine("<li> QuotasDisabled: " & objItem.QuotasDisabled & "</li>")
        htmlFile.writeLine("<li> QuotasIncomplete: " & objItem.QuotasIncomplete & "</li>")
        htmlFile.writeLine("<li> QuotasRebuilding: " & objItem.QuotasRebuilding & "</li>")
        htmlFile.writeLine("<li> Size: " & objItem.Size & "</li>")
        htmlFile.writeLine("<li> Status: " & objItem.Status & "</li>")
        htmlFile.writeLine("<li> StatusInfo: " & objItem.StatusInfo & "</li>")
        htmlFile.writeLine("<li> SupportsDiskQuotas: " & objItem.SupportsDiskQuotas & "</li>")
        htmlFile.writeLine("<li> SupportsFileBasedCompression: " & objItem.SupportsFileBasedCompression & "</li>")
        htmlFile.writeLine("<li> SystemCreationClassName: " & objItem.SystemCreationClassName & "</li>")
        htmlFile.writeLine("<li> SystemName: " & objItem.SystemName & "</li>")
        htmlFile.writeLine("<li> VolumeDirty: " & objItem.VolumeDirty & "</li>")
        htmlFile.writeLine("<li> VolumeName: " & objItem.VolumeName & "</li>")
        htmlFile.writeLine("<li> VolumeSerialNumber: " & objItem.VolumeSerialNumber & "</li>")
    Next
    htmlFile.writeLine("</div>")

End Function

' Placas Controladoras '

' Unidade de Backup '

' Usuários '

' Softwares Instalados '

' Status dos Serviços '

' Compartilhamentos '

' Impressoras '

' Portas de Impressoras '

' Visualizador de Eventos '


htmlFile.writeLine("</body>")
htmlFile.writeLine("</html>")

' Funções '
Function PrintArray(Jobs)
    htmlFile.writeLine("<ol>")
    For Each objItem In Jobs
        htmlFile.writeLine("<li>" & objItem & "</li>")
    Next
    htmlFile.writeLine("</ol>")
End Function

Function Country( numCountry )
	' Extracted from Microsoft Internet Explorer 6 Resource Kit, '
	' Appendix F: Country/Region and Language Codes '
	' http://technet.microsoft.com/en-us/library/dd346950.aspx '

	Dim dctCountryCodes, wshShell

	Country = ""

	Set dctCountryCodes = CreateObject("Scripting.Dictionary")
	dctCountryCodes.Add    1, "United States"
	dctCountryCodes.Add    2, "Canada"
	dctCountryCodes.Add    7, "Russia"
	dctCountryCodes.Add   20, "Egypt"
	dctCountryCodes.Add   27, "South Africa"
	dctCountryCodes.Add   30, "Greece"
	dctCountryCodes.Add   31, "The Netherlands"
	dctCountryCodes.Add   32, "Belgium"
	dctCountryCodes.Add   33, "France"
	dctCountryCodes.Add   34, "Spain"
	dctCountryCodes.Add   36, "Hungary"
	dctCountryCodes.Add   39, "Italy"
	dctCountryCodes.Add   40, "Romania"
	dctCountryCodes.Add   41, "Switzerland"
	dctCountryCodes.Add   43, "Austria"
	dctCountryCodes.Add   44, "United Kingdom"
	dctCountryCodes.Add   45, "Denmark"
	dctCountryCodes.Add   46, "Sweden"
	dctCountryCodes.Add   47, "Norway"
	dctCountryCodes.Add   48, "Poland"
	dctCountryCodes.Add   49, "Germany"
	dctCountryCodes.Add   51, "Peru"
	dctCountryCodes.Add   52, "Mexico"
	dctCountryCodes.Add   53, "Cuba"
	dctCountryCodes.Add   54, "Argentina"
	dctCountryCodes.Add   55, "Brazil"
	dctCountryCodes.Add   56, "Chile"
	dctCountryCodes.Add   57, "Colombia"
	dctCountryCodes.Add   58, "Venezuela"
	dctCountryCodes.Add   60, "Malaysia"
	dctCountryCodes.Add   61, "Australia"
	dctCountryCodes.Add   62, "Indonesia"
	dctCountryCodes.Add   63, "Philippines"
	dctCountryCodes.Add   64, "New Zealand"
	dctCountryCodes.Add   65, "Singapore"
	dctCountryCodes.Add   66, "Thailand"
	dctCountryCodes.Add   81, "Japan"
	dctCountryCodes.Add   82, "Korea"
	dctCountryCodes.Add   84, "Viet Nam"
	dctCountryCodes.Add   86, "China"
	dctCountryCodes.Add   90, "Turkey"
	dctCountryCodes.Add   91, "India"
	dctCountryCodes.Add   92, "Pakistan"
	dctCountryCodes.Add   93, "Afghanistan"
	dctCountryCodes.Add   94, "Sri Lanka"
	dctCountryCodes.Add   95, "Myanmar"
	dctCountryCodes.Add   98, "Iran"
	dctCountryCodes.Add  101, "Anguilla"
	dctCountryCodes.Add  102, "Antigua and Barbuda"
	dctCountryCodes.Add  103, "The Bahamas"
	dctCountryCodes.Add  104, "Barbados"
	dctCountryCodes.Add  105, "Bermuda"
	dctCountryCodes.Add  106, "British Virgin Islands"
	dctCountryCodes.Add  108, "Cayman Islands"
	dctCountryCodes.Add  109, "Dominica"
	dctCountryCodes.Add  110, "Dominican Republic"
	dctCountryCodes.Add  111, "Grenada"
	dctCountryCodes.Add  112, "Jamaica"
	dctCountryCodes.Add  113, "Montserrat"
	dctCountryCodes.Add  115, "St. Kitts and Nevis"
	dctCountryCodes.Add  116, "St. Vincent and the Grenadines"
	dctCountryCodes.Add  117, "Trinidad and Tobago"
	dctCountryCodes.Add  118, "Turks and Caicos Islands"
	dctCountryCodes.Add  120, "Antigua and Barbuda"
	dctCountryCodes.Add  121, "Puerto Rico"
	dctCountryCodes.Add  122, "St. Lucia"
	dctCountryCodes.Add  123, "Virgin Islands"
	dctCountryCodes.Add  124, "Guam"
	dctCountryCodes.Add  212, "Morocco"
	dctCountryCodes.Add  213, "Algeria"
	dctCountryCodes.Add  216, "Tunisia"
	dctCountryCodes.Add  218, "Libya"
	dctCountryCodes.Add  220, "Gambia"
	dctCountryCodes.Add  221, "Senegal"
	dctCountryCodes.Add  222, "Mauritania"
	dctCountryCodes.Add  223, "Mali"
	dctCountryCodes.Add  224, "Guinea"
	dctCountryCodes.Add  225, "Côte d'Ivoire"
	dctCountryCodes.Add  226, "Burkina Faso"
	dctCountryCodes.Add  227, "Niger"
	dctCountryCodes.Add  228, "Togo"
	dctCountryCodes.Add  229, "Benin"
	dctCountryCodes.Add  230, "Mauritius"
	dctCountryCodes.Add  231, "Liberia"
	dctCountryCodes.Add  232, "Sierra Leone"
	dctCountryCodes.Add  233, "Ghana"
	dctCountryCodes.Add  234, "Nigeria"
	dctCountryCodes.Add  235, "Chad"
	dctCountryCodes.Add  236, "Central African Republic"
	dctCountryCodes.Add  237, "Cameroon"
	dctCountryCodes.Add  238, "Cape Verde"
	dctCountryCodes.Add  239, "São Tomé and Príncipe"
	dctCountryCodes.Add  240, "Equatorial Guinea"
	dctCountryCodes.Add  241, "Gabon"
	dctCountryCodes.Add  242, "Congo"
	dctCountryCodes.Add  243, "Congo (DRC)"
	dctCountryCodes.Add  244, "Angola"
	dctCountryCodes.Add  245, "Guinea-Bissau"
	dctCountryCodes.Add  246, "Diego Garcia"
	dctCountryCodes.Add  247, "Ascension Island"
	dctCountryCodes.Add  248, "Seychelles"
	dctCountryCodes.Add  249, "Sudan"
	dctCountryCodes.Add  250, "Rwanda"
	dctCountryCodes.Add  251, "Ethiopia"
	dctCountryCodes.Add  252, "Somalia"
	dctCountryCodes.Add  253, "Djibouti"
	dctCountryCodes.Add  254, "Kenya"
	dctCountryCodes.Add  255, "Tanzania"
	dctCountryCodes.Add  256, "Uganda"
	dctCountryCodes.Add  257, "Burundi"
	dctCountryCodes.Add  258, "Mozambique"
	dctCountryCodes.Add  260, "Zambia"
	dctCountryCodes.Add  261, "Madagascar"
	dctCountryCodes.Add  262, "Reunion"
	dctCountryCodes.Add  263, "Zimbabwe"
	dctCountryCodes.Add  264, "Namibia"
	dctCountryCodes.Add  265, "Malawi"
	dctCountryCodes.Add  266, "Lesotho"
	dctCountryCodes.Add  267, "Botswana"
	dctCountryCodes.Add  268, "Swaziland"
	dctCountryCodes.Add  269, "Mayotte"
	dctCountryCodes.Add  290, "St. Helena"
	dctCountryCodes.Add  291, "Eritrea"
	dctCountryCodes.Add  297, "Aruba"
	dctCountryCodes.Add  298, "Faroe Islands"
	dctCountryCodes.Add  299, "Greenland"
	dctCountryCodes.Add  350, "Gibraltar"
	dctCountryCodes.Add  351, "Portugal"
	dctCountryCodes.Add  352, "Luxembourg"
	dctCountryCodes.Add  353, "Ireland"
	dctCountryCodes.Add  354, "Iceland"
	dctCountryCodes.Add  355, "Albania"
	dctCountryCodes.Add  356, "Malta"
	dctCountryCodes.Add  357, "Cyprus"
	dctCountryCodes.Add  358, "Finland"
	dctCountryCodes.Add  359, "Bulgaria"
	dctCountryCodes.Add  370, "Lithuania"
	dctCountryCodes.Add  371, "Latvia"
	dctCountryCodes.Add  372, "Estonia"
	dctCountryCodes.Add  373, "Moldova"
	dctCountryCodes.Add  374, "Armenia"
	dctCountryCodes.Add  375, "Belarus"
	dctCountryCodes.Add  376, "Andorra"
	dctCountryCodes.Add  377, "Monaco"
	dctCountryCodes.Add  378, "San Marino"
	dctCountryCodes.Add  379, "Vatican City"
	dctCountryCodes.Add  380, "Ukraine"
	dctCountryCodes.Add  381, "Yugoslavia"
	dctCountryCodes.Add  385, "Croatia"
	dctCountryCodes.Add  386, "Slovenia"
	dctCountryCodes.Add  387, "Bosnia and Herzegovina"
	dctCountryCodes.Add  389, "Macedonia"
	dctCountryCodes.Add  420, "Czech Republic"
	dctCountryCodes.Add  421, "Slovakia"
	dctCountryCodes.Add  423, "Liechtenstein"
	dctCountryCodes.Add  500, "Falkland Islands (Islas Malvinas)"
	dctCountryCodes.Add  501, "Belize"
	dctCountryCodes.Add  502, "Guatemala"
	dctCountryCodes.Add  503, "El Salvador"
	dctCountryCodes.Add  504, "Honduras"
	dctCountryCodes.Add  505, "Nicaragua"
	dctCountryCodes.Add  506, "Costa Rica"
	dctCountryCodes.Add  507, "Panama"
	dctCountryCodes.Add  508, "St. Pierre and Miquelon"
	dctCountryCodes.Add  509, "Haiti"
	dctCountryCodes.Add  590, "Guadeloupe"
	dctCountryCodes.Add  591, "Bolivia"
	dctCountryCodes.Add  592, "Guyana"
	dctCountryCodes.Add  593, "Ecuador"
	dctCountryCodes.Add  594, "French Guiana"
	dctCountryCodes.Add  595, "Paraguay"
	dctCountryCodes.Add  596, "Martinique"
	dctCountryCodes.Add  597, "Suriname"
	dctCountryCodes.Add  598, "Uruguay"
	dctCountryCodes.Add  599, "Netherlands Antilles"
	dctCountryCodes.Add  670, "East Timor"
	dctCountryCodes.Add  672, "Norfolk Island"
	dctCountryCodes.Add  673, "Brunei"
	dctCountryCodes.Add  674, "Nauru"
	dctCountryCodes.Add  675, "Papua New Guinea"
	dctCountryCodes.Add  676, "Tonga"
	dctCountryCodes.Add  677, "Solomon Islands"
	dctCountryCodes.Add  678, "Vanuatu"
	dctCountryCodes.Add  679, "Fiji Islands"
	dctCountryCodes.Add  680, "Palau"
	dctCountryCodes.Add  681, "Wallis and Futuna"
	dctCountryCodes.Add  682, "Cook Islands"
	dctCountryCodes.Add  683, "Niue"
	dctCountryCodes.Add  684, "American Samoa"
	dctCountryCodes.Add  685, "Samoa"
	dctCountryCodes.Add  686, "Kiribati"
	dctCountryCodes.Add  687, "New Caledonia"
	dctCountryCodes.Add  688, "Tuvalu"
	dctCountryCodes.Add  689, "French Polynesia"
	dctCountryCodes.Add  690, "Tokelau"
	dctCountryCodes.Add  691, "Micronesia"
	dctCountryCodes.Add  692, "Marshall Islands"
	dctCountryCodes.Add  705, "Kazakhstan"
	dctCountryCodes.Add  850, "North Korea"
	dctCountryCodes.Add  852, "Hong Kong SAR"
	dctCountryCodes.Add  853, "Macau SAR"
	dctCountryCodes.Add  855, "Cambodia"
	dctCountryCodes.Add  856, "Laos"
	dctCountryCodes.Add  880, "Bangladesh"
	dctCountryCodes.Add  886, "Taiwan"
	dctCountryCodes.Add  960, "Maldives"
	dctCountryCodes.Add  961, "Lebanon"
	dctCountryCodes.Add  962, "Jordan"
	dctCountryCodes.Add  963, "Syria"
	dctCountryCodes.Add  964, "Iraq"
	dctCountryCodes.Add  965, "Kuwait"
	dctCountryCodes.Add  966, "Saudi Arabia"
	dctCountryCodes.Add  967, "Yemen"
	dctCountryCodes.Add  968, "Oman"
	dctCountryCodes.Add  971, "United Arab Emirates"
	dctCountryCodes.Add  972, "Israel"
	dctCountryCodes.Add  973, "Bahrain"
	dctCountryCodes.Add  974, "Qatar"
	dctCountryCodes.Add  975, "Bhutan"
	dctCountryCodes.Add  976, "Mongolia"
	dctCountryCodes.Add  977, "Nepal"
	dctCountryCodes.Add  992, "Tajikistan"
	dctCountryCodes.Add  993, "Turkmenistan"
	dctCountryCodes.Add  994, "Azerbaijan"
	dctCountryCodes.Add  995, "Georgia"
	dctCountryCodes.Add  996, "Kyrgyzstan"
	dctCountryCodes.Add  998, "Uzbekistan"
	dctCountryCodes.Add 2691, "Comoros"
	dctCountryCodes.Add 5399, "Guantanamo Bay"
	dctCountryCodes.Add 6101, "Cocos (Keeling) Islands"

	If numCountry < 1 Then
		Set wshShell = CreateObject("Wscript.Shell")
		numCountry = wshShell.RegRead("HKEY_CURRENT_USER\Control Panel\International\iCountry")
		Set wshShell = Nothing
		If IsNumeric(numCountry) Then numCountry = CInt(numCountry)
	End If

	Country = dctCountryCodes(numCountry)

	Set dctCountryCodes = Nothing
End Function

Function GetFormatedTimeZone(timeZone)
    timeZone = timeZone / 60
    If timeZone < 10 AND timeZone >= 0 Then
        strTimeZone = "UTC +0" & timeZone & ":00"
    ElseIf timeZone < 0 AND timeZone > -10 Then
        timeZone = timeZone * -1
        strTimeZone = "UTC -0" & timeZone & ":00"
    ElseIf timeZone >= 10 Then
        strTimeZone = "UTC +" & timeZone & ":00"
    Else
        strTimeZone = "UTC " & timeZone & ":00"
    End If
    GetFormatedTimeZone = strTimeZone
End Function

Function GetProductType(code)
    If code = 1 Then
        productType = "Estação de Trabalho"
    ElseIf code = 2 Then
        productType = "Controlador de Domínio"
    ElseIf code = 3 Then
        productType = "Servidor"
    End If

    GetProductType = productType
End Function

Function GetArchitecture(code)
    If code = 0 Then
        architecture = "x86"
    ElseIf code = 1 Then
        architecture = "MIPS"
    ElseIf code = 2 Then
        architecture = "Alpha"
    ElseIf code = 3 Then
        architecture = "PowerPC"
    ElseIf code = 5 Then
        architecture = "ARM"
    ElseIf code = 6 Then
        architecture = "ia64"
    ElseIf code = 9 Then
        architecture = "x64"
    End If

    GetArchitecture = architecture
End Function

Function GetAvailability(code)
    If code = 1 Then
        availability = "Other"
    ElseIf code = 2 Then
        availability = "Unknown"
    ElseIf code = 3 Then
        availability = "Running/Full Power"
    ElseIf code = 4 Then
        availability = "Warning"
    ElseIf code = 5 Then
        availability = "In Test"
    ElseIf code = 6 Then
        availability = "Not Applicable"
    ElseIf code = 7 Then
        availability = "Power Off"
    ElseIf code = 8 Then
        availability = "Off Line"
    ElseIf code = 9 Then
        availability = "Off Duty"
    ElseIf code = 10 Then
        availability = "Degraded"
    ElseIf code = 11 Then
        availability = "Not Installed"
    ElseIf code = 12 Then
        availability = "Install Error"
    ElseIf code = 13 Then
        availability = "Power Save - Unknown"
    ElseIf code = 14 Then
        availability = "Power Save - Low Power Mode"
    ElseIf code = 15 Then
        availability = "Power Save - Standby"
    ElseIf code = 16 Then
        availability = "Power Cycle"
    ElseIf code = 17 Then
        availability = "Power Save - Warning"
    ElseIf code = 18 Then
        availability = "Paused"
    ElseIf code = 19 Then
        availability = "Not Ready"
    ElseIf code = 20 Then
        availability = "Not Configured"
    ElseIf code = 21 Then
        availability = "Quiesced"
    End If

    GetAvailability = availability
End Function

Function GetCpuStatus(code)
    If code = 0 Then
        cpuStatus = "Unknown"
    ElseIf code = 1 Then
        cpuStatus = "CPU Enabled"
    ElseIf code = 2 Then
        cpuStatus = "CPU Disabled by User via BIOS Setup"
    ElseIf code = 3 Then
        cpuStatus = "CPU Disabled By BIOS (POST Error)"
    ElseIf code = 4 Then
        cpuStatus = "CPU is Idle"
    ElseIf code = 5 OR code = 6 Then
        cpuStatus = "Reserved"
    ElseIf code = 7 Then
        cpuStatus = "Other"
    End If

    GetCpuStatus = cpuStatus
End Function

Function GetClock(clockMhz)
    If IsNull(clockMhz) Then
        clockStr = "Unknown"
    Else
        clockGhz = clockMhz / 1000
        clockStr = clockGhz & " GHz"
        clockStr = Replace(clockStr, ",", ".")
    End If

    GetClock = clockStr
End Function

Function GetDate(intDate)
    Set objDate = CreateObject("WbemScripting.SWbemDateTime")

    objDate.Value = intDate

    GetDate = objDate.GetVarDate
End Function

Function GetTargetSO(code)
    If code = 0 Then
        targetSO = "Unknown"
    ElseIf code = 1 Then
        targetSO = "Other"
    ElseIf code = 2 Then
        targetSO = "MACOS"
    ElseIf code = 3 Then
        targetSO = "ATTUNIX"
    ElseIf code = 4 Then
        targetSO = "DGUX"
    ElseIf code = 5 Then
        targetSO = "DECNT"
    ElseIf code = 6 Then
        targetSO = "Digital Unix"
    ElseIf code = 7 Then
        targetSO = "OpenVMS"
    ElseIf code = 8 Then
        targetSO = "HPUX"
    ElseIf code = 9 Then
        targetSO = "AIX"
    ElseIf code = 10 Then
        targetSO = "MVS"
    ElseIf code = 11 Then
        targetSO = "OS400"
    ElseIf code = 12 Then
        targetSO = "OS/2"
    ElseIf code = 13 Then
        targetSO = "JavaVM"
    ElseIf code = 14 Then
        targetSO = "MSDOS"
    ElseIf code = 15 Then
        targetSO = "WIN3x"
    ElseIf code = 16 Then
        targetSO = "WIN95"
    ElseIf code = 17 Then
        targetSO = "WIN98"
    ElseIf code = 18 Then
        targetSO = "WINNT"
    ElseIf code = 19 Then
        targetSO = "WINCE"
    ElseIf code = 20 Then
        targetSO = "NCR3000"
    'Continua até 61'
    End If

    GetTargetSO = targetSO
End Function

Function GetRAMType(code)
    If code = 0 Then
        ramType = "Unknown"
    ElseIf code = 1 Then
        ramType = "Other"
    ElseIf code = 2 Then
        ramType = "DRAM"
    ElseIf code = 3 Then
        ramType = "Synchronous DRAM"
    ElseIf code = 4 Then
        ramType = "Cache DRAM"
    ElseIf code = 5 Then
        ramType = "EDO"
    ElseIf code = 6 Then
        ramType = "EDRAM"
    ElseIf code = 7 Then
        ramType = "VRAM"
    ElseIf code = 8 Then
        ramType = "SRAM"
    ElseIf code = 9 Then
        ramType = "RAM"
    ElseIf code = 10 Then
        ramType = "ROM"
    ElseIf code = 11 Then
        ramType = "Flash"
    ElseIf code = 12 Then
        ramType = "EEPROM"
    ElseIf code = 13 Then
        ramType = "FEPROM"
    ElseIf code = 14 Then
        ramType = "EPROM"
    ElseIf code = 15 Then
        ramType = "CDRAM"
    ElseIf code = 16 Then
        ramType = "3DRAM"
    ElseIf code = 17 Then
        ramType = "SDRAM"
    ElseIf code = 18 Then
        ramType = "SGRAM"
    ElseIf code = 19 Then
        ramType = "RDRAM"
    ElseIf code = 20 Then
        ramType = "DDR"
    ElseIf code = 21 Then
        ramType = "DDR2"
    ElseIf code = 22 Then
        ramType = "DDR2 FB-DIMM"
    ElseIf code = 24 Then
        ramType = "DDR3"
    ElseIf code = 25 Then
        ramType = "FBD2"
    End If

    GetRAMType = ramType
End Function

Function GetNetDeviceAvailability(code)
    If code = 1 Then
        NetDeviceAvailability = "Other"
    ElseIf code = 2 Then
        NetDeviceAvailability = "Unknown"
    ElseIf code = 3 Then
        NetDeviceAvailability = "Running/Full Power"
    ElseIf code = 4 Then
        NetDeviceAvailability = "Warning"
    ElseIf code = 5 Then
        NetDeviceAvailability = "In Test"
    ElseIf code = 6 Then
        NetDeviceAvailability = "Not Applicable"
    ElseIf code = 7 Then
        NetDeviceAvailability = "Power Off"
    ElseIf code = 8 Then
        NetDeviceAvailability = "Off Line"
    ElseIf code = 9 Then
        NetDeviceAvailability = "Off Duty"
    ElseIf code = 10 Then
        NetDeviceAvailability = "Degraded"
    ElseIf code = 11 Then
        NetDeviceAvailability = "Not Installed"
    ElseIf code = 12 Then
        NetDeviceAvailability = "Install Error"
    ElseIf code = 13 Then
        NetDeviceAvailability = "Power Save - Unknown"
    ElseIf code = 14 Then
        NetDeviceAvailability = "Power Save - Low Power Mode"
    ElseIf code = 15 Then
        NetDeviceAvailability = "Power Save - Standby"
    ElseIf code = 16 Then
        NetDeviceAvailability = "Power Cycle"
    ElseIf code = 17 Then
        NetDeviceAvailability = "Power Save - Warning"
    ElseIf code = 18 Then
        NetDeviceAvailability = "Paused"
    ElseIf code = 19 Then
        NetDeviceAvailability = "Not Ready"
    ElseIf code = 20 Then
        NetDeviceAvailability = "Not Configured"
    ElseIf code = 21 Then
        NetDeviceAvailability = "Quiesced"
    End If

    GetNetDeviceAvailability = NetDeviceAvailability
End Function

Function GetNetDeviceSpeed(bps)
    If IsNull(bps) Then
        mbps = 0
    Else
        mbps = bps / 1000000
    End If

    GetNetDeviceSpeed = (mbps & " Mbps")
End Function

Function GetConfigManagerErrorCode(code)
    If code = 0 Then
        ConfigManagerError = "This device is working properly."
    ElseIf code = 1 Then
        ConfigManagerError = "This device is not configured correctly."
    ElseIf code = 2 Then
        ConfigManagerError = "Windows cannot load the driver for this device."
    ElseIf code = 3 Then
        ConfigManagerError = "The driver for this device might be corrupted, or your system may be running low on memory or other resources."
    ElseIf code = 4 Then
        ConfigManagerError = "This device is not working properly. One of its drivers or your registry might be corrupted."
    ElseIf code = 5 Then
        ConfigManagerError = "The driver for this device needs a resource that Windows cannot manage."
    ElseIf code = 6 Then
        ConfigManagerError = "The boot configuration for this device conflicts with other devices."
    ElseIf code = 7 Then
        ConfigManagerError = "Cannot filter."
    ElseIf code = 8 Then
        ConfigManagerError = "The driver loader for the device is missing."
    ElseIf code = 9 Then
        ConfigManagerError = "This device is not working properly because the controlling firmware is reporting the resources for the device incorrectly."
    ElseIf code = 10 Then
        ConfigManagerError = "This device cannot start."
    ElseIf code = 11 Then
        ConfigManagerError = "This device failed."
    ElseIf code = 12 Then
        ConfigManagerError = "This device cannot find enough free resources that it can use."
    ElseIf code = 13 Then
        ConfigManagerError = "Windows cannot verify this device's resources."
    ElseIf code = 14 Then
        ConfigManagerError = "This device cannot work properly until you restart your computer."
    ElseIf code = 15 Then
        ConfigManagerError = "This device is not working properly because there is probably a re-enumeration problem."
    ElseIf code = 16 Then
        ConfigManagerError = "Windows cannot identify all the resources this device uses."
    ElseIf code = 17 Then
        ConfigManagerError = "This device is asking for an unknown resource type."
    ElseIf code = 18 Then
        ConfigManagerError = "Reinstall the drivers for this device."
    ElseIf code = 19 Then
        ConfigManagerError = "Failure using the VxD loader."
    ElseIf code = 20 Then
        ConfigManagerError = "Your registry might be corrupted."
    ElseIf code = 21 Then
        ConfigManagerError = "System failure: Try changing the driver for this device. If that does not work, see your hardware documentation. Windows is removing this device."
    ElseIf code = 22 Then
        ConfigManagerError = "This device is disabled."
    ElseIf code = 23 Then
        ConfigManagerError = "System failure: Try changing the driver for this device. If that doesn't work, see your hardware documentation."
    ElseIf code = 24 Then
        ConfigManagerError = "This device is not present, is not working properly, or does not have all its drivers installed."
    ElseIf code = 25 Then
        ConfigManagerError = "Windows is still setting up this device."
    ElseIf code = 26 Then
        ConfigManagerError = "Windows is still setting up this device."
    ElseIf code = 27 Then
        ConfigManagerError = "This device does not have valid log configuration."
    ElseIf code = 28 Then
        ConfigManagerError = "The drivers for this device are not installed."
    ElseIf code = 29 Then
        ConfigManagerError = "This device is disabled because the firmware of the device did not give it the required resources."
    ElseIf code = 30 Then
        ConfigManagerError = "This device is using an Interrupt Request (IRQ) resource that another device is using."
    ElseIf code = 31 Then
        ConfigManagerError = "This device is not working properly because Windows cannot load the drivers required for this device."
    End If

    GetConfigManagerErrorCode = ConfigManagerError
End Function

Function GetNetConnectionStatus(code)
    If code = 0 Then
        NetConnectionStatus = "Disconnected"
    ElseIf code = 1 Then
        NetConnectionStatus = "Connecting"
    ElseIf code = 2 Then
        NetConnectionStatus = "Connected"
    ElseIf code = 3 Then
        NetConnectionStatus = "Disconnecting"
    ElseIf code = 4 Then
        NetConnectionStatus = "Hardware Not Present"
    ElseIf code = 5 Then
        NetConnectionStatus = "Hardware Disabled"
    ElseIf code = 6 Then
        NetConnectionStatus = "Hardware Malfunction"
    ElseIf code = 7 Then
        NetConnectionStatus = "Media Disconnected"
    ElseIf code = 8 Then
        NetConnectionStatus = "Authenticating"
    ElseIf code = 9 Then
        NetConnectionStatus = "Authentication Succeeded"
    ElseIf code = 10 Then
        NetConnectionStatus = "Authentication Failed"
    ElseIf code = 11 Then
        NetConnectionStatus = "Invalid Address"
    ElseIf code = 12 Then
        NetConnectionStatus = "Credentials Required"
    Else
        NetConnectionStatus = "Other"
    End If

    GetNetConnectionStatus = NetConnectionStatus
End Function

Function IsCollection(param)
    On Error Resume Next
    For Each p In param
        Exit For
    Next
    If Err Then
        If Err.Number = 451 Then
            IsCollection = False
        Else
            WScript.Echo "Unexpected error (0x" & Hex(Err.Number) & "): " & _
                Err.Description
            WScript.Quit 1
        End If
    Else
        IsCollection = True
    End If
End Function

Function PrintIpArray(Jobs)
    On Error Resume Next
    htmlFile.writeLine("<ul>")
    For Each objItem In Jobs
        htmlFile.writeLine("<li> " & UCase(objItem) & "</li>")
    Next
    htmlFile.writeLine("</ul>")
End Function

