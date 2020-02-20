On Error Resume Next

Function FileExists(FilePath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
        FileExists=CBool(1)
    Else
        FileExists=CBool(0)
    End If
End Function

'Objects
Set fso = CreateObject("Scripting.FileSystemObject")
Set shl = CreateObject("WScript.Shell")

path=".\" 'path to folder
exists = fso.FolderExists(path)

Set objFSO = CreateObject("Scripting.FileSystemObject")

    Dim oFileSys
    Dim oFolder
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    Set objNetwork = CreateObject("WScript.Network")
    STRcomputer =objNetwork.ComputerName
    STRTipoServer = "MS"
    Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
    'On Error Resume Next
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    if Not(exists) then
        Set oFolder = oFileSys.CreateFolder(path)
    end if
    If FileExists(path & STRcomputer & ".HTML") Then
        Set objtextfile = objFSO.deleteFile(path & STRcomputer & ".HTML")
    End If
    Set objtextfile = objFSO.CreateTextFile(path & STRcomputer & ".HTML", Forwritting)

    Objtextfile.writeline "<html>"
    Objtextfile.writeline "<!-----Inventario de Estações----->"
    Objtextfile.writeline "<!-----SJ----->"
    Objtextfile.writeline "<!-----2019----->"
    Objtextfile.writeline "<!-----Versão 1.0----->"
    Objtextfile.writeline "<TFOOT STYLE='font-weight:bold; color:#FFFFFF'>"
    Objtextfile.writeline "<TR>"
    Objtextfile.writeline "<TD COLSPAN=5 ALIGN='center'>"
    Objtextfile.writeline "</B>"
    Objtextfile.writeline "</TR>"
    Objtextfile.writeline "</TFOOT>"

    Lin_log = 4
    Set objNetwork = CreateObject("WScript.Network")

    objtextfile.WriteLine "<head><meta charset=" & Chr(34) & "utf-8" & Chr(34) & "></head>"
    objtextfile.WriteLine "<center><a name='#menu'></center></a>"
    objtextfile.WriteLine "<br>"
    objtextfile.WriteLine "<center><H1><b> Inventario de Estações de Trabalho PMMC </b></h1></center>"
    objtextfile.WriteLine "<!WKS" & STRcomputer & "Fim_WKS>"
    objtextfile.WriteLine "<center><H1><b>" & STRcomputer & "</b></h1></center>"
   '
    Objtextfile.writeline "<Font face='arial'><ol>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#SO'>Sistema Operacional</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#proc'>Processadores</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#bios'>Bios e Hardware</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#rede'>Configurações de Rede</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#tcp'>Configurações TCP/IP</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#discos'>Configurações de Discos</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#controladores'>Placas  Controladores</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#backup'>Unidade de Backup</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#usuarios'>Usuarios Locais</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#software'>Softwares Instalados</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#servicos'>Status dos Serviços</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#compartilhamentos'>Compartilhamentos  Locais</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#impressora'>Impressoras Locais</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#portas'>Portas de Impressora</a></li></h5>"
    Objtextfile.writeline "<Font face='arial'><h5><li><a href='#event'>Event Viewer</a></li></h5>"
    Objtextfile.writeline "</ol></font>"
   '
    set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" _
                  & STRcomputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    objtextfile.WriteLine "<ul>"
      objtextfile.WriteLine "<li><b><h2><a name='#SO'> Sistema Operacional</a></h2></b>"
    For Each objItem In colItems
     dtmConvertedDate.Value = objItem.InstallDate
       dtmInstallDate = dtmConvertedDate.GetVarDate
     objtextfile.WriteLine "<ul>"
     objtextfile.WriteLine "<!SO" & objItem.Caption & "Fim_SO>"
       objtextfile.WriteLine "<Font face='arial'><li><pre>" & objItem.Caption & "</font>"
           objtextfile.WriteLine "<Font face='arial'><li>Versão..............................: " &   objItem.Version & "</font>"
       objtextfile.WriteLine "<Font face='arial'><li>Service Pack....................: " &   objItem.ServicePackMajorVersion & "</font>"
       objtextfile.WriteLine "<Font face='arial'><li>Outras descrições...........: " & objItem.OtherTypeDescription & "</font>"
     objtextfile.WriteLine "<Font face='arial'><li>Boot Device.....................: " & objItem.BootDevice & "</font>"
     objtextfile.WriteLine "<Font face='arial'><li>Diretorio Instalação..........: " & objItem.WindowsDirectory & "</font>"
     objtextfile.WriteLine "<Font face='arial'><li>Data de Instalação...........: " & dtmInstallDate & "</font>"
     objtextfile.WriteLine "<Font face='arial'><li>Organização....................: " & objItem.Organization & "</font>"
     objtextfile.WriteLine "<Font face='arial'><li>Usuario Registrado..........: " & objItem.RegisteredUser & "</font>"
     objtextfile.WriteLine "<Font face='arial'><li>Serial Number..................: " & objItem.SerialNumber & "</font>"
   Next
   Set ZoneSet = GetObject("winmgmts:").InstancesOf ("Win32_TimeZone")
   for each System in ZoneSet
    objtextfile.WriteLine "<Font face='arial'><li>Time Zone........................:" & System.StandardName & "</font></pre>"
   next
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "</font></ul>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   ' Coletando Processadores
   Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
   objtextfile.WriteLine "<ul>"
     objtextfile.WriteLine "<li><b><h2><a name='#proc'> Processadores</a></h2></b>"

   For Each objItem In colItems
     objtextfile.WriteLine "<Font face='arial'><ul>"
     objtextfile.WriteLine "<Font face='arial'><li>" & objItem.Description
     Objtextfile.writeline "<!Processador>"
     objtextfile.WriteLine "<Font face='arial'>" & objItem.Name
     Objtextfile.writeline "<!Fim_Proc>"
     objtextfile.WriteLine "<Font face='arial'>" & objItem.MaxClockSpeed & " MHZ"
     objtextfile.WriteLine "</ul>"
   Next
   objtextfile.WriteLine "</ul></font>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   ' Coletando Informa��es da BIOS
   Set colBIOS = objWMIService.ExecQuery("Select * from Win32_BIOS")
   objtextfile.WriteLine "<ul>"
     objtextfile.WriteLine "<li><b><a name='#bios'>BIOS / Hardware</a></b>"

   For Each objbios In colBIOS
     objtextfile.WriteLine "<Font face='arial'><ul>"
     objtextfile.WriteLine "<!Fabricante>"
     objtextfile.WriteLine "<li><pre>Fabricante         : " & objbios.Manufacturer
     objtextfile.WriteLine "<!Fim_Fabricante>"
     objtextfile.WriteLine "<!Serie>"
     objtextfile.WriteLine "<li>Serie/Service Tag  : " & objbios.SerialNumber
     objtextfile.WriteLine "<!Fim_Serie>"
     objtextfile.WriteLine "<li>" & objbios.Name
     objtextfile.WriteLine "<li>Release Date       : " & (Mid(objbios.ReleaseDate, 7, 2)) & "/" & (Mid(objbios.ReleaseDate, 5, 2)) & "/" & (Left(objbios.ReleaseDate, 4))
     objtextfile.WriteLine "<li>SMBIOS Version     : " & objbios.SMBIOSBIOSVersion
     objtextfile.WriteLine "<li>BIOS Version       : " & CStr(objbios.Version)
     objtextfile.WriteLine "</ul>"
   Next

   ' Coletando Modelo do Equipamento
   Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
   For Each objComputer In colSettings
     objtextfile.WriteLine "<ul><Font face='arial'>"
     objtextfile.WriteLine "<!Modelo>"
     objtextfile.WriteLine "<li><pre>Modelo Equipamento : " & objComputer.Model & "</pre>"
     objtextfile.WriteLine "<!Fim_Modelo>"
     objtextfile.WriteLine "</ul>"
   Next
   objtextfile.WriteLine "</ul></font>"
   objtextfile.WriteLine "</ul>"
   w_tipos = Array("Unknown", "Other", "DRAM,Synchronous DRAM", "Cache DRAM", "EDO,EDRAM", "VRAM", "SRAM", _
          "RAM", "ROM", "Flash", "EEPROM", "FEPROM", "EPROM", "CDRAM", "3DRAM", "SDRAM", "SGRAM", _
          "RDRAM", "DDR")
   totalslots = 0
   objtextfile.WriteLine "<ul>"
     objtextfile.WriteLine "<li><b><h2> Memória</h2></b>"

   Set SlotMem = objWMIService.ExecQuery _
     ("Select * from Win32_PhysicalMemoryArray")
     For Each objItem In SlotMem
       totalslots = totalslots + objItem.MemoryDevices
     Next
   totalpentes = 0
   Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
   cont_mem=0
   For Each objItem In colItems
     objtextfile.WriteLine "<Font face='arial'><ul>"
     objtextfile.WriteLine " <li>" & objItem.Tag
     objtextfile.WriteLine " " & Int(objItem.Capacity / 1024 / 1024)
     'objtextfile.WriteLine " " & w_tipos(objItem.MemoryType)
     objtextfile.WriteLine " " & objItem.BankLabel
     objtextfile.WriteLine " " & "Ativa"
     totalpentes = totalpentes + 1
     objtextfile.WriteLine "</ul>"
   cont_mem = cont_mem + objItem.Capacity
   Next

     objtextfile.WriteLine "<!Memoria" & Int(cont_mem / 1024 / 1024) & "Fim_Mem>"

   For i = totalpentes + 1 To totalslots
     objtextfile.WriteLine "<Font face='arial'><ul>"
     objtextfile.WriteLine " <li> Physical Memory "  & i - 1
     objtextfile.WriteLine " Vazio"
     objtextfile.WriteLine "</ul>"
   Next
   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery( _
     "SELECT * FROM Win32_PageFileUsage",,48)
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<ul><br>"
   objtextfile.WriteLine "<Font face='arial'><b> Arquivo de Paginação</b><br>"
   objtextfile.WriteLine "<table border=2 width=500>"
   objtextfile.WriteLine "<tr><td><center>localização</center></td><td><center>Tamanho</center></td><td><center>Utilização Atual</center></td><td><center>" & _
   "Pico de Utilização</center></td></tr>"
   For Each objItem In colItems
     wlinha = " "
     wlinha = "<center><tr><td>"
     wlinha=wlinha & objItem.Description & "</td></center><td><center>"
     wlinha=wlinha & objItem.AllocatedBaseSize & "MB" & "</td></center><td><center>"
     wlinha=wlinha & objItem.CurrentUsage & "MB"& "</td></center><td><center>"
     wlinha=wlinha & objItem.PeakUsage & "MB" & "</td></center></tr>"
     objtextfile.WriteLine wlinha
   Next
   objtextfile.WriteLine "</table>"
   objtextfile.WriteLine "</ul>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   ' Coletando Placas de Rede
   objtextfile.WriteLine "<ul>"
     objtextfile.WriteLine "<li><b><h2><a name='#rede'> Componentes de Rede</a></h2></b>"

   w_Status = Array("Device is working properly.", _
   "Device is not configured correctly.", _
   "Windows cannot load the driver for this device.", _
   "Driver for this device might be corrupted, or the system may be low on memory or other resources.", _
   "Device is not working properly. One of its drivers or the registry might be corrupted.", _
   "Driver for the device requires a resource that Windows cannot manage.", _
   "Boot configuration for the device conflicts with other devices.", _
   "Cannot filter.", "Driver loader for the device is missing.", _
   "Device is not working properly; the controlling firmware is incorrectly reporting the resources for the device.", _
   "Device cannot start.", "Device failed.", "Device cannot find enough free resources to use.", _
   "Windows cannot verify the device's resources.", "Device cannot work properly until the computer is restarted.", _
   "Device is not working properly due to a possible re-enumeration problem.", _
   "Windows cannot identify all of the resources that the device uses.", _
   "Device is requesting an unknown resource type.", "Device drivers need to be reinstalled.", _
   "Failure using the VxD loader.", "Registry might be corrupted.", _
   "System failure. If changing the device driver is ineffective, see the hardware documentation. Windows is removing the device.", _
   "Device is disabled.", "System failure. If changing the device driver is ineffective, see the hardware documentation.", _
   "Device is not present, not working properly", _
   "Windows is still setting up the device.", "Windows is still setting up the device.", _
   "Device does not have valid log configuration.", "Device drivers are not installed.", _
   "Device is disabled; the device firmware did not provide the required resources.", _
   "Device is using an IRQ resource that another device is using.", _
   "Device is not working properly; Windows cannot load the required device drivers.")

   w_statusinfo = Array("Disconnected", "Connecting", "Connected", "Disconnecting", _
              "Hardware Not present", "Hardware disabled", "Hardware malfunction", _
              "Media disconnected", "Authenticating", "Authentication succeeded", _
              "Authentication failed", "Invalid Address", "Credentials required")


   Set colItems2 = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter") '
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<table border=2 width=1100>"
   objtextfile.WriteLine "<tr><td><center>Tipo de Adaptador</center></td><td><center>Descriçao</center></td><td><center>Mac Address</center></td><td><center>" & _
   "Fabricante</center></td><td><center>Status</center></td><td><center>Nome da Conexao</center></td><td><center>Velocidade</center></td></tr>"
   For Each objItem In colItems2
     If IsNull(objItem.AdapterType) then
      wTipo = "-"
     else
      wTipo = objItem.AdapterType
     end if
     If IsNull(objItem.MACAddress) then
      wMac = "- "
     Else
      wMac = objItem.MACAddress
     end if
     If IsNull(objItem.StatusInfo) then
      wMac = "-"
     Else
      wMac = objItem.StatusInfo
    end if
     If IsNull(objItem.NetConnectionID) Then
      wconnection = "-"
     Else
      wconnection = objItem.NetConnectionID
     end if
     If IsNull(objItem.speed) Then
      wvelocidade = "-"
     Else
      wvelocidade = objItem.speed
     end if

     objtextfile.WriteLine "<Font face='arial'><tr><td>" & wTipo & "</td><td>" & objItem.Description & "</td><td>" & _
     wMac & "</td><td>" & objItem.Manufacturer & "</td><td >" & _
     w_Status(objItem.ConfigManagerErrorCode) & "</td><td>" & wconnection & "</td><td>" & wvelocidade & "</td></tr>"

   Next
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   objtextfile.writeline "<li><b><h2><a name='#tcp'> Endereços TCP/IP</a></h2></b>"
   Set colftp = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
   objtextfile.writeline "</ul>"
   'objtextfile.writeline "<table border=2 width=100%>"
   'objtextfile.writeline "<tr><td><center>Placa de rede</center></td><td><center>Endereco TCP/IP</center></td><td><center>Mascara</center></td><td><center>" & _
   ''"Tipo IP</center></td><td><center>Default Gateway</center></td><td><center>DNS Server</center></td><td><center>Wins Server</center></td></tr>"


'Nome placa
For Each objftp In colftp
  wlinha = " "
  wlinha = "<table border=2 width=100% bgcolor=#666666>"
  wlinha = wlinha + "<tr><td width=15%><center><b><font color=#FFFFFF>Placa</b></center></font></td><td><center><b><font color=#FFFFFF><!Nome_Pl" & objftp.Description & "Fim_Pl>" + objftp.Description + "</td></b></center></font><td><center><b><font color=#FFFFFF>"

  'DHCP ou Estatico
  If objftp.dhcpenabled Then
    wlinha = wlinha + "IP Dinamico" + "</td></b></center></font><td>"
  Else
    wlinha = wlinha + "IP Estatico" + "</td></b></center></font><td>"
  End If

wlinha = wlinha + "</table>"
wlinha = wlinha + "<table border=1 width=100%>"

'End_IP
  wlinha = wlinha + "<tr><td width=15%>IP</td><td>"
  strIP = 1
  For Each StrIPaddress In objftp.IPAddress
    wlinha = wlinha + "<!IP" & strIP & StrIPaddress & "Fim_IP" & strIP & ">" + StrIPaddress + "</td><td>"
    strIP = strIP + 1
    'objtextfile.writeline "<!End_IP" & StrIPaddress & "Fim_IP>"
  Next

'Mascara
  wlinha = wlinha + "<tr><td width=15%>Mascara</td><td>"
  For Each strIPSubnet In objftp.IPSubnet
    wlinha = wlinha + strIPSubnet + "</td><td>"
  Next

'Gateway
  wlinha = wlinha + "<tr><td width=15%>Gateway</td><td>"
  For Each strDefaultIPGatewaY In objftp.DefaultIPGateway
    If IsEmpty(strDefaultIPGatewaY) Then
      wlinha = wlinha + "0.0.0.0" + "<br>"
    Else
      wlinha = wlinha + strDefaultIPGatewaY + "<br>"
    End If
  Next
  wlinha = wlinha + "</td><td>"

'DNS
  wlinha = wlinha + "<tr><td width=15%>DNS</td><td>"
  For Each strDNSServer In objftp.DNSServerSearchOrder
    If IsEmpty(strDNSServer) Then
      wlinha = wlinha + "0.0.0.0" + "<br>"
    Else
      wlinha = wlinha + strDNSServer + "<br>"
    End If
  Next
  wlinha = wlinha + "</td><td>"

'Wins
  wlinha = wlinha + "<tr><td width=15%>Wins</td><td>"
  If IsNull(objftp.WINSPrimaryServer) Then
    wlinha = wlinha + "0.0.0.0" + "<br>"
  Else
    wlinha = wlinha + objftp.WINSPrimaryServer + "<br>"
  End If
  If IsNull(objftp.WINSSecondaryServer) Then
    wlinha = wlinha + "0.0.0.0" + "</td><td>"
  Else
    wlinha = wlinha + objftp.WINSSecondaryServer + "</td><td>"
  End If

wlinha = wlinha + "</table>"
objtextfile.writeline wlinha
Next

   objtextfile.writeline "</table>"
   objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   objtextfile.WriteLine "<ul>"
     objtextfile.WriteLine "<li><b><h2><a name='#discos'> Discos Locais</a></h2></b>"
     objtextfile.WriteLine "</ul>"
     objtextfile.WriteLine "<table border=2 width=400>"
     objtextfile.WriteLine "<tr><td><center>Unidade</center></td><td><center>Tipo</center></td></tr>"
     Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      Set colDisks = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk")
      For Each objDisk in colDisks
         W_ID= objDisk.DeviceID
     Select Case objDisk.DriveType
       Case 1
         W_ID1 = "Tipo de Disco não Detectado"
       Case 2
         W_ID1 = "Disco removível ou Disquete"
       Case 3
         W_ID1 = "Disco Rígido Local"
       Case 4
         W_ID1 ="Drive de Rede"
       Case 5
         W_ID1 = "Unidade de CD"
       Case 6
         W_ID1 ="RAM disk."
       Case Else
         W_ID1 = "Tipo de Disco não Detectado"
     End Select
     objtextfile.WriteLine "<tr><td><center>" & W_ID & "</center></td><td><center>" & W_ID1 & "</center></td></tr>"
   Next
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"

   '
   objtextfile.WriteLine "<ul>"
     objtextfile.WriteLine "<li><b><h2><a name='#discos'> Detalhe dos Discos Locais</a></h2></b>"
     objtextfile.WriteLine "</ul>"
     Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
   Set colDiskDrives = objWMIService.ExecQuery ("Select * from Win32_DiskDrive")
     strDisk = 1
   For each objDiskDrive in colDiskDrives
     objtextfile.WriteLine "<Font face='arial'><ul>"
     objtextfile.WriteLine "<li><pre><b>Caption:         </b>" & ltrim(objDiskDrive.Caption )
     objtextfile.WriteLine "<li><b>Device ID:        </b>" & ltrim(objDiskDrive.DeviceID )
     objtextfile.WriteLine "<li><b>Interface Type:     </b>" & ltrim(objDiskDrive.InterfaceType)
     objtextfile.WriteLine "<li><b>Manufacturer:      </b>" & ltrim(objDiskDrive.Manufacturer)
     objtextfile.WriteLine "<li><b>Model:          </b>" & ltrim(objDiskDrive.Model)
     objtextfile.WriteLine "<li><b>Name:          </b>" & ltrim(objDiskDrive.Name)
     objtextfile.WriteLine "<li><b>Partitions:       </b>" & ltrim(objDiskDrive.Partitions)
     objtextfile.WriteLine "<li><b>SCSI Bus:        </b>" & ltrim(objDiskDrive.SCSIBus)
     objtextfile.WriteLine "<li><b>SCSI Logical Unit:    </b>" & ltrim(objDiskDrive.SCSILogicalUnit)
     objtextfile.WriteLine "<li><b>SCSI Port:        </b>" & ltrim(objDiskDrive.SCSIPort)
     objtextfile.WriteLine "<li><b>SCSI TargetId:      </b>" & ltrim(objDiskDrive.SCSITargetId)
     objtextfile.WriteLine "<!Disco" & strDisk & int(objDiskDrive.Size/1024/1024) & "Fim_Disco" & strDisk & ">"
     objtextfile.WriteLine "<li><b>Size:          </b>" & int(objDiskDrive.Size/1024/1024) & "MB"
     objtextfile.WriteLine "<li><b>Status:         </b>" & ltrim(objDiskDrive.Status)
     objtextfile.WriteLine "</ul>"
	 strDisk = strDisk+1
   Next

   '
   ' Discos l�gicos
   Const HARD_DISK = 3
   objtextfile.WriteLine "<ul>"
     objtextfile.WriteLine "<li><b><h2><a name='#discos'> Partições</a></h2></b>"
   Set colDiskDrives = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "")
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<table border=2 width=500>"
   objtextfile.WriteLine "<tr><td><center>Unidade</center></td><td><center>Tamanho (MB)</center></td><td><center>Tipo Partiçao" & _
   "</td><td><center>Espaço Livre(MB)</center></td></tr>"
   For Each objdisk In colDiskDrives
     objtextfile.WriteLine "<tr><td><center>" & objdisk.DeviceID & "</center></td><td><center>" & " " _
     & Int(objdisk.Size / 1024 / 1024) & "</center></td><td><center>" & " " _
     & objdisk.FileSystem & "</center></td><td><center>" _
      & int(objDisk.FreeSpace/(1024*1024)) & "</center></td></tr>"
   Next
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"

   '
   objtextfile.WriteLine "<ul>"
     objtextfile.WriteLine "<li><b><h2><a name='#discos'> Unidade de CDROM</a></h2></b>"
   objtextfile.WriteLine "</ul>"
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_CDROMDrive")
   For Each objItem in colItems
      objtextfile.WriteLine "<tr><td>" & objItem.Caption & "</td></td>"
   Next


   '
   'Controladoras de Disco"
   objtextfile.WriteLine "<ul>"
   objtextfile.WriteLine "<li><b><h2><a name='#controladoras'> Controladoras de Disco</a></h2></b>"
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<table border=2 width=600>"
   objtextfile.WriteLine "<tr><td><center>Controladora</center></td><td><center>Driver</center></td><td><center>Fabricante</center></td><td><center>Status</center></td></tr>"
   Set wplaca = objWMIService.ExecQuery("Select * from Win32_SCSIController")
   For Each objplaca In wplaca
   objtextfile.WriteLine "<tr><td>" & objplaca.Name & "</td><td>" & " " _
   & objplaca.DriverName & "</td><td>" & " " _
   & objplaca.Manufacturer & "</td><td>" & " " _
   & objplaca.Status & "</td></tr>"
   Next
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "</table>"
   '
   objtextfile.WriteLine "<ul>"
   objtextfile.WriteLine "<li><b><h2><a name='#backup'> Unidade de Backup</a></h2></b></ul>"
   objtextfile.WriteLine "<table>"
   Set wfita = objWMIService.ExecQuery("Select * from Win32_TapeDrive")
   For Each objfita In wfita
     objtextfile.WriteLine "<tr><td>" & objfita.Caption & "</td></tr>"
   Next
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   'Coleta usuarios locais
   Set colAccounts = GetObject("WinNT://" & STRcomputer & "")
   Set colGroups = GetObject("WinNT://" & STRcomputer & "")
   colAccounts.Filter = Array("user")
   colGroups.Filter = Array("group")
   objtextfile.WriteLine "<ul>"
   objtextfile.WriteLine "<li><b><h2><a name='#usuarios'> Usuarios Locais</a></h2></b>"
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<table border=2 width=1000>"
    objtextfile.WriteLine "<tr><td><center>Login</center></td><td><center>Nome Completo</center></td><td><center>Descriçao</center>" & _
   "</td><td><center>Status</center></td><td><center>Grupos</center></td></tr>"
   For Each objUser In colAccounts
     Wgrupo = " "
     If objUser.AccountDisabled then
    wstatus = "Inativa"
     Else
    wstatus = "Ativa"
     End if
     For Each objGroup In colGroups
       For Each objuserMBR In objGroup.Members
         If objuserMBR.Name = objUser.Name Then
           Wgrupo = Wgrupo & objGroup.Name & "<br>"
         End If
       Next
     Next
     If Len(objUser.FullName) <= 1 Then
       wfulname = "-"
     Else
      wfulname = objUser.FullName
     End If
     If Len(objUser.Description) <= 1 then
    wdescricao = "-"
     else
    wdescricao = objUser.Description
     end if
     objtextfile.WriteLine "<tr><td>" & objUser.Name & "</td><td>" & " " _
     & wfulname & "</td><td>" _
     & wdescricao & "</td><td>" _
     & wstatus & "</td><td>" _
     & Wgrupo & "</td></tr> "
   Next
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   'Coleta Software

	ssig="Unable to open registry key"

	set wshShell= Wscript.CreateObject("WScript.Shell")
	strKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
	on error resume next
	present = WshShell.RegRead(strKey)
	if err.number<>0 then
		if right(strKey,1)="\" then    'strKey is a registry key
			if instr(1,err.description,ssig,1)<>0 then
				'bExists=true
				strKey = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
			else
				'bExists=false
				strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
			end if
		else    'strKey is a registry valuename
			'bExists=false
			strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
		end if
		err.clear
	else
		'bExists=true
		strKey = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
	end if

   Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
   'strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
   strEntry1a = "DisplayName"
   strEntry1b = "QuietDisplayName"
   strEntry2 = "InstallDate"
   strEntry3 = "VersionMajor"
   strEntry4 = "VersionMinor"

   objtextfile.WriteLine "<ul>"
   objtextfile.WriteLine "<li><b><h2><a name='#software'> Softwares Instalados</a></h2></b>"
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<table border=2 width=800>"
   objtextfile.WriteLine "<tr><td><b><center>Software</center></b></td><td><b><center>Data de Instalação</center></b></td>"& _
   "<td><b><center>Versão</center></b></td>"  & _
   "</tr>"

'BACKUP


Set objReg = GetObject("winmgmts://" & strComputer & _
 "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey, arrSubkeys
'WScript.Echo "Installed Applications" & VbCrLf
For Each strSubkey In arrSubkeys
  tableLine = ""
  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, _
   strEntry1a, strValue1)
  If intRet1 <> 0 Then
    objReg.GetStringValue HKLM, strKey & strSubkey, _
     strEntry1b, strValue1
  End If
  If strValue1 <> "" Then
    'WScript.Echo VbCrLf & "Display Name: " & strValue1
  End If
  objReg.GetStringValue HKLM, strKey & strSubkey, _
   strEntry2, strValue2
  If strValue2 <> "" Then
    'WScript.Echo "Install Date: " & strValue2
  End If
  objReg.GetDWORDValue HKLM, strKey & strSubkey, _
   strEntry3, intValue3
  objReg.GetDWORDValue HKLM, strKey & strSubkey, _
   strEntry4, intValue4
  If intValue3 <> "" Then
     'WScript.Echo "Version: " & intValue3 & "." & intValue4
  End If
  objReg.GetDWORDValue HKLM, strKey & strSubkey, _
   strEntry5, intValue5
  If intValue5 <> "" Then
    'WScript.Echo "Estimated Size: " & Round(intValue5/1024, 3) & " megabytes"
  End If
  'objtextfile.WriteLine wlinha + "<br>"
  if Len(strValue1) > 0 then
    tableLine = "<tr><td><i>" & wfonte & strValue1 & "</font></i></td><td><i>"
    if strValue1 = "AR" then
        alreadyExists(STRcomputer)
    end if
  end if
  if Len(strValue2) > 0 then
    tableLine = tableLine + (wfonte & Int(strValue2 Mod 100) & "/" & (Int(strValue2 - 10000 * Int(strValue2 / 10000)) \ 100) & "/" & Int(strValue2 / 10000) & "</font></i></td><td><i>")
  Else
    if Len(strValue1) > 0 then
        tableLine = tableLine + (wfonte & "</font></i></td><td><i>")
    end if
  end if
  if Len(intValue3) > 0 then
    tableLine = tableLine + (wfonte & intValue3 & "." & intValue4 & "</font></i></td>")
  end if
  'objtextfile.WriteLine "<tr><td><i>" & wfonte & strValue1 & "</font></i></td><td><i>" '& _
  'wfonte & strValue2.GetVarDate & "</font></i></td><td><i>" & _
  'wfonte & intValue3 & "." & intValue4 & "</font></i></td><td><i>" & _
  'wfonte & Round(intValue5/1024, 3) & " Megabytes" & "</font></i></td><td><i>" & _
  'wfonte & strValue3 & "." & strValue2 & "</font></i></td><td><i>"
  objtextfile.WriteLine tableLine
Next
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   'Coleta Serviços
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & STRcomputer & "\root\cimv2")
    Set colListOfServices = objWMIService.ExecQuery _
       ("Select * from Win32_Service")
   objtextfile.WriteLine "<ul>"
   objtextfile.WriteLine "<li><b><h2><a name='#servicos'> Serviços</a></h2></b>"
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<table border=2 width=800>"
   objtextfile.WriteLine "<tr><td><b><center>Serviço</center></b></td><td><b><center>Startup</center></b></td><td><b>" & _
   "<center>Status</center></b></td><td><b><center>Usuario de Startup</center></b></b></td></tr>"
   For Each objService In colListOfServices
     if objService.State = "Stopped" then
    wfonte= "<font color=red>"
     else
    wfonte = "<font>"
     end if
     objtextfile.WriteLine "<tr><td><i>" & wfonte & objService.Caption & "</font></i></td><td><i>" & _
     wfonte & objService.StartMode & "</font></i></td><td><i>" & _
     wfonte & objService.State & "</font></i></td><td><i>" & _
     wfonte & objService.StartName & "</font></i></td></tr>"
   Next
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   'Coleta Tamanho dos Diretorios
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & STRcomputer & "\root\cimv2")

   Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")

   objtextfile.WriteLine "<ul>"
   objtextfile.WriteLine "<li><b><h2><a name='#compartilhamentos'>Compartilhamentos</a></h2></b>"
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<table border=2>"
   objtextfile.WriteLine "<tr><td><b><center>Compartilhamento</center></b></td><td><b><center>" & _
   "Caminho</center></b></b></td></tr>"

   For Each objShare In colShares
     objtextfile.WriteLine "<tr><td><i>" & objShare.Name & "</i></td><td><i>" _
     & objShare.Path & "</i></td></tr>"
   Next
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   'Coleta Impressora
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & STRcomputer & "\root\cimv2")
   Set colInstalledPrinters = objWMIService.ExecQuery _
     ("Select * from Win32_PrinterDriver")
   objtextfile.WriteLine "<ul>"
   objtextfile.WriteLine "<li><b><h2><a name='#impressora'> Drivers de Impressoras</a></h2></b>"
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<table border=2>"
   objtextfile.WriteLine "<tr><td><b>Impressora Local</b></b></td></tr>"

   For Each objPrinter In colInstalledPrinters
     objtextfile.WriteLine "<tr><td><i>" & objPrinter.Name & "</i></td></tr>"
   Next
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   'Coleta portas de impressora
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & STRcomputer & "\root\cimv2")

   Set colPorts = objWMIService.ExecQuery _
     ("Select * from Win32_TCPIPPrinterPort")
   objtextfile.WriteLine "<ul>"
   objtextfile.WriteLine "<li><b><h2><a name='#portas'> Portas de Impressora</a></h2></b>"
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<table border=2>"
   objtextfile.WriteLine "<tr><td><b><center>Impressora Local</center></b></td><td><b><center>Endereco Host</center></b></td><td><b>" & _
   "<center>Porta</center></b></td><td><b><center>Protocolo</center></b></b></td></tr>"
   For Each objPort In colPorts
     objtextfile.WriteLine "<tr><td><i>" & objPort.Description & "</i></td><td><i>" & _
     objPort.HostAddress & "</i></td><td><i>" & _
     objPort.Name & "</i></td><td><i>" & _
     objPort.PortNumber & "</i></td><td><i>" & _
     objPort.Protocol & "</i></td></tr>"
   Next
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   '
   objtextfile.WriteLine "<ul>"
   objtextfile.WriteLine "<li><b><h2><a name='#event'> Event Viewer</a></h2></b>"
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "<table border=2 width=100>"
   Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate,(Security)}!\\" & _
       strComputer & "\root\cimv2")
   Set colLogFiles = objWMIService.ExecQuery _
     ("Select * from Win32_NTEventLogFile " _
       & "Where LogFileName='Security'")
   For Each objLogFile in colLogFiles
    objtextfile.WriteLine "<tr><td><i>Security </i></td><td><i>" & int(objLogfile.MaxFileSize/1024) & "MB" & "</i></td><td><i>"
   Next
   Set colLogFiles = objWMIService.ExecQuery _
     ("Select * from Win32_NTEventLogFile " _
       & "Where LogFileName='application'")
   For Each objLogFile in colLogFiles
    objtextfile.WriteLine "<tr><td><i>Application</i></td><td><i>" & int(objLogfile.MaxFileSize/1024) & "MB" & "</i></td><td><i>"
   Next
   Set colLogFiles = objWMIService.ExecQuery _
     ("Select * from Win32_NTEventLogFile " _
       & "Where LogFileName='system'")
   For Each objLogFile in colLogFiles
    objtextfile.WriteLine "<tr><td><i>System</i></td><td><i>" & int(objLogfile.MaxFileSize/1024) & "MB" & "</i></td><td><i></tr>"
   Next
   objtextfile.WriteLine "</ul>"
   objtextfile.WriteLine "</table>"
   Objtextfile.writeline "<font size=1 face='Helvetica' color=blue><a href='#menu'>Menu</a></font>"
   Objtextfile.writeline "</html>"

Function alreadyExists(str)
    Set objFile = Wscript.CreateObject("Scripting.FileSystemObject")
    Set objShell = Wscript.CreateObject("Wscript.Shell")
    Const FORREADING = 1
    Const FORWRITING = 2
    Const FORAPPENDING = 8
    Dim sToSearch: sToSearch = str
    Dim sFileName: sFileName = "\\pmmcad02\VMs\inventario\lista_AR.txt"
    Dim sContent, Found
    If Not objFile.FileExists(sFileName) Then
        ' MsgBox "File Not Found"
        WScript.Quit 0
    End If
    Set TxtFile = objFile.OpenTextFile(sFileName,FORREADING)
    sContent = TxtFile.ReadAll
    If InStr(sContent,sToSearch) Then
        Found = True
    End If
    Set TxtFile = Nothing
    If Not Found Then
        Set TxtFile = objFile.OpenTextFile(sFileName,FORAPPENDING)
        TxtFile.WriteLine sToSearch
    End If
End Function
