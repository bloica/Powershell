
# =========================================================
# ===     	   ADHC_PortScan2012.01.PS1                 ===
# =========================================================
# === Created by:     Peter Van Keymeulen               ===
# ===              on:     24.04.2012      Ver. 2012.01 ===
# === Last Updated by: Ritchie Henuset                  ===
# ===              on: 17/06/2015                       ===
# =========================================================
#cls

# -----------------------------------------------------------------------------------------------------------------------
# set variables
# -----------------------------------------------------------------------------------------------------------------------
<#
Windows 2008
Client port		Server port		Type of traffic
TCP Dynamic		TCP 135, 49152–65535	RPC, EPM
TCP and UDP Dynamic	TCP and UDP 389		LDAP
TCP Dynamic		TCP 636			LDAP SSL
TCP Dynamic		TCP 3268		GC
TCP Dynamic		TCP 3269		GC SSL
TCP and UDP 53, Dynamic	TCP and UDP 53		DNS
TCP and UDP Dynamic	TCP and UDP 88		Kerberos
TCP and UDP Dynamic	TCP-NP and UDP-NP 445	Security Accounts Manager (SAM), LSA
TCP Dynamic		UDP 138			NetBIOS Datagram Service
			TCP Static 53248	FRsRpc  (needed for RODC)
			TCP 1352		Lotus Notes
#>
import-module ActiveDirectory
#$ErrorActionPreference = "SilentlyContinue"
$PortsToSCan=@(53,88,135,139,389,445,464,636,3268,3269,53248)
$DCComputerName = $env:COMPUTERNAME
$timeout=3000
$verbose = $False
$LogFile = "\\sptw0087\data$\LogFiles\DCPorts\$($DCComputerName)_AD_Ports_Overview.log"
#$LogFile = "C:\Temp\AD\Reports\AD_Ports_Overview.log"

#Log Headers
Out-File -FilePath $LogFile -Append -InputObject "DC Source,DC Destination,53,53,88,135,139,389,445,464,636,3268,3269,53248"

# -----------------------------------------------------------------------------------------------------------------------
# Functions
# -----------------------------------------------------------------------------------------------------------------------
function Logging 
{ 
$datetime=get-date
$datetime=$datetime.ToShortDateString() +" : "+ $datetime.ToShortTimeString()
write-host $datetime  " : "  $args[0] 
}

# -----------------------------------------------------------------------------------------------------------------------
# Start Main Procedure
# -----------------------------------------------------------------------------------------------------------------------




#get-all domain controllers
$DomainControllers=get-addomaincontroller -Filter * -server ebm.infra.shared.etex| select-object Name
$DomainControllers+=get-addomaincontroller -Filter * -server infra.shared.etex | select-object Name
$DomainControllers+=get-addomaincontroller -Filter * -server systems.infra.shared.etex | select-object Name
#Out-File -FilePath $LogFile -Append -InputObject $DomainControllers
$iFormat = 0
$DomainControllers | %{
    $PortsLog = "$DCComputerName,$($_.Name),"
    $srv=$_.Name.ToString()
    foreach ($port in $PortsToSCan){
        If($iFormat -eq 1){$PortsLog+= ","}
        # Create TCP Client
        $failed = $false
        $tcpclient = new-Object system.Net.Sockets.TcpClient
        $iar = $tcpclient.BeginConnect($srv,$port,$null,$null)
        $wait = $iar.AsyncWaitHandle.WaitOne($timeout,$false)
        if(!$wait)
        {
            $tcpclient.Close()
            $failed = $true
        }else{
            $error.Clear()
            $tcpclient.EndConnect($iar) | out-Null
            if(!$?){if($verbose){write-host $error[0]};$failed = $true}
            $tcpclient.Close()
        }
        if($failed){
            #Logging ('Testing port: ' + $srv + ':'  + $port + ' Failed !!!')
            $PortsLog+= "NOK" 
        }else{
            #Logging ('Testing port: ' + $srv + ':'  + $port + ' OK') 
            $PortsLog+= "OK"         
        }
        If($iFormat -eq 0){$iFormat = 1}

    }
    $iFormat = 0
    Out-File -FilePath $LogFile -Append -InputObject $PortsLog
        #Dynamic Port Rangs
        #49152..65535 | %{
            # Create TCP Client
        #    $failed = $false
        #    $tcpclient = new-Object system.Net.Sockets.TcpClient
        #    $iar = $tcpclient.BeginConnect($srv,$_,$null,$null)
        #    $wait = $iar.AsyncWaitHandle.WaitOne($timeout,$false)
        #    if(!$wait){
        #        $tcpclient.Close()
        #        $failed = $true
        #    }else{
        #        $error.Clear()
        #       $tcpclient.EndConnect($iar) | out-Null
        #        if(!$?){if($verbose){write-host $error[0]};$failed = $true}
        #        $tcpclient.Close()
        #    }
        #   if($failed){
        #        Logging ('Testing port: ' + $srv + ':'  + $_ + ' Failed !!!')
        #    }else{
        #        Logging ('Testing port: ' + $srv + ':'  + $_ + ' OK')
        #    }
        #}
}


 

