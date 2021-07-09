<#

LocationName = The Arbitrary location where the lab resides. It's used as an identifier when running functions. 
SharedFolder = The shared folder where the Json file will be stored g
JsonFileName = The filename of the Json file
OUpath = The LDAP path of the computer's you'd like to manipulate. 
#>


$Locations = @{
    LCH = @{
        SharedFolder = '\\SERVER\SHARE\'
        OUPath = 'ou=Lab1 Computers, ou= Desktops ,dc=Contoso, dc=local'

    }
    SUL =@{
        SharedFolder = '\\SERVER2\SHARE2'
        OuPath = 'ou=Lab2 Computers,ou= Desktops ,dc=safetycouncil, dc=local'

    }
}


## Logging 
$logfilepath="C:\Logs\PowerMLog.txt"

function Write-Log ($message){
$message +" - "+ (Get-Date).ToString() >> $logfilepath
}

## Log Rotation. 
function Rotate-Log {
    $LogDir = Split-Path $logfilepath
    [int64]$MaxFileSize = 20mb
    if (Test-Path $logfilepath){
        $ActiveLogFile = Get-ChildItem $LogFilePath
            if($ActiveLogFile.Length -ige $MaxFileSize){
                $OldLogFileName = $ActiveLogFile.Name+".OLD"
                if(Test-Path "$($LogDir)\$($OldLogFileName)"){
                    Remove-Item -Path "$($Logdir)\$($OldLogFileName)"
                          }
                Rename-Item -Path $logfilepath -NewName "$($ActiveLogFile.Name).OLD"   
                Write-Log "Log File Exceeded $MaxFileSize KB. Renamed File to $($ActiveLogfile.Name).OLD and started a Fresh Log"
            }
    }
}
Rotate-Log


## End Logging 




if (Get-Module -ListAvailable -Name ActiveDirectory) {
    Write-Host "Module exists"
    } 
    else {
        Write-Host "Missing Active Directory Module, You need to install RSAT Tools on your Computer. Install that before continuing."
        Exit
}


## Get Mac addresses of Lab Computer

    
function Get-Macs{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Location

    )
     
     $Config = $Locations.$Location
     $LabComputersJsonFile  = "$($Config.Sharedfolder)$($location)Computers.Json"
     if (!(Test-Path $LabComputersJsonFile)){
        $emptyDict = @{}
        $emptyDict | ConvertTo-Json | Out-File $LabComputersJsonFile
    }
     $OUPath = $Config.OUPath
     $LabComputersInAD = Get-ADComputer -Filter * -SearchBase $OUpath | Select-object DNSHostName
     $LabComputers = Get-Content $LabComputersJsonFile | Out-String | ConvertFrom-Json
     $LabComputersHash = @{}
     foreach ($property in $LabComputers.PSObject.Properties) {
        $LabComputershash[$property.Name] = $property.Value
     }

    ## GET COMPUTERS IN AD AND CHECK IF ALIVE      
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, 5)
    $RunspacePool.Open()
    $ComputersToCheck = @()
    $ScriptBlock = {
        param($Computer)
        $IsAlive = Test-Connection -ComputerName $computer -count 1 -BufferSize 16 -Quiet
        $IsAlive
        }

    foreach ($Computer in $LabComputersInAD.DNSHostName){
        $Powershell = [PowerShell]::Create().AddScript($ScriptBlock).AddArgument($Computer)
        $PowerShell.RunspacePool = $RunspacePool
        $Invoke = $Powershell.BeginInvoke()
        $Data = $Powershell.EndInvoke($Invoke)
        if($Data){
            if ($LabComputersHash.Keys -contains $Computer){ 
                $LabComputersHash[$computer].timestamp = (get-date -format "MM/dd/yyyy")
            }
            else{
                 $ComputersToCheck += $Computer
            }
        }
        if(!$Data){
              if ($LabComputersHash.Keys -contains $Computer){
                        Write-Log "$Computer is currently in JSON File, but Didn't Respong to a ping this time. If no response for 14 days, it will be removed from JSON"
                        Write-Host "$Computer is currently in JSON File, but Didn't Respong to a ping this time. If no response for 14 days, it will be removed from JSON"
              }
              else{
                Write-Log "$computer is listed in AD, but didn't respond to ping. Not Adding To Json"
              }  
        }
    }
    $RunspacePool.Dispose()


## END GET COMPUTERS IN AD AND CHECK IF ALIVE     

    if (!($ComputerstoCheck)){
        Write-Log "No Changes made to $LabComputersJsonFile . All Computers $LabComputersJsonFile are up to date."
        Return 
    }
    else{      
        Write-Log "The following computers are In AD and are responding. Will Attempt to get MAC and add to JsonFile :"
        foreach ($computer in $ComputersToCheck){Write-Log $computer}
        }
       

    ### GET MAC ADDRESSES 
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, 15)
    $RunspacePool.Open()
    $Macs = @{}
    $ScriptBlock = {
        param($Computer)
        try{
            $mac = Invoke-Command -ComputerName $Computer -scriptblock {Get-NetAdapter | Where-Object {($_.Status -eq 'up' -and $_.Name -eq "Ethernet")} | Select-Object MacAddress} -ErrorAction Stop
            $mac.MacAddress
            }
        catch{
            $mac = 'ERROR'
            $mac
            }
    }
            
    foreach ($Computer in $ComputersToCheck){
        $Powershell = [PowerShell]::Create().AddScript($ScriptBlock).AddArgument($Computer)
        $PowerShell.RunspacePool = $RunspacePool
        $Invoke = $Powershell.BeginInvoke()
        $Data = $Powershell.EndInvoke($Invoke)
        if (!$Data[0]){
            $Macs += @{$Computer = 'ERROR'}
            }
        if($Data[0] -eq 'ERROR'){
            $Macs += @{$Computer = 'ERROR'}
            }
        else{
            $Macs += @{$computer = $Data[0]}  
        }
    }
    $RunspacePool.Dispose()

    foreach ($Computer in $Macs.GetEnumerator()){
        if ($computer.value -ne "ERROR"){
            $LabComputersHash[$computer.Key] = @{}
            $LabComputersHash[$computer.Key] = $LabComputersHash[$computer.Key] +  @{mac = $Computer.Value}
            $LabComputersHash[$computer.Key] =$LabComputersHash[$computer.Key] +  @{timestamp = (get-date -format "MM/dd/yyyy")}
            Write-Log "Added $($computer.Key) to Json File"
            }
        if ($computer.value -eq "ERROR"){
            Write-Log "Could not Add $($Computer.key) to Json file, Getting Mac Address Timed out" 
            Write-Host "Could not Add $($Computer.key) to Json file, Getting Mac Address Timed out" 
            }
    }
    $LabComputersHash| ConvertTo-Json | Set-Content $LabComputersJsonFile
    Write-Log "Fin Mac Check"
    Write-Host "Fin Mac Check"
}

## END GET MAC ADDRESSES 


function Stop-Computers{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Location
    )
    Write-Log "Running Stop-Computers"
    Write-Log "Checking to see if any updates need to made to JSON file (Adding or Removing computers), before Stopping"
    Get-Macs -Location $Location
    $Config = $Locations.$Location
    $LabComputersJsonFile = "$($Config.Sharedfolder)$($location)Computers.Json"
    if (!(Test-Path  $LabComputersJsonFile )){
    $emptyDict = @{}
    $emptyDict | ConvertTo-Json | Out-File $LabComputersJsonFile
        }
    $LabComputers = Get-Content $LabComputersJsonFile | Out-String | ConvertFrom-Json

   

    $LabComputersHash = @{}
    foreach ($property in $LabComputers.PSObject.Properties) {
        $LabComputershash[$property.Name] = $property.Value
    } 

    ## Check the Timestamp of when the MAC was added to the JSON. If it's older than 14 days, remove it. 
    $ComputerstoRemove = @()
    foreach ($computer in $LabComputersHash.GetEnumerator()){
            if ((get-date $computer.value.timestamp) -lt ((Get-Date).AddDays(-14))){
                $ComputerstoRemove += $computer.key
                Write-Output $computer.Key "Removed from json file. - Older than 14 days" 
                Write-Log "$computer.Key Removed from json file. - Older than 14 days" 
                }
    }
    if($ComputerstoRemove){
        foreach ($computer in $ComputerstoRemove){
            $LabComputersHash.Remove($Computer)
            $LabComputersHash | ConvertTo-Json | Set-Content $LabComputersJsonFile
            }
    }
    

    $ComputersArray = $LabComputersHash.Keys | % ToString    
    foreach ($Computer in $ComputersArray){
        Invoke-Command -ComputerName $Computer -ScriptBlock {shutdown /s /f }
        Write-Log "Sent Stop-Computer Command to $Computer"
        Write-Output "Sent Stop-Computer Command to $Computer"
        Sleep 1
        }  
    Write-Log "Done Stopping Computers" 
}


function Get-BroadcastAddress() {
    $defaultRouteNic = Get-NetRoute -DestinationPrefix 0.0.0.0/0 | Sort-Object -Property RouteMetric | Select-Object -ExpandProperty ifIndex
    $address = Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex $defaultRouteNic | Select-Object -ExpandProperty IPAddress
    $prefixlength = Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex $defaultRouteNic | Select-Object -ExpandProperty PrefixLength
    $subnetmask = (0..3|%{+("`0Ààðøüþÿ"+"`0"*99)[($prefixlength,8)[$prefixlength-ge8]];$prefixlength-=8})-join'.'
    $addressBytes = $address -split "\." | ForEach { [byte] $_ };
    $subnetMaskBytes = $subnetMask -split "\." | ForEach { [byte] $_ };
 
    $broadcastBytes = New-Object byte[] 4;
    for($i = 0; $i -lt 4; ++$i) {
        $inverted = -bnot [byte] $subnetMaskBytes[$i] + 256;
        $broadcastBytes[$i] = $addressBytes[$i] -bor $inverted;
    }
 
    return $broadcastBytes -join ".";
}


function Start-Computers{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Location
    )
    Write-Log "Running Start-Computers"
    $Config = $Locations.$Location
    $LabComputersJsonFile  = "$($Config.Sharedfolder)$($location)Computers.Json"
    if (!(Test-Path $LabComputersJsonFile)){
        Write-Output "Json File does not exist here $LabComputersJsonFile, therefore there are no computers to turn on. You'll need to manually turn on the computers, and run 'Get-Macs -Location $Location'"
    }
    $LabComputers = Get-Content $LabComputersJsonFile | Out-String | ConvertFrom-Json
    
    $UDPClient = New-Object System.Net.Sockets.UdpClient
    
    $LabComputersHash = @{}
    foreach ($property in $LabComputers.PSObject.Properties) {
        $LabComputershash[$property.Name] = $property.Value
    }
    
     
    foreach ($computer in $LabComputersHash.GetEnumerator()){ 
       $mac = $computer.Value.mac
       $Broadcast = Get-BroadcastAddress
       $MacInBytes = $mac -split "[:-]" | ForEach-Object { [Byte] "0x$_"} # Convert Mac STring to Byte array 
       $MagicPacket = [Byte[]] $MagicPacket = (,0xFF * 6) + ($MacInBytes  * 16) ## Add 6 255s, repeat Mac Byte array 16 times 
       $UDPClient.Connect($Broadcast,4343)
       $UDPClient.Send($MagicPacket , $MagicPacket.Length) | Out-null
       $UDPClient.Send($MagicPacket , $MagicPacket.Length) | Out-null
       Write-Output "Sent Power ON Request to $($computer.Name)"
       Write-Log "Sent Power ON Request to $($computer.Name)"
    }
    $UDPClient.close()
}


