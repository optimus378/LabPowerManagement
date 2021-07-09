# LabPowerManagmenet
Starts and Stops Computers in a Lab Setting on a Schedule. 



## Overview

This script was created for a very specific environment, but could be adapted for other purposes. 
Read the actual code. I did my best to explain how it works given the amount of time I felt like putting into the explanation. 

- One goal of the script was to acheive this witout using any outside executables or supporting modules. 
- Another goal was design this in a "set it and forget it" sort of way. 
   - This is acheived by auto pruning the list of computers that will be started/stopped before each shutdown. 

This Module manages the powering on and off of set of computers listed in a specified Active Directly OU. 
It works by using set of function to dynamically update a JSON file with Mac Addresses of the Lab Computers based an Active Directory OU path that you set in the script. 
It can Power on or Power off the computers using the properties defined in the JSON.  

The JSON File contains the Computer Name, Mac Address, and a timestamp. The functions in this module reference the JSON file to Shutdown all the computers via Hostname or Power On all the computers by sending Wake On Lan Packets to each Mac Address in said JSON File 
The script can be run on schedule as well as manually triggered from a workstation using batch files. 

** Do note that there are little to no error handling in this module and I don't expect to create any. It's quick and dirty. If you get errors, you'll have to read the Powershell &output.  **   

## Prerequisites:

- A Windows Domain Environment,
- A computer with the Active Directory PowerShell Module installed. 
- WinRM properly enabled on all computers that you'd like to run this script against.  
- Firewall Properly configured to allow WinRM access on all computers. (In essense you need to be able to access all computers in the lab remotely via PowerShell)

## First-time Usage Instructions or Manual Running. -- 

  If you will be running this Module from multiple computers, you should share a folder(s) in a central location(s) where the JSON file(s) is/are accessible to the computers running the script.
  
  You'll need to define some things at the top of the script: 
           $Locations 
           -  Location Name  
           -  Shared Folder path
           -  OU Path of the computers you'd like to target 
           $LogFilePath: Location of the Log File

    $Locations = @{
        Lab1= @{
            SharedFolder = '\\SERVER\SHARE\'
            OUPath = 'ou=Lab1 Computers, ou= Desktops,dc=contoso, dc=local'

        }
        Lab2 =@{
            SharedFolder = '\\SERVER2\SHARE\'
            OUPath = 'ou=Lab2 Computers, ou= Desktops,dc=contoso, dc=local'

        }
    }



  Save the Module and also place it in your Shared Folder or whereover you'll run it from. 
  Open PowerShell and navigate to the the powershell module path. 

  Type 'Import-Module SCPowerManagement.psm1'
  Nice, the module is imported. From here you are able to run the functions listed below. 

## Functions Overview --  

### Function: Get-Macs

  Notes: 
  You shouldn't need to manually run this function unless you want to prepopulate the JSON files with mac addresses, Or just manually refresh them. 
  This function gets all of the mac addresses of the computers in the OU you specified providing prerequisites listed above have been met.
  Of course the computers must be powered on in order to get Mac Addresses. 
  If a JSON file doesn't exist for the location and the SharedFolder path you specified in $locations, it will create one for you and attempt to populate it. 

  How it works: 
  It checks for the existence of a JSON file in the Location you provided and SharedFolder path you provided. 
     - If one is not found, it creates it. 
  It then grabs all computers from Active Directory based on the $OUPath for that $location.
  It pings every computer. If the computer is not responding, It removes it from the Mac Address Scan 
  If the computer is not already in the JSON file, it scans for the mac address and adds the new computer to the JSON file with a timestamp 
  If the computer is already in the JSON file, it updates time stamp and skips a check for the mac address. 
  If the Timestamp is older than 14 days, it prunes the Computer from the JSON File. 

  Usage Example: 
  'Get-Macs -location LCH'


### Function: Stop-Computers

  How it Works: 
  It first runs the Get-Macs function above in order to clean up/update the JSON File 
  It then sends a Shutdown command to each computer in the JSON File
  Usage Example: 
  'Stop-Computers -location SUL' 

### Function: Start-Computers 

  How it works:
  This Loads the JSON file. 
  For each computer in the JSON File, it sends a Wake-on-Lan packet. 
  If there is no JSON file found, it will complain and instruct you to turn on all the computers manually and run the Get-Macs Function.  
  Ex: 'Start-Computers -location LCH'


## Logging  
  I didn't get too fancy here.
  
  * Specify the $LogFilePath in the script. 
  * When Log File meets or exceeds 20MB(You can change this by setting $MaxFileSize), It's renamed with a .OLD extension. 
  * If there's already a LogFile with .OLD it deletes it. 
  The Log file will tell you:
  - A list of computers that were started and what Date/Time 
  - A list of computers that were stopped and what Date/Time 
  - When Getting Mac addresses... 
             - Which computers are in Active Directory, but not responding to a ping 
             - Which computers are in Active Directoy, responding to ping, and will be added to the JSON File 
             - Which computers are in JSON file, but didn't respond to a ping. 
             - If computer was removed because it was listed in the JSON file, but didn't respond for 14 days. 


## Considerations  

1. Once again, this script was made for a very specfic environment, but could be modified to suite your needs. 

2. The way this script gets MAC addresses assumes that computer's NIC is in the 'up' state and it's named "Ethernet"

       Invoke-Command -ComputerName $Computer -scriptblock {Get-NetAdapter | Where-Object {($_.Status -eq 'up' -and $_.Name -eq "Ethernet")} | Select-Object MacAddress}

  I know there's better ways you could be more resilient when getting the mac address. 
  For instance, you could get the Interface Index on the Computer that hase the default gateway set. (In most cases only one NIC per computer that has the default gateway set.)
  Then plug that InterfaceIndex to get the Interface Alias of the Nic with the Default Gateway Set
  Next, Run the Command to get the Mac Address of the $NicAlias, Like Below. I haven't tested that yet. 

    $defaultRouteNic = Get-NetRoute -DestinationPrefix 0.0.0.0/0 | Sort-Object -Property RouteMetric | Select-Object -ExpandProperty ifIndex 
    $NICAlias = Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex $defaultRouteNic | Select-Object -ExpandProperty InterfaceAlias
    Invoke-Command -ComputerName $Computer -scriptblock {Get-NetAdapter | Where-Object {($_.Status -eq 'up' -and $_.Name -eq $NicAlias)} | Select-Object MacAddress}


3. Retrieving the correct Broadcast Address to send the magic packet.
   Used to be, when sending magic packets using the broadcast address of 255.255.255.255 would work perfectly well, but this seemed stopped working in later version of Windows.
   Instead, you must contsrtuct the magic packet with the Network address. example 192.168.1.255. It has something to do with the way the latest Windows Network Stack works.
   If I recall, it seems that when using 255.255.255.255 as the broadcast address on a computer with multiple NICs, it does not decern which NIC to try to broadcast out from.
   Anyway, if someone can provide more insight on that, that would be great. Wireshark was very helpful in determining the issue. 
   
   There's a function in the script called Get-BroadCast Address. It does it's best to determine the network address of the network. 
   If you're having troubles with computers not waking up, investigate the broadcast address.  
    

