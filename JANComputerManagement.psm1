Function Confirm-JANADComputer{
<#
.SYNOPSIS
Check if computer allready exists in AD.
.DESCRIPTION
Checks in AD if computer with this name already exists. If it exists it cannot be
added again with the same name. If you are trying to change existing computer, you 
must use Modify-JANComputer.
Requires ActiveDirectory module to work.
.PARAMETER ComputerName
One computer name.
.EXAMPLE
Confirm-JANADComputer -Computername 'ComputerName'
#>
    Param (
    [string]$ComputerName
    )

    Import-Module ActiveDirectory

    $return | Out-Null

    try{
        Get-ADComputer -Identity $ComputerName | Out-Null
        $return = 1
    }
    catch{
        $return = 0
    }
    Remove-Module ActiveDirectory -ErrorAction SilentlyContinue
    $return
}

Function New-JANSCCMComputer{
<#
.SYNOPSIS
Creates new object in SCCM using WMI. If record already exists, it will be overwritten.
.DESCRIPTION
Creates new object in SCCM using WMI. If record already exists, it will be overwritten.
You must change SiteServer and SiteCode names.
.PARAMETER ComputerName
One computer name.
.PARAMETER MACAddress
Computers MAC Address. Must be MAC address of the card you will be booting computer
from, wired network adapter.
.EXAMPLE
New-JANSCCMComputer -ComputerName 'TestCName' -MACAddress 'AA:BB:CC:DD:EE:FF'
#>
    param(
    [string]$ComputerName, 
    [string]$MACAddress
    )

    $SiteServer = 'YourSiteServer'
    $SiteCode = 'YourSiteCode'
    try
    {
        #New computer account information
        $WMIConnection = ([WMIClass]"\\$SiteServer\root\SMS\Site_$($SiteCode):SMS_Site")
        $NewEntry = $WMIConnection.psbase.GetMethodParameters("ImportMachineEntry")
        $NewEntry.MACAddress = $MACAddress
        $NewEntry.NetbiosName = $ResourceName
        $NewEntry.OverwriteExistingRecord = $True
        $Resource = $WMIConnection.psbase.InvokeMethod("ImportMachineEntry",$NewEntry,$null)
    }
    catch{
        #FAILED! No check if adding to SCCM failed....for now.
    }
}

Function New-JANMDTComputer{
<#
.SYNOPSIS
Creates new object in SCCM. If record already exists, it will be overwritten.
Requires MDT module.
.DESCRIPTION
Creates new object in SCCM. If record already exists, it will be overwritten.
You must change SiteServer and SiteCode names.
.PARAMETER ComputerName
One computer name.
.PARAMETER MACAddress
Computers MAC Address. Must be MAC address of the card you will be booting computer
from, wired network adapter.
.PARAMETER CompanyCode
Company code to which this computer belongs to, MDT field OrgName
.PARAMETER SerialNumber
Computers serial number, MDT SerialNumber identifyer
.EXAMPLE
New-JANMDTComputer -ComputerName 'ComputerName' -MACAddress 'AA:BB:CC:DD:EE:FF' -CompanyCode 'OrgName' -SerialNumber 'S123456'
#>
    param
    (
    [string]$ComputerName, 
    [string]$MACAddress, 
    [string]$CompanyCode, 
    [string]$SerialNumber
    )
    Import-Module 'C:\scripts\SCCM Scripts\MDTDB.psm1'

    $NEW_COMPUTER_OU = "OU=Test,DC=Contoso,DC=com"
    $mdtdbsrv2012 = 'ndtDBserver'
    $mdtdb = "MDTConfigDB"

    try
    {
        $tt = Connect-MDTDatabase –sqlServer $mdtdbsrv2012 –database MDTConfigDB
        #Test if computer exists and clean beforehand!
        $settings = @{
			        OSInstall='YES';
			        OSDComputerName=$ComputerName;
			        OrgName=$CompanyCode;
			        OSDINSTALLSILENT='1';
			        Location=MDTCountries ($ComputerName.Split('-'))[1];
			        MachineObjectOU=$NEW_COMPUTER_OU
		        }
    
        New-MDTComputer -Description $ComputerName -SerialNumber $SerialNumber -MACAddress $MACAddress –Settings $settings | Out-Null
        $return = 0
    }

    Catch
    {
        $return = 1
    }
    Remove-Module 'C:\scripts\SCCM Scripts\MDTDB.psm1' -ErrorAction SilentlyContinue
    $return
}

Function Get-JANCollIDfromOSNameNEW{
<#
.SYNOPSIS
Transforms operating system name to collection ID for OSD New.
.DESCRIPTION
Transforms operating system name to collection ID for OSD New. Returns 
collection ID for requested OSD and OS type.
.PARAMETER OperatingSystem
Operating system name. Valid values are "Windows XP", "Windows 7"
.EXAMPLE
Get-JANCollIDfromOSName -OperatingSystem 'Windows 7'
#>
    [CmdletBinding()]
    param
    (
    [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   HelpMessage="Operating system:")]
    [ValidateSet("Windows XP", "Windows 7")] 
    [string]$OperatingSystem
    )

    $return = 0
    if($OperatingSystem -eq 'Windows XP')
    {
        $return = 'Collection ID'
    }
    elseif($OperatingSystem -eq 'Windows 7')
    {
        $return = 'Collection ID'    
    }
    else
    {
        $return = 1
    }
}

function Add-JANComputerToCollection{
<#
.SYNOPSIS
Add computer to collection for OS deploy based on desired operating system.
Requires SCCM module installed.
.DESCRIPTION
Add computer to collection for OS deploy based on desired operating system.
Requires SCCM module installed.
.PARAMETER ComputerName
One computer name.
.PARAMETER OperatingSystem
Operating system to be isntalled on computer.
.EXAMPLE
New-JANSCCMComputer -ComputerName 'TestCName' -OperatingSystem 'Windows 7'
#>
    param
    (
    [string]$OperatingSystem, 
    [string]$ComputerName
    )

    Import-Module "C:\Program Files (x86)\Microsoft Configuration Manager Console 2012\bin\ConfigurationManager.psd1"

    $SiteServer = 'YourSiteServer'
    $SiteCode = 'YourSiteCode'

    # Switch to CM Drive
    cd $SiteCode':'

    $CollectionID = Get-JANCollIDfromOSNameNEW -OperatingSystem $OperatingSystem

    Add-CMDeviceCollectionDirectMembershipRule -CollectionId $CollectionID -ResourceId $ResourceID
}

function Update-JANCollection
{
<#

#>
    param
    (
    [string]$OperatingSystem
    )

    $SiteServer = 'YourSiteServer'
    $SiteCode = 'YourSiteCode'

    $CollectionID = Get-JANCollIDfromOSNameNEW -OperatingSystem $OperatingSystem

    $Return = Invoke-WmiMethod -Path "ROOT\SMS\Site_$($SiteCode):SMS_Collection.CollectionId='$CollectionId'" -Name RequestRefresh -ComputerName $SiteServer

}

function New-JANComputer {
<#
.SYNOPSIS
Create new computer in SCCM, MDT
.DESCRIPTION
New-JANComputer creates new computer in SCCM and MDT. It also
checks whether user has permissions to do this action. It is to
be used with SCSM or simmilar system.
.PARAMETER ComputerName
One computer name.
.PARAMETER SerialNumber
SN for this computer
.PARAMETER CompanyCode
CC Reason code for this computer to be filled in BIOS
.PARAMETER OperatingSystem
Desired operating system to be isntalled on computer, desides collection
in SCCM for OSD.
.PARAMETER MACAddress
Computers MAC address. Must be MAC address of the card you will be booting 
computer from, wired network adapter.
.PARAMETER UserName
User name of user managing computer in question
.PARAMETER LogErrors
Specify this switch to create a text log file of computers
that could not be queried.
.PARAMETER ErrorLog
When used with -LogErrors, specifies the file path and name
to which failed computer names will be written. Defaults to
C:\Retry.txt.
.EXAMPLE
 Get-Content  | New-JANComputer
.EXAMPLE
New-JANComputer -ComputerName "ComputerName" -SerialNumber "S123456" -CompanyCode "OrgName" -OperatingSystem "Windows 7" -MACAddress 'AA:BB:CC:DD:EE:FF' 
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   HelpMessage="Computer name or IP address")]
        [Alias('hostname')]
        [string[]]$ComputerName,

        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   HelpMessage="Computer serial number")]
        [Alias('SN')]
        [string[]]$SerialNumber,

        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   HelpMessage="Company Code for computer")]
        [Alias('CCReason')]
        [string[]]$CompanyCode,

        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   HelpMessage="Desired operating system")]
        [ValidateSet("Windows XP", "Windows 7")] 
        [Alias('OS')]
        [string[]]$OperatingSystem,

        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   HelpMessage="Computer MAC address")]
        [Alias('MAC')]
        [string[]]$MACAddress,

        [string]$ErrorLog = $JANErrorLogPreference,

        [switch]$LogErrors
    )
    BEGIN {
        if ($LogErrors){
            Write-Verbose "Error log will be $ErrorLog"
        }
        else{
            Write-Verbose "No Error logging."
        }


    }
    PROCESS {
        New-JANSCCMComputer -ComputerName $ComputerName -MACAddress $MACAddress
        New-JANMDTComputer -ComputerName $ComputerName -MACAddress $MACAddress -CompanyCode $CompanyCode -SerialNumber $SerialNumber
        Add-JANComputerToCollection -OperatingSystem $OperatingSystem -ComputerName $ComputerName
        Update-JANCollection($OperatingSystem)

    }
    END{}
}

Export-ModuleMember -Function New-JANComputer