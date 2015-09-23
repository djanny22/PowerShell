Function GET-JANCountryISOCode {
	Param ([string]$CountryCode)
	
	switch ($CountryCode) {
		AL {return 'Albania'}
		BA {return 'Bosnia'}
		BG {return 'Bulgaria'}
		CZ {return 'Czech'}
		EE {return 'Estonia'}
		GR {return 'Greece'}
		HU {return 'Hungary'}
		KO {return 'Kosovo'}
		KS {return 'Kosovo'}
		LV {return 'Latvia'}
		LT {return 'Lithuania'}
		MK {return 'Macedonia'}
		MD {return 'Moldova'}
		ME {return 'Montenegro'}
		PL {return 'Poland'}
		RO {return 'Romania'}
		RU {return 'Russia'}
		RS {return 'Serbia'}
		SK {return 'Slovakia'}
		SI {return 'Slovenia'}
		HR {return 'Croatia'}
		CH {return 'Switzerland'}
		TR {return 'Turkey'}
		UA {return 'Ukraine'}
		US {return 'United States'}
		GB {return 'United Kingdom'}
		CA {return 'Canada'}		
		
		default {return $CountryCode}
	}
}

Function Get-JANUMGroups {
	Param ([string]$CountryCode)
	
	switch ($CountryCode) {
		AL {return "SR.AL.ComputerManagement"}
		BA {return "SR.BA.ComputerManagement"}
		BG {return "SR.BG.ComputerManagement"}
		CZ {return "SR.CZ.ComputerManagement"}
		CH {return "SR.CH.ComputerManagement"}
		EE {return "SR.EE.ComputerManagement"}
		HR {return "SR.HR.ComputerManagement"}
		HU {return "SR.HU.ComputerManagement"}
		KO {return "SR.KO.ComputerManagement"}
		LT {return "SR.LT.ComputerManagement"}
		LV {return "SR.LV.ComputerManagement"}
		MD {return "SR.MD.ComputerManagement"}
		ME {return "SR.ME.ComputerManagement"}
		MK {return "SR.MK.ComputerManagement"}
		PL {return "SR.PL.ComputerManagement"}
		RO {return "SR.RO.ComputerManagement"}
		RS {return "SR.RS.ComputerManagement"}
		RU {return "SR.RU.ComputerManagement"}
		SI {return "SR.SI.ComputerManagement"}
		SK {return "SR.SK.ComputerManagement"}
		TR {return "SR.TR.ComputerManagement"}
		UA {return "SR.UA.ComputerManagement"}
		GB {return "SR.GB.ComputerManagement"}
		CA {return "SR.CA.ComputerManagement"}
		US {return "SR.US.ComputerManagement"}
		SMFG {return "SR.SMFG.ComputerManagement"}
		default {return ''}
	}
}

Function Test-JANComputerNameStructure{
<#
.SYNOPSIS
Test if computer name is following enterprise naming convention.
.DESCRIPTION
Test computer name is following naming convention:
(CC/CL/NL)-CC-X0****
Returns 0 if computer name is constructed correctly, 1 if CC is 
incorrect and 2 if computer name does not begin with either
CC, CL or NL.
.PARAMETER ComputerName
One computer name.
.EXAMPLE
Test-JANComputerNameStructure -Computername 'CL-SI-X01234'
#>
    Param ([string]$ComputerName)

    $parts = $ComputerName.split('-')
    $return | Out-Null
    if (($parts[0] -eq 'CC') -or ($parts[0] -eq 'NL') -or ($parts[0] -eq 'CL')){
        if((GET-JANCountryISOCode -CountryCode $parts[1]) -ne $parts[1]){
            $return = 0
        }
        else{
            # Computer Name CC Incorrect
            $return = 1
        }
    }
    else{
        # Computer Name Incorrect
        $return = 2
    }
    $return
}

Function Confirm-JANUserPermissions{
<#
.SYNOPSIS
Check if user has permissions to manage this computer, by computer
name.
.DESCRIPTION
Checks in AD if user is member of group that has permissions to manage 
computer by computer name.
Requires ActiveDirectory module to work.
Returns 0 if user has permissions and 1 if user does not have permissions
to manage this computer.
.PARAMETER ComputerName
One computer name.
.PARAMETER UserName
User name of user managing computer in question.
.EXAMPLE
Confirm-JANUserPermissions -Computername 'CL-SI-X01234' -UserName 'domain\name.surname'
#>
    Param (
    [string]$ComputerName,
    [string]$UserName
    )
    Import-Module ActiveDirectory

    $return | Out-Null
    $parts = ($ComputerName.Split('-'))
    $cc = $parts[1]
    $group = Get-JANUMGroups -CountryCode $cc
    $User = ($UserName.Split('\'))[1]
    $users = Get-ADGroupMember -Identity $group -Recursive
    if($User -in $users.SamAccountName){
        $return = 0
    }
    else{
        # User does not have permissions.
        $return = 1
    }
    Remove-Module ActiveDirectory -ErrorAction SilentlyContinue
    $return
}

Function Confirm-JANADComputer{
<#
.SYNOPSIS
Check if computer allready exists in AD.
.DESCRIPTION
Checks in AD if computer with this name allready exists. If it exists it cannot be
added again with the same name. If you are trying to change existing computer, you 
must use Modify-JANComputer.
Requires ActiveDirectory module to work.
.PARAMETER ComputerName
One computer name.
.EXAMPLE
Confirm-JANADComputer -Computername 'CL-SI-X01234'
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

    $SiteServer = "SR-SI-SCCMPS1-P"
    $SiteCode = "SI5"
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
New-JANMDTComputer -ComputerName 'CL-SI-X01234' -MACAddress 'AA:BB:CC:DD:EE:FF' -CompanyCode 'SI01' -SerialNumber 'S345782'
#>
    param
    (
    [string]$ComputerName, 
    [string]$MACAddress, 
    [string]$CompanyCode, 
    [string]$SerialNumber
    )
    Import-Module 'C:\scripts\SCCM Scripts\MDTDB.psm1'

    $NEW_COMPUTER_OU = "OU=Test,OU=Studio moderna,DC=sm-group,DC=local"
    $mdtdbsrv2012 = "CS-HQ-SQL10-P"
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
        $return = 'SI50029B'
    }
    elseif($OperatingSystem -eq 'Windows 7')
    {
        $return = 'SI500297'    
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

    $SiteServer = "SR-SI-SCCMPS1-P"
    $SiteCode = "SI5"

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

    $SiteServer = "SR-SI-SCCMPS1-P"
    $SiteCode = "SI5"

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
New-JANComputer
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
        [Alias('OS')]
        [string[]]$OperatingSystem,

        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   HelpMessage="Computer MAC address")]
        [Alias('MAC')]
        [string[]]$MACAddress,

        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   HelpMessage="User name")]
        [Alias('User')]
        [string[]]$UserName,

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
        if((Test-JANComputerNameStructure -ComputerName $ComputerName) -eq 0)
        {
            if((Confirm-JANUserPermission -UserName $UserName -ComputerName $ComputerName) -eq 0)
            {
                if(Confirm-JANADComputer -ComputerName $ComputerName)
                {
                    New-JANSCCMComputer -ComputerName $ComputerName -MACAddress $MACAddress
                    New-JANMDTComputer -ComputerName $ComputerName -MACAddress $MACAddress -CompanyCode $CompanyCode -SerialNumber $SerialNumber
                    Add-JANComputerToCollection -OperatingSystem $OperatingSystem -ComputerName $ComputerName
                    Update-JANCollection($OperatingSystem)
                }
            }
            else{}
        }
        else{}
    }
    END{}
}

Export-ModuleMember -Function New-JANComputer