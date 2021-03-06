﻿
<#
.Synopsis
   Gets Username;Name;Surname;Department;Job title;Manager parameters for users in a specified country.
.DESCRIPTION
   Gets Username;Name;Surname;Department;Job title;Manager parameters for users in a specified country and
   outputs them to display. If parameter FilePath is specified it outpust them to specified file.
   Requires AD module installed to work.
.EXAMPLE
   Get-ADCountryUsers.ps1 -Country UA
.EXAMPLE
   Get-ADCountryUsers.ps1 -Country UA -FilePath "c:\temp\ttt.csv"
.PARAMETER Country
   ISO Alfa-2 Country code https://en.wikipedia.org/wiki/ISO_3166-1
.PARAMETER FilePath
   If this parameter is specified list of users will be ouput to file specified.
.NOTES
    This script requires AD module to function.
    Written By: Jan Bocko Kuhar
    Twitter:	http://twitter.com/JanBockoKuhar

    Change Log
    V0.01, 22.09.2015
#>
[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [Alias("CC")] 
    [string]
    $Country,

    [Parameter()]
    [AllowNull()]
    [string]
    $FilePath
)

Begin
{
    $TestADSnapin = get-pssnapin | where { $_.Name -eq "ActiveDirectory"}
    if($TestADSnapin -eq $null)
    {
        try
        {
            import-module ActiveDirectory -ErrorAction Stop
        }
        catch
        {
            Write-Warning "Could not import module Active Directory"
        }
    }
}
Process
{
    Add-Type -AssemblyName System.speech
    $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $Speak.SpeakAsync(("take me to your heaven" | Out-String)) | out-null
    $users = Get-ADUser -LDAPFilter "(c=$Country)" -Properties sn,department,title,manager
    $speak.dispose() | out-null
    if($FilePath)
    {
        Out-File -FilePath $FilePath -InputObject "Username;Name;Surname;Department;Job title;Manager"
    }
    foreach($user in $users)
    {
        $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
        $Speak.SpeakAsync(("Processing user $($user.name)" | Out-String)) | out-null
        $speak.dispose() | out-null
        if($FilePath)
        {
            out-file -FilePath $FilePath -InputObject "$($user.SamAccountName);$($user.GivenName);$($user.sn);$($user.Department);$($user.Title);$($user.Manager)" -Append
        }
        else
        {
            $param = @{
                'Username' = $user.SamAccountName
                'Name' = $user.GivenName
                'Surname' = $user.sn
                'Departnemt' = $user.Department
                'Job Title' = $user.Title
                'Manager' = $user.Manager
            }
            New-Object -TypeName psobject -Property $param
        }
    }
}
End
{
}
