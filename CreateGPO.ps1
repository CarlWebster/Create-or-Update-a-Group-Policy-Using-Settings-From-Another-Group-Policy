#requires -Module GroupPolicy

#region help text

<#
.SYNOPSIS
	Creates an unlinked Group Policy Object from an imported CSV file.
	
	If the Group Policy Object exists, linked or unlinked, it is updated with the 
	settings from the imported CSV file.
	
.DESCRIPTION
	Creates an unlinked Group Policy Object from an imported CSV file.
	
	The CSV file is created using Darren Mar-Elia's Registry .Pol Viewer Utility
	https://sdmsoftware.com/gpoguy/gpo-freeware-registration-form/

	The Registry.pol files only contain settings from the Administrative Templates
	node in the Group Policy Management Console.
	
	This script is designed for consultants and trainers who may create Group Policies 
	in a lab and need a way to recreate those policies at a customer or training site.
	
	The GroupPolicy module is required for the script to run.
	
	On a server, Group Policy Management must be installed.
	On a workstation, RSAT is required.
	
	Remote Server Administration Tools for Windows 7 with Service Pack 1 (SP1)
		http://www.microsoft.com/en-us/download/details.aspx?id=7887
		
	Remote Server Administration Tools for Windows 8 
		http://www.microsoft.com/en-us/download/details.aspx?id=28972
		
	Remote Server Administration Tools for Windows 8.1 
		http://www.microsoft.com/en-us/download/details.aspx?id=39296
		
	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520
	
.PARAMETER PolicyName
	Group Policy Name.
	
	This is a required parameter.
	
	This parameter has an alias of PN.
	
.PARAMETER PolicyType
	The type of Group Policy settings to import.
	M - Machine (Computer)
	C - Machine (Computer)
	U - User
	
	This is a required parameter.
	
	This parameter has an alias of PT.
	
	Default is Machine.
	
	If a Group Policy is to contain both Computer and User settings, there will be
	two .Pol files. One for Machine (Computer) settings and a separate .Pol file 
	for User settings. Each .Pol file is named Registry.pol. The CSV files will need
	to be have different names. There will then be two CSV files to import.
	
	The script will detect if the Group Policy Name exists and if it does, the next
	set of settings will be added to the existing script.
	
	If the Group Policy Name does not exist, a new Group Policy is created with the 
	imported settings.
.PARAMETER CSVFile
	CSV file created by Registry PolViewer utility.
	
	This is a required parameter.
	
	The parameter has an alias of CSV.
	
	The CSV file will contain either Machine or User policy settings. It should
	contain both.
	
.EXAMPLE
	PS C:\PSScript > .\CreateGPO.ps1 -PolicyName "Sample Policy Name" -PolicyType M -CSVFIle C:\PSScript\GPOMSettings.csv
	
	Darren's PolViewer.exe utility is used first to create the 
	C:\PSScript\GPOMSettings.csv file.
	
	If a Group Policy named "Sample Policy Name" does not exist, a policy will be 
	created using the Machine settings imported from the GPOMSettings.csv file.
	
	If a Group Policy named "Sample Policy Name" exists, the policy will be 
	updated using the Machine settings imported from the GPOMSettings.csv file.

.EXAMPLE
	PS C:\PSScript > .\CreateGPO.ps1 -PolicyName "Sample Policy Name" -PolicyType U -CSVFIle C:\PSScript\GPOUSettings.csv
	
	Darren's PolViewer.exe utility is used first to create the 
	C:\PSScript\GPOUSettings.csv file.
	
	If a Group Policy named "Sample Policy Name" does not exist, a policy will be 
	created using the User settings imported from the GPOUSettings.csv file.
	
	If a Group Policy named "Sample Policy Name" exists, the policy will be 
	updated using the User settings imported from the GPOUSettings.csv file.

.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates or updates a Group Policy Object.
.NOTES
	NAME: CreateGPO.ps1
	VERSION: 1.00
	AUTHOR: Carl Webster
	LASTEDIT: March 18, 2016
#>

#endregion

#region script parameters

[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Machine") ]

Param(
	[parameter(ParameterSetName="Machine",Mandatory=$True)] 
	[parameter(ParameterSetName="User",Mandatory=$True)] 
	[Alias("PN")]
	[ValidateNotNullOrEmpty()]
	[String]$PolicyName,
    
	[parameter(ParameterSetName="Machine",Mandatory=$True,
	HelpMessage="Enter M or C for Machine or U for User")] 
	[parameter(ParameterSetName="User",Mandatory=$True,
	HelpMessage="Enter M or C for Machine or U for User")] 
	[Alias("PT")]
	[String][ValidateSet("M", "C", "U")]$PolicyType="M",

	[parameter(ParameterSetName="Machine",Mandatory=$True,
    HelpMessage="Enter path to CSV file.")] 
	[parameter(ParameterSetName="User",Mandatory=$True, 
    HelpMessage="Enter path to CSV file.")] 
	[Alias("CSV")]
	[ValidateNotNullOrEmpty()]
	[String]$CSVFile

	)
#endregion

#region script change log	

#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on November 7, 2015

#released to the community on 18-Mar-2016

#endregion

#region initial variable testing and setup
Set-StrictMode -Version 2

If(!(Test-Path $CSVFile))
{
	Write-Error "$(Get-Date): CSV file $($CSVFile) does not exist.`n`nScript cannot continue"
	Exit
}

#endregion

#region functions

Function SetPolicySetting
{
	Param([string]$Key, [string]$ValueName, [string]$ValueType, $Data)
	
	If($Key -eq "" -or $ValueName -eq "" -or $ValueType -eq "" -or $Null -eq $Key -or $Null -eq $ValueName -or $Null -eq $ValueType)
	{
		#failure
		Write-Warning "$(Get-Date): Missing data: `nKey: $($Key) `nValueName: $($ValueName) `nValueType: $($ValueType) `nData: $($Data)`n`n"
	}
	Else
	{

		#Specifies the data type for the registry-based policy setting. You can specify one of the following data
		#types: String, ExpandString, Binary, DWord, MultiString, or Qword.
		$NewType = ""
		$NewData = ""
		Switch ($ValueType)
		{
			"REG_DWORD"		{$NewType = "DWord"; $NewData = Invoke-Expression ("0x" + $Data)}
			"REG_SZ"		{$NewType = "String"; $NewData = $Data}
			"REG_EXPAND_SZ"	{$NewType = "ExpandString"; $NewData = $Data}
			"REG_BINARY" 	{$NewType = "Binary"; $NewData = $Data}
			"REG_MULTI_SZ" 	{$NewType = "MultiString"; $NewData = $Data}
			"REG_QDWORD" 	{$NewType = "QWord"; $NewData = Invoke-Expression ("0x" + $Data)}
		}

		If($PolicyType -eq "M" -or $PolicyType -eq "C")
		{
			$NewKey = "HKLM\$($Key)"
		}
		ElseIf($PolicyType -eq "U")
		{
			$NewKey = "HKCU\$($Key)"
		}
		
		$results = Set-GPRegistryValue -Name $PolicyName -key $NewKey -ValueName $ValueName -Type $NewType -value $NewData -EA 0
		
		If($? -and $Null -ne $results)
		{
			#success
			Write-Host "$(Get-Date): Successfully added: $($NewKey) $($ValueName) $($NewType) $($NewData)"
		}
		Else
		{
			#failure
			Write-Warning "$(Get-Date): Problem adding: $($NewKey) $($ValueName) $($NewType) $($NewData)"
		}
	}
}

#endregion

Write-Host "Processing $($PolicyName) Group Policy Object"

$results = Get-GPO -Name $PolicyName -EA 0

If($? -and $results -ne $Null)
{
	#gpo already exists
	Write-Host
	Write-Host "$(Get-Date): $($PolicyName) exists and will be updated"
	Write-Host
}
ElseIf(!($?) -and $results -eq $Null)
{
	#gpo does not exist. Create it
	$results = New-GPO -Name $PolicyName -EA 0
	
	If($? -and $results -ne $null)
	{
		#success
		Write-Host
		Write-Host "$(Get-Date): $($PolicyName) does not exist and will be created"
		Write-Host
	}
	Else
	{
		#something went wrong
		Write-Error "$(Get-Date): Unable to create a GPO named $($PolicyName).`n`nScript cannot continue"
		Exit
	}
}
Else
{
	#error
	Write-Error "$(Get-Date): Error verifying if GPO $($PolicyName) exists or not.`n`nScript cannot continue"
	Exit
}

Write-Host "$(Get-Date): Read in CSV file"
Write-Host
$Settings = Import-CSV $CSVFile

If($? -and $Settings -ne $Null)
{
	#success
	Write-Host "$(Get-Date): Setting policy settings"
	Write-Host

	ForEach($Setting in $Settings)
	{
		SetPolicySetting $Setting.'Registry Key' $Setting.'Registry Value' $Setting.'Value Type' $Setting.Data
	}
}
ElseIf($? -and $Settings -eq $Null)
{
	#success but no contents
	Write-Warning "$(Get-Date): CSV file $($CSVFile) exists but has no contents.`n`nScript cannot continue"
	Exit
}
Else
{
	#failure
	Write-Error "$(Get-Date): Error importing CSV file $($CSVFile).`n`nScript cannot continue"
	Exit
}

Write-Host
Write-Host "$(Get-Date): Successfully created/updated GPO $($PolicyName)"
Write-Host
Get-GPO -Name $PolicyName

