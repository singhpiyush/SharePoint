# #############################################################################
# NAME: Export-TermSets.ps1
# 
# AUTHOR:  Piyush Kumar Singh
# DATE:  2015/06/19
# EMAIL: piyushksingh11@gmail.com
# 
# COMMENT:  This script will....
#
# VERSION HISTORY
# 1.0 Initial Version.
#
# TO ADD
# -Implement it for SharePoint Online as well
# #############################################################################

$termGrpNm = "TEST"
$svLoc = "C:\Piyush\ChronicleMigration\TermSet` Operations\OrderDeprecateTerms\TermFolder"

[Byte[]] $amp = 0xEF,0xBC,0x86         # Ampersands are stored as fullwidth ampersands (see http://www.fileformat.info/info/unicode/char/ff06/index.htm)
$commaEncd = "&#44;"                   # Text to be used in place of commam as the CSV file's delimiter is ','.

Function ReplaceChars($stringName)
{
	return ($stringName.Replace([System.Text.Encoding]::UTF8.GetString($amp), "&").Replace(",", $commaEncd))
}

function Add-Snapin {
	if ((Get-PSSnapin -Name Microsoft.Sharepoint.Powershell -ErrorAction SilentlyContinue) -eq $null) {
		$global:SPSnapinAdded = $true
		Write-Host "Adding Sharepoint module to PowerShell" -NoNewline
		Add-PSSnapin Microsoft.Sharepoint.Powershell -ErrorAction Stop
		Write-Host " - Done."
	}
	
	Write-Host "Adding Microsoft.Sharepoint assembly" -NoNewline
	Add-Type -AssemblyName "Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
	# Disable the above line and enable the line below for SharePoint 2013
	# Add-Type -AssemblyName "Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
	Write-Host " - Done."
}

function Remove-Snapin {
	if ($global:SPSnapinAdded -eq $true) {
		Write-Host "Removing Sharepoint module from PowerShell" -NoNewline
		Remove-PSSnapin Microsoft.Sharepoint.Powershell -ErrorAction SilentlyContinue
		Write-Host " - Done."
	}
}
function Export-SPTerms {
    param (
        [string]$termGroupName = $(Read-Host -prompt "Please provide the term group name to export"),
        [string]$saveLocation = $(Read-Host -prompt "Please provide the path of the folder to save the CSV file to")
    )
	
	if ([IO.Directory]::Exists($saveLocation) -eq $false)
	{
		New-Item ($saveLocation) -Type Directory | Out-Null
	}
	
	#Connect to Central Admin
	$caWebApp = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local
	$CAsiteCollectionUrl = $caWebApp.Sites[0].Url
	$CAsite = Get-SPSite $CAsiteCollectionUrl
	write-host "Connection made with Central Admin -"$CAsite
	
	$taxonomySession = Get-SPTaxonomySession -site $CAsite
	$taxonomyTermStore =  $taxonomySession.TermStores | Select Name
	$termStore = $taxonomySession.TermStores[$taxonomyTermStore.Name]
	
	foreach ($group in $termStore.Groups) {
	
		if ($group.Name -eq $termGroupName) {
		
		    foreach ($termSet in $group.TermSets) {
            	# Remove unsafe file system characters from filename
				$parsedFilename =  [regex]::replace($termSet.Name, "[^a-zA-Z0-9\\-]", "_")
				$file = New-Object System.IO.StreamWriter($saveLocation + "\termset_" + $parsedFilename + ".csv")
                
		        # Write out the headers
		        $file.Writeline("Term Set Name,Term Set Description,LCID,Available for Tagging,Term Description,Deprecated,Level 1 Term, Level 2 Term,Level 3 Term,Level 4 Term,Level 5 Term,Level 6 Term,Level 7 Term")
		        
				try {
					Export-SPTermSet $termSet.Terms
				}
				finally {
			        $file.Flush()
			        $file.Close()
				}
			}
		}
	}
}

function Export-SPTermSet {
    param (
        [Microsoft.SharePoint.Taxonomy.TermCollection]$terms,
		[int]$level = 1,
		[string]$previousTerms = ""
    )
	
	if ($level -ge 1 -or $level -le 7)
	{
		if ($terms.Count -gt 0 ) {
			$termSetName = ""
			if ($level -eq 1) {
				#$termSetName =  """" + $terms[0].TermSet.Name.Replace([System.Text.Encoding]::UTF8.GetString($amp), "&") + """"
                $termSetName =  """" + (ReplaceChars $terms[0].TermSet.Name) + """"
			}
			$terms | ForEach-Object {
				#$currentTerms = $previousTerms + ",""" + $_.Name.Replace([System.Text.Encoding]::UTF8.GetString($amp), "&") + """";
                $currentTerms = $previousTerms + ",""" + (ReplaceChars $_.Name) + """";
				
				$file.Writeline($termSetName +
								",""" + $_.TermSet.Description + """" + 
								"," + $_.ID +
								"," + $_.IsAvailableForTagging +
								"," + $_.GetDescription() + 
								"," + $_.IsDeprecated + $currentTerms);
				
				if ($level -lt 7) {
					#Export-SPTermSet $_.Terms ($level + 1) ($previousTerms + $currentTerms)
					Export-SPTermSet $_.Terms ($level + 1) $currentTerms
				}
			}
		}
	}
}

try {
	Write-Host "Starting export of Metadata Termsets" -ForegroundColor Green
	$ErrorActionPreference = "Stop"
	Add-Snapin
	
	Export-SPTerms $termGrpNm $svLoc
}
catch {
	Write-Host ""
    Write-Host "Error : " $Error[0] -ForegroundColor Red
	throw
}
finally {
	Remove-Snapin
}
Write-Host Finished -ForegroundColor Blue
