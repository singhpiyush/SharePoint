# #############################################################################
# NAME: Import-TermSets.ps1
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

$fileName = "D:\Piyush\TermTest.csv"
$termStoreGroupName = "TEST1"
$isTermAvailable = $true
$lcid = 2
$availableId = 3
$descId = 4
$deprecateId = 5
$fstTrmId = 6
$delimiter = ','
$nameDelimiter = "~~"

$dictTerms = @{}
$dictGuids = @{}

$snapinName = "Microsoft.SharePoint.PowerShell"
if ((Get-PSSnapin | Where-Object {$_.Name -eq $snapinName }) -eq $NULL) {
  write-host "SharePoint SnapIn not loaded. Loading..."
  Add-PSSnapin $snapinName -ErrorAction SilentlyContinue
}

#Connect to Central Admin
$caWebApp = [Microsoft.SharePoint.Administration.SPAdministrationWebApplication]::Local
$CAsiteCollectionUrl = $caWebApp.Sites[0].Url
$CAsite = Get-SPSite $CAsiteCollectionUrl
write-host "Connection made with Central Admin -"$CAsite

#Connect to Term Store in the Managed Metadata Service Application
$taxonomySession = Get-SPTaxonomySession -site $CAsite
$MMSStore = "Managed Metadata Service"

$termStore = $taxonomySession.TermStores[$MMSStore]
write-host "Connection made with term store -"$termStore.Name

#Get TermGroup and TermSet
$termStoreGroup = $termStore.Groups[$termStoreGroupName]

function ClearAll()
{
	$dictTerms.Clear()
	$dictGuids.Clear()

	$CAsite.Dispose()
	$taxonomySession = $null
	$termStore = $null
}

function ApplyOrdering
{
	$isCommit = $false
	foreach($key in $dictGuids.Keys)
	{ 
		$dictTerms.Get_Item($key).CustomSortOrder = $dictGuids.Get_Item($key)

		if($isCommit -eq $false)
		{
			$isCommit = $true
		}
	}

	if($isCommit -eq $true)
	{
		$termStore.CommitAll()
	}
}

function CacheTermOrder($ordTerm, $term)
{
	if($ordTerm -ne $null)
	{
		Write-Host ("Ord " + $ordTerm.Name)
		Write-Host ("Exst " + $term.Name)

		$trmId = [string]$term.Id
		if($dictGuids.ContainsKey($trmId))
		{
			$existId = $dictGuids.Get_Item($trmId)
			$newId = [string]$ordTerm.Id
			if($existId.Contains($newId) -eq $false)
			{
				$dictGuids.Set_Item($trmId, ($existId + ":" + $newId))
			}
		}
		else
		{
			$dictGuids.Add($trmId, ([string]$ordTerm.Id))
			$dictTerms.Add($trmId, $term)
		}
	}
}

function DeprecateEnableTerms($term, $termDetail)
{
	$isCommit = $false
	if(($term.IsAvailableForTagging -eq $true) -and ($termDetail[$availableId] -eq "FALSE"))
	{
		$term.IsAvailableForTagging = $false
		$isCommit = $true
	}
	if(($term.IsAvailableForTagging -eq $false) -and ($termDetail[$availableId] -eq "TRUE"))
	{
		$term.IsAvailableForTagging = $true
		$isCommit = $true
	}
	if(($term.IsDeprecated -eq $true) -and ($termDetail[$deprecateId] -eq "FALSE"))
	{
		$term.Deprecate($false)
		$isCommit = $true
	}
	if(($term.IsDeprecated -eq $false) -and ($termDetail[$deprecateId] -eq "TRUE"))
	{
		Write-Host "Deprecating..."

		$term.Deprecate($true)
		$isCommit = $true
	}
	if($isCommit -eq $true)
	{
		Write-Host "Commit"

		$termStore.CommitAll()
		Write-Host "Deprecate"
	}
}

function GetNewName($termSetName)
{
	$splitNameNew =  if(([string[]]$termSetName).Length -eq 1){ ([string[]]$termSetName)[0] } else { ([string[]]$termSetName)[1] }
	Write-Host ("Afetr Split " + $splitNameNew)

	return $splitNameNew
}

function GetNamePair($name)
{
	return ($name.Split($nameDelimiter, [StringSplitOptions]::RemoveEmptyEntries))
}

function IsAvailableForTagging($termVal)
{
	Write-Host($termVal)

	if($termVal)
	{
		$termVal = $termVal.Replace("`"", "").Trim() #"
	}

	if ($termVal -eq "FALSE")
	{ 
		Write-Host("Is False")

		$isAvailable = $FALSE
	} 
	else 
	{ 
		$isAvailable = $TRUE
	}
}

function CreateTerm($termName, $termDesc, $termSet, $isAvailable, $termGuid)
{
	write-host ("Term not Found -" + $termName + " " + $isAvailable)


	if(([string]::IsNullOrEmpty($termGuid)) -eq $false)
	{
		Write-Host ($termName + $termGuid)

		$term = $termSet.CreateTerm($termName, 1033, ([GUID]$termGuid))
	}
	else
	{
		$term = $termSet.CreateTerm($termName, 1033)
	}
	$term.IsAvailableForTagging = $isAvailable

	if($termDesc)
	{
		$termDesc = $termDesc.Replace("`"", "").Trim() #"
	}

	$term.SetDescription($termDesc, 1033)
	#$term.Description = "$termDesc.Replace("`"", "").Trim()"

	$termStore.CommitAll()

	return $term
}

function ManageTerm($fstTrmId, $termDetail, $termSet)
{
	$newTerm = $null

	++$fstTrmId
	if(($fstTrmId -lt $termDetail.Length) -and ($termDetail[$fstTrmId]))
	{
		$newTermNames = GetNamePair ($termDetail[$fstTrmId].Replace("`"", "").Trim()) #"
		Write-Host ([string[]]$newTermNames)[0]

		$newTerm = $termSet.Terms[([string[]]$newTermNames)[0]]
		if($newTerm -eq $null)
		{      
			$isAvailable = $TRUE
			$newTermName =  GetNewName $newTermNames

			Write-Host $newTermName
			Write-Host "New Name ............................"
			Write-Host $termSet.Name

			$newTerm = $termSet.Terms[$newTermName]
			if($newTerm -eq $null)
			{
				if ($termDetail[$availableId] -eq "FALSE")
				{ 
					 Write-Host("Is False")
					 
					 $newTerm = CreateTerm $newTermName $termDetail[$descId] $termSet $FALSE $termDetail[$lcid]
				} 
				else 
				{ 
					$newTerm = CreateTerm $newTermName $termDetail[$descId] $termSet $TRUE $termDetail[$lcid]
				}
			}

		}
		elseif (([string[]]$newTermNames).Length -eq 2)
		{
			$newTerm.Name = $newTermNames[1]
			$termStore.CommitAll()
		}

		if([string]::IsNullOrEmpty($termDetail[($fstTrmId + 1)]))
		{
			DeprecateEnableTerms $newTerm $termDetail
		}

		ManageTerm $fstTrmId $termDetail $newTerm


		CacheTermOrder $newTerm $termSet
	}
	else
	{
		Write-Host("Not ENTERED")
	}
 
}

function ManageTerms()
{
	$termData = Get-Content $fileName
	$termSetName = GetNamePair ($termData[1].Split($delimiter)[0].Replace("`"", "").Trim()) #"
	$termSetDesc = $termData[1].Split($delimiter)[1].Replace("`"", "").Trim() #"

	write-host ("Creating TermSet " + ([string[]]$termSetName)[0]) -backgroundcolor blue

	if ($termStoreGroup -ne $null)
	{
		write-host "TermGroup Found -"$termStoreGroupName
		$termSet = $termStoreGroup.TermSets[([string[]]$termSetName)[0]]
		if ($termSet -eq $null)
		{
			$termNewName = GetNewName $termSetName

			$termSet = $termStoreGroup.TermSets[$termNewName]
			if($termSet -eq $null)
			{
				$termSet = $termStoreGroup.CreateTermSet($termNewName)
				#$termSet.SetDescription($termSetDesc, 1033)

				$termStore.CommitAll()

				write-host ("Created the TermSet - " + $termSetName)
			}
			else
			{
				write-host "TermSet Found with new name -"$termSetName
			}
		}
		elseif (([string[]]$termSetName).Length -eq 2)
		{
			write-host "Renaming Term from  - " + $termSetName[0] + " to " + $termSetName[1]

			$termSet.Name = $termSetName[1]
			$termStore.CommitAll()
		}
		else
		{
			write-host "TermSet Found -"$termSetName
		}

		$i = 0
		foreach($line in $termData)
		{   
			if($i -eq 0)
			{ 
				$i++
				continue 
			}

			$termDetail = $line.Split($delimiter)
			$isTermAvailable = $termDetail[$availableId].Replace("`"", "").Trim() #"

			$termNames = GetNamePair ($termDetail[$fstTrmId].Replace("`"", "").Trim()) #"

			$term = $termSet.Terms[([string[]]$termNames)[0]]
			if ($term -eq $null)
			{
			$termNewName = GetNewName $termNames

			$term = $termSet.Terms[$termNewName]

			if ($term -eq $null)
			{
				 Write-Host $termNames
				 Write-Host $termNewName
				 
				 if ($termDetail[$availableId] -eq "FALSE")
				 { 
					  Write-Host("Is False")
					  
					  $term = CreateTerm $termNewName $termDetail[$descId] $termSet $FALSE $termDetail[$lcid]
				 } 
				 else
				 { 
					$term = CreateTerm $termNewName $termDetail[$descId] $termSet $TRUE $termDetail[$lcid]
				 }
				}
				else
				{ 
					write-host "TermSet Found with new name -"$termName
				}
			}
			elseif (([string[]]$termNames).Length -eq 2)
			{
				write-host "Renaming Term from  - " + $termNames[0] + " to " + $termNames[1]

				$term.Name = $termNames[1]
				$termStore.CommitAll()
			}
			else
			{
				write-host "Term already present- "$termName -foregroundcolor green
				ManageTerm $fstTrmId $termDetail $term				
			}
			if([string]::IsNullOrEmpty($termDetail[($fstTrmId + 1)]))
			{
				Write-Host "True Found"
				Write-Host $termDetail

				DeprecateEnableTerms $term $termDetail

				CacheTermOrder $term $termSet
			}

			$i++
		} 
	}
	else
	{
		write-host "TermStore Group Not Found-"$termStoreGroupName -foregroundcolor red
	}
}

ManageTerms 

#Update the Term Store
$termStore.CommitAll()
Write-Host "Taxonomy created"

ApplyOrdering


ClearAll
