Param([string]$Url)

# Check to ensure Microsoft.SharePoint.PowerShell is loaded
$snapin = Get-PSSnapin | Where-Object {$_.Name -eq 'Microsoft.SharePoint.Powershell'}
if ($snapin -eq $null) {

Write-Host "Test:"
"Loading SharePoint Powershell Snapin"
Add-PSSnapin "Microsoft.SharePoint.Powershell"

}

$JsonFilePath = 'C:\Users\aafoxdm2\Desktop\powerShellConverter\thing.json'
$json1 = Get-Content -Raw -Path $JsonFilePath | ConvertFrom-Json

	Function UpdateField
	{
		Param([string]$webName, [string] $listName)

		$theWeb = $theSite.OpenWeb("/" + $webName)
		if ($theWeb -eq $null)
		{
			Write-Host "Could not open Web : " + $webName -ForegroundColor Red
			return
		}

		[Microsoft.SharePoint.SPList]$theList = $theWeb.Lists.TryGetList($listName)
		if ($theList -eq $null)
		{
			Write-Host "Could not open List : " + $listName -ForegroundColor Red
			return
		}else{
			Write-Output "theList is $theList"
		}


		# $JsonFilePath = 'C:\Users\aafoxdm2\Desktop\powerShellConverter\thing.json'
		# $json1 = Get-Content -Raw -Path $JsonFilePath | ConvertFrom-Json
		# Write-Output "Start"
		# Write-Output $json1
		# Write-Output "End"

		# $getLI = $json1.projects[0].LeadInvestigators
		# $getLI = $getLI.split("(")
		# $getLI = $getLI.split(" ")
		# $clientName = $getLI[2] + ', ' + $getLI[0]

		# $sam = ""
		# $sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($clientName))" | select samaccountname
		# $loginName = $sam.split("=")[1];
		# $loginName = $loginName.split("}")[0];
		# $reportTitle = $theWeb.ensureuser($loginName)
		$reportTitle = $json1.projects[$num]["Title"]
		$nihProjectID = $json1.projects[$num]["NIH Project ID"]
		$ziaIdNumber = $json1.projects[$num]["ZIA ID Number"]
		$ncatsDivision = $json1.projects[$num]["NCATS Division"]
		$2018ProjectStatus = $json1.projects[$num]["2018 Project Status"]
		$2019ProjectStatus = $json1.projects[$num]["2019 Project Status"]
		$leadInvestagors = $json1.projects[$num]["Lead Investigators"]
		$supervisorOrRecord = $json1.projects[$num]["Supervisor of Record"]
		$ncatsTeamMembers = $json1.projects[$num]["NCATS Team Members"]
		$intCollabs = $json1.projects[$num]["Intramural Collaborators (Affiliation)"]
		$extCollabs = $json1.projects[$num]["Extramural Collaborators (Affiliation)"]
		$humanCells = $json1.projects[$num]["Does project use human cells, human subjects or human tissues?"]
		$keywords = $json1.projects[$num]["Keywords"]
		$distinguishingKeyword = $json1.projects[$num]["Distinguishing Keyword"]
		$goalsAndObjectives = $json1.projects[$num]["Goals and Objectives"]
		$summary = $json1.projects[$num]["Summary"]

		if($reportTitle -ne $Null){
			$newItem = $theList.items.add()
			$newitem["Title"] = $reportTitle
			$newItem["NIH Project ID"] = $nihProjectID
			$newItem["ZIA ID Number"] = $ziaIdNumber
			$newItem["NCATS Division"] = $ncatsDivision
			$newItem["2018 Project Status"] = $2018ProjectStatus
			$newItem["2019 Project Status"] = $2019ProjectStatus
			$newItem["Lead Investigators"] = $leadInvestagors
			$newItem["Supervisor of Record"] = $supervisorOrRecord
			$newItem["NCATS Team Members"] = $ncatsTeamMembers
			$newItem["Intramural Collaborators (Affiliation)"] = $intCollabs
			$newItem["Extramural Collaborators (Affiliation)"] = $extCollabs
			$newItem["Does project use human cells, human subjects or human tissues?"] = $humanCells
			$newItem["Keywords"] = $keywords
			$newItem["Distinguishing Keyword"] = $distinguishingKeyword
			$newItem["Goals and Objectives"] = $goalsAndObjectives
			$newItem["Summary"] = $summary
			# $newitem.Update();
		}
	}

	# $theWeb.Close()
	# $theWeb.Dispose()
	

	Function Update-Projects
	{
		UpdateField "" "Projects2"
	}

	$theSite = New-Object Microsoft.SharePoint.SPSite($Url)

	# Update-Projects
	$count = 0
	# foreach ($num in $json1) {
	# 	# Update-Projects
	# 	# Write-Output $num["Report Title"]
	# 	$count++
	# 	Write-Output $count
	# }
	Write-Output $json1[0]


    $theSite.Close()
	$theSite.Dispose()
