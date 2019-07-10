Param([string]$Url)

# Import-Module ActiveDirectory
Import-Module ('ActiveDirectory')


# Check to ensure Microsoft.SharePoint.PowerShell is loaded
$snapin = Get-PSSnapin | Where-Object {$_.Name -eq 'Microsoft.SharePoint.Powershell'}
if ($snapin -eq $null) {
	Write-Host "Test:"
	"Loading SharePoint Powershell Snapin"
	Add-PSSnapin "Microsoft.SharePoint.Powershell"
}

$JsonFilePath = 'C:\Users\aafoxdm2\Desktop\powerShellConverter\thing.json'
$json1 = Get-Content -Raw -Path $JsonFilePath | ConvertFrom-Json
$numberOfErrorsFound = 0
$errorList

	Function UpdateField
	{
		Param([string]$webName, [string] $listName)

		$theWeb = $theSite.OpenWeb("/" + $webName)
		if ($theWeb -eq $null)
		{
			$numberOfErrorsFound++
			Write-Host "Could not open Web : " + $webName -ForegroundColor Red
			return
		}

		[Microsoft.SharePoint.SPList]$theList = $theWeb.Lists.TryGetList($listName)
		if ($theList -eq $null)
		{
			$numberOfErrorsFound++
			Write-Host "Could not open List : " + $listName -ForegroundColor Red
			return
		}

		$JsonFilePath = 'C:\Users\aafoxdm2\Desktop\powerShellConverter\thing.json'
		$json1 = Get-Content -Raw -Path $JsonFilePath | ConvertFrom-Json

		# $count = 56

		$reportTitle = $json1[$count]."Report Title"
		$nihProjectID = $json1[$count]."NIH Project ID"
		$ziaIdNumber = $json1[$count]."ZIA ID Number"
		$ncatsDivision = $json1[$count]."NCATS Division"
		$2018ProjectStatus = $json1[$count]."2018 Project Status"
		$2019ProjectStatus = $json1[$count]."2019 Project Status"
		$leadInvestagors = $json1[$count]."Lead Investigators"
		$supervisorOfRecord = $json1[$count]."Supervisor of Record"
		$ncatsTeamMembers = $json1[$count]."NCATS Team Members"
		$intCollabs = $json1[$count]."Intramural Collaborators (Affiliation)"
		$extCollabs = $json1[$count]."Extramural Collaborators (Affiliation)"
		$humanCells = $json1[$count]."Does project use human cells, human subjects or human tissues?"
		$keywords = $json1[$count]."Keywords"
		$distinguishingKeyword = $json1[$count]."Distinguishing Keyword"
		$goalsAndObjectives = $json1[$count]."Goals and Objectives"
		$summary = $json1[$count]."Summary"

		if($reportTitle -ne $Null){
			$newItem = $theList.items.add()
			$newitem["ReportTitle"] = $reportTitle
			$newItem["NIH Project ID"] = $nihProjectID
			$newItem["ZIA ID Number"] = $ziaIdNumber
			$newItem["NCATS Division"] = $ncatsDivision
			$newItem["2018 Project Status"] = $2018ProjectStatus
			$newItem["2019 Project Status"] = $2019ProjectStatus

			############# Lead Investigator
			[Microsoft.SharePoint.SPFieldUserValueCollection]$tempLeadInvestigators = new-object Microsoft.SharePoint.SPFieldUserValueCollection
			foreach ($i in $leadInvestagors) {
				$sam = ""
				$sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($i))" | select samaccountname
				if($sam[1] -ne $Null){
					$loginName = $sam.split("=")[1];
				}
				if($loginName -ne $Null){
					$loginName = $loginName.split("}")[0];	
				}

				if($loginName){
					$User = $theWeb.EnsureUser($loginName)
					$UserFieldValue = new-object Microsoft.SharePoint.SPFieldUserValue($theWeb, $User.ID, $User.LoginName)
					# $tempLeadInvestigators += $UserFieldValue
					$tempLeadInvestigators.Add($UserFieldValue)
				}else{
					$numberOfErrorsFound++
					$errorList += "Error for 'Lead Investigator' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory.`n"
					# Write-Output "Error for 'Lead Investigator' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
				}
			}
			$newItem["Lead Investigators"] = $tempLeadInvestigators

			
			############# Supervisor of Record
			[Microsoft.SharePoint.SPFieldUserValueCollection]$tempSupervisofOfRecord = new-object Microsoft.SharePoint.SPFieldUserValueCollection
			foreach ($i in $supervisorOfRecord) {
				# Write-Output $i
				$sam = ""
				$sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($i))" | select samaccountname
				if($sam[1] -ne $Null){
					$loginName = $sam.split("=")[1];
				}
				if($loginName -ne $Null){
					$loginName = $loginName.split("}")[0];	
				}
				if($Null -ne $loginName){
					$User = $theWeb.EnsureUser($loginName)
					$UserFieldValue = new-object Microsoft.SharePoint.SPFieldUserValue($theWeb, $User.ID, $User.LoginName)
					if($UserFieldValue){
						Write-Output $tempSupervisorOfRecord
						$tempSupervisofOfRecord.Add($UserFieldValue)
					}
				}else{
					$numberOfErrorsFound++
					$errorList += "Error for 'Supervisor of Record' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory.`n"
					# Write-Output "Error for 'Supervisor of Record' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
				}
			}
			$newItem["Supervisor of Record"] = $tempSupervisofOfRecord


			############# NCATS Team Members
			[Microsoft.SharePoint.SPFieldUserValueCollection]$tempNcatsTeamMembers = new-object Microsoft.SharePoint.SPFieldUserValueCollection
			foreach ($i in $ncatsTeamMembers) {
				# Write-Output $i
				if($i -eq "Li, Rong"){
					$errorList += "'Li, Rong' is no longer in Active Directory`n"
					# Write-Output "'Li, Rong' is no longer in Active Directory"
				}
				if($i -eq "Lu, Billy"){
					$errorList += "'Lu, Billy' is no longer in Active Directory`n"
					# Write-Output "'Lu, Billy' is no longer in Active Directory"
				}
				if($i -eq "Yang, Shu"){
					$errorList += "'Yang, Shu' is no longer in Active Directory`n"
					# Write-Output "There are 2 'Yang, Shu's in Active Directory. Need a way to select NCATS' one."
				}
				$sam = ""
				$sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($i))" | select samaccountname
				if($sam[1] -ne $Null){
					$loginName = $sam.split("=")[1];
				}
				if($Null -ne $loginName){
					$loginName = $loginName.split("}")[0];	
				}
				if($loginName -ne $Null){
					try{
						$User = $theWeb.EnsureUser($loginName)
					} 
					catch{
						$numberOfErrorsFound++
						$errorList += "Error for 'NCATS Team Memeber' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory.`n"
						# Write-Output "Error for 'NCATS Team Memeber' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
					}
					
					$UserFieldValue = new-object Microsoft.SharePoint.SPFieldUserValue($theWeb, $User.ID, $User.LoginName)
					$tempNcatsTeamMembers.Add($UserFieldValue)
				}else{
					$numberOfErrorsFound++
					$errorList += "Error for 'NCATS Team Memeber' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory.`n"
					# Write-Output "Error for 'NCATS Team Memeber' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
				}
			}
			$newItem["NCATS Team Members"] = $tempNcatsTeamMembers


			############# Intramural Collabs
			[Microsoft.SharePoint.SPFieldUserValueCollection]$tempIntCollabs = new-object Microsoft.SharePoint.SPFieldUserValueCollection
			foreach ($i in $intCollabs) {
				$sam = ""
				$sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($i))" | select samaccountname
				if($sam[1] -ne $Null){
					$loginName = $sam.split("=")[1];
				}
				if($sam[0] -ne $Null){
					$loginName = $loginName.split("}")[0];	
				}
				if($loginName -ne $Null){
					$User = $theWeb.EnsureUser($loginName)
					$UserFieldValue = new-object Microsoft.SharePoint.SPFieldUserValue($theWeb, $User.ID, $User.LoginName)
					$tempIntCollabs.Add($UserFieldValue)
				}else{
					$numberOfErrorsFound++
					$errorList += "Error for 'Internal Collaborators' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
					# Write-Output "Error for 'Internal Collaborators' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
				}	
			}
			$newItem["IntramuralCollaborators"] = $tempIntCollabs
			
			$newItem["Extramural Collaborators (Affiliation)"] = $extCollabs
			$newItem["Keywords"] = $keywords
			$newItem["DoesProjectUseHumanBiospecimen"] = $humanCells
			$newItem["Distinguishing Keyword"] = $distinguishingKeyword
			$newItem["Goals and Objectives"] = $goalsAndObjectives
			$newItem["Summary"] = $summary
			$newItem["Title"] = "Test Title"
			$newitem.Update();
		}
		$theWeb.Close()
		$theWeb.Dispose()
		if($numberOfErrorsFound -gt 0){
			Write-Output "******** Start of $ziaIdNumber **************** `n"
			Write-Output "Number of errors found: $numberOfErrorsFound"
			Write-Output "$errorList" "********* End of $ziaIdNumber ****************************** `n"
		}
		
	}

	Function Update-Projects
	{
		UpdateField "" "Projects2"
	}

	$theSite = New-Object Microsoft.SharePoint.SPSite($Url)

	# Update-Projects

	$count = 0
	foreach ($num in $json1) {
		Update-Projects
		$count++
	}

    $theSite.Close()
	$theSite.Dispose()
