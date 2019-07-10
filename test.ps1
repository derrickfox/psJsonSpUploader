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
				if($i -eq "Ferrer-Alegre, Marc"){
					$i = "Ferrer, Marc"
				}
				if($i -eq "Michael, Samuel"){
					$i = "Michael, Sam"
				}
				if($i -eq "Lal, Madhu"){
					$errorList += "Lead Investigators -> 'Lal, Madhu' no longer works for NCATS.`n"
				}
				if($i -eq "Zhang, Li"){
					$errorList += "Lead Investigators -> There are 3 'Zhang, Li's in the system. Need a way to select the NCATS one.`n"
				}
				if($i -eq "Zhao, Jinghua"){
					$errorList += "Lead Investigators -> 'Zhao, Jinghua' is not in Active Directory.`n"
				}
				if($i -eq "Ching-Tze, Donald"){
					$errorList += "Lead Investigators -> 'Ching-Tze, Donald' is not in Active Directory.`n"
				}
				if($i -eq "Yang, Na"){
					$errorList += "Lead Investigators -> 'Yang, Na' exists in Active Directory; however, the code is not selecting their username.`n"
				}
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
					# Write-Output "ERROR Lead Investigator"
					# Write-Output $i

					# $errorList += "Error for 'Lead Investigator' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory.`n"
					# Write-Output "Error for 'Lead Investigator' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
				}
			}
			$newItem["Lead Investigators"] = $tempLeadInvestigators

			
			############# Supervisor of Record
			[Microsoft.SharePoint.SPFieldUserValueCollection]$tempSupervisofOfRecord = new-object Microsoft.SharePoint.SPFieldUserValueCollection
			foreach ($i in $supervisorOfRecord) {
				if($i -eq "undefined, n/a"){
					$tempSupervisofOfRecord = $Null
				}else{
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
							# Write-Output $tempSupervisorOfRecord
							$tempSupervisofOfRecord.Add($UserFieldValue)
						}
					}else{
						$numberOfErrorsFound++
						# Write-Output "ERROR Supervisor of Record"
						# Write-Output $i

						# $errorList += "Error for 'Supervisor of Record' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory.`n"
						# Write-Output "Error for 'Supervisor of Record' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
					}
				}
			}
			$newItem["Supervisor of Record"] = $tempSupervisofOfRecord


			############# NCATS Team Members
			[Microsoft.SharePoint.SPFieldUserValueCollection]$tempNcatsTeamMembers = new-object Microsoft.SharePoint.SPFieldUserValueCollection
			foreach ($i in $ncatsTeamMembers) {
				if($i -eq "Lal, Madhu"){
					$errorList += "NCATS Team Members -> 'Lal, Madhu' no longer works for NCATS.`n"
				}
				if($i -eq "Lu, Billy"){
					$errorList += "NCATS Team Members -> 'Lu, Billy' is no longer in Active Directory`n"
				}
				if($i -eq "Lee, Tobie"){
					$errorList += "NCATS Team Members -> 'Lee, Tobie' is no longer in Active Directory`n"
				}
				if($i -eq "Guha, Rajarshi"){
					$errorList += "NCATS Team Members -> 'Guha, Rajarshi' is no longer in Active Directory`n"
				}
				if($i -eq "Boxer, Matthew"){
					$errorList += "NCATS Team Members -> 'Boxer, Matthew' is no longer in Active Directory`n"
				}
				if($i -eq "Yang, Shu"){
					$errorList += "NCATS Team Members -> 'Yang, Shu' is no longer in Active Directory`n"
				}
				if($i -eq "Chen, Lu"){
					$errorList += "NCATS Team Members -> 'Chen, Lu' exists in Active Directory; however, the code is not selecting their username.`n"
				}
				if($i -eq "Li, Rong"){
					$errorList += "NCATS Team Members -> 'Li, Rong' exists in Active Directory; however, the code is not selecting their username.`n"
				}
				if($i -eq "Xu, Xin"){
					$errorList += "NCATS Team Members -> There are 3 'Xu, Xin's in the system. Need a way to select the NCATS' one.`n"
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
						# Write-Output "ERROR NCATS Team Members - catch block"
						# Write-Output $i

						# $errorList += "Error for 'NCATS Team Memeber' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory.`n"
						# Write-Output "Error for 'NCATS Team Memeber' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
					}
					
					$UserFieldValue = new-object Microsoft.SharePoint.SPFieldUserValue($theWeb, $User.ID, $User.LoginName)
					$tempNcatsTeamMembers.Add($UserFieldValue)
				}else{					
					$numberOfErrorsFound++
					# Write-Output "ERROR NCATS Team Members - else block"
					# Write-Output $i

					# $errorList += "Error for 'NCATS Team Memeber' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory.`n"
					# Write-Output "Error for 'NCATS Team Memeber' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
				}
			}
			$newItem["NCATS Team Members"] = $tempNcatsTeamMembers


			############# Intramural Collabs
			[Microsoft.SharePoint.SPFieldUserValueCollection]$tempIntCollabs = new-object Microsoft.SharePoint.SPFieldUserValueCollection
			foreach ($i in $intCollabs) {
				if($i -eq "indicated, none"){
					$tempIntCollabs = $null
				}else{
					$sam = ""
					$sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($i))" | select samaccountname
					if($sam[1] -ne $Null){
						$loginName = $sam.split("=")[1];
					}
					if($loginName -ne $Null){
						$loginName = $loginName.split("}")[0];	
					}
					if($loginName -ne $Null){
						$User = $theWeb.EnsureUser($loginName)
						$UserFieldValue = new-object Microsoft.SharePoint.SPFieldUserValue($theWeb, $User.ID, $User.LoginName)
						$tempIntCollabs.Add($UserFieldValue)
					}else{
						$numberOfErrorsFound++
						# Write-Output "ERROR Internal Collaborators"
						# Write-Output $i

						# $errorList += "Error for 'Internal Collaborators' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
						# Write-Output "Error for 'Internal Collaborators' on ZIA ID: $ziaIdNumber. Cannot find '$i' in Active Directory."
					}	
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
			Write-Output "******** Start of $ziaIdNumber ****************"
			Write-Output "Number of errors found: $numberOfErrorsFound"
			Write-Output "$errorList `n"
		}
		
	}

	Function Update-Projects
	{
		UpdateField "" "ProjectsEmpty"
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
