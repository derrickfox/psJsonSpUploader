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
		}

		$JsonFilePath = 'C:\Users\aafoxdm2\Desktop\powerShellConverter\thing.json'
		$json1 = Get-Content -Raw -Path $JsonFilePath | ConvertFrom-Json
		# $getLI = $json1.projects[0].LeadInvestigators
		# $getLI = $getLI.split("(")
		# $getLI = $getLI.split(" ")
		# $clientName = $getLI[2] + ', ' + $getLI[0]

		# $sam = ""
		# $sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($clientName))" | select samaccountname
		# $loginName = $sam.split("=")[1];
		# $loginName = $loginName.split("}")[0];
		# $reportTitle = $theWeb.ensureuser($loginName)

		# $rawName = $json1[12]."Supervisor of Record"
		# Write-Output $rawName
		# $testName2 = Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($rawName))" | select samaccountname
		# Write-Output $testName2

		$count = 12

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
			$tempLeadInvestigators
			foreach ($i in $leadInvestagors) {
				# # $testName = 'wzheng'
				# Write-Output $i
				# $testName = Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($i))" | select samaccountname
				# Write-Output "testName $testName"
				# Write-Output $testName
				$sam = ""
				$sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($i))" | select samaccountname
				$loginName = $sam.split("=")[1];
				$loginName = $loginName.split("}")[0];
				Write-Output $loginName
				# $targetUser = $theWeb.ensureuser($loginName)

				if($loginName -ne $Null){
					$User = $theWeb.EnsureUser($loginName)
					$UserFieldValue = new-object Microsoft.SharePoint.SPFieldUserValue($theWeb, $User.ID, $User.LoginName)
					Write-Output $UserFieldValue
					$tempLeadInvestigators += $UserFieldValue
				}
			}
			$newItem["Lead Investigators"] = $tempLeadInvestigators

			
			############# Supervisor of Record
			$tempSupervisofOfRecord
			foreach ($i in $supervisorOfRecord) {
				$sam = ""
				$sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($i))" | select samaccountname
				$loginName = $sam.split("=")[1];
				$loginName = $loginName.split("}")[0];
				if($loginName -ne $Null){
					$User = $theWeb.EnsureUser($loginName)
					$UserFieldValue = new-object Microsoft.SharePoint.SPFieldUserValue($theWeb, $User.ID, $User.LoginName)
					$tempSupervisorOfRecord += $UserFieldValue
				}
			}
			$newItem["Supervisor of Record"] = $tempSupervisofOfRecord


			############# NCATS Team Members
			$tempNcatsTeamMembers
			foreach ($i in $ncatsTeamMembers) {
				$sam = ""
				$sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($i))" | select samaccountname
				$loginName = $sam.split("=")[1];
				$loginName = $loginName.split("}")[0];
				if($loginName -ne $Null){
					$User = $theWeb.EnsureUser($loginName)
					$UserFieldValue = new-object Microsoft.SharePoint.SPFieldUserValue($theWeb, $User.ID, $User.LoginName)
					$tempNcatsTeamMembers += $UserFieldValue
				}
			}
			$newItem["NCATS Team Members"] = $tempNcatsTeamMembers

			
			############# Intramural Collabs
			$tempIntCollabs
			foreach ($i in $intCollabs) {
				$sam = ""
				$sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($i))" | select samaccountname
				$loginName = $sam.split("=")[1];
				$loginName = $loginName.split("}")[0];
				if($loginName -ne $Null){
					$User = $theWeb.EnsureUser($loginName)
					$UserFieldValue = new-object Microsoft.SharePoint.SPFieldUserValue($theWeb, $User.ID, $User.LoginName)
					$tempIntCollabs += $UserFieldValue
				}	
			}
			$newItem["Internal Collaborators (affiliation)"] = $tempIntCollabs

			
			$newItem["Extramural Collaborators (Affiliation)"] = $extCollabs
			$newItem["Keywords"] = $keywords
			$newItem["DoesProjectUseHumanBiospecimen"] = $humanCells
			$newItem["Distinguishing Keyword"] = $distinguishingKeyword
			$newItem["Goals and Objectives"] = $goalsAndObjectives
			$newItem["Summary"] = $summary
			$newitem.Update();
		}
		$theWeb.Close()
		$theWeb.Dispose()
	}

	Function Update-Projects
	{
		UpdateField "" "Projects2"
	}

	$theSite = New-Object Microsoft.SharePoint.SPSite($Url)

	Update-Projects

	# $count = 0
	# foreach ($num in $json1) {
	# 	# Update-Projects
	# 	$count++
	# }

    $theSite.Close()
	$theSite.Dispose()
