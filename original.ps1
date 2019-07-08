Param([string]$Url)

# Check to ensure Microsoft.SharePoint.PowerShell is loaded
$snapin = Get-PSSnapin | Where-Object {$_.Name -eq 'Microsoft.SharePoint.Powershell'}
if ($snapin -eq $null) {

Write-Host "Test:"
"Loading SharePoint Powershell Snapin"
Add-PSSnapin "Microsoft.SharePoint.Powershell"

}

	Function UpdateField
	{

		Param([string]$webName, [string] $listName, [string] $fieldName)

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

		if($theList.Fields.ContainsField($fieldName) -eq $true)
		{
			$JsonFilePath = 'C:\scripts\New folder\Projects.json'
			$json1 = Get-Content -Raw -Path $JsonFilePath | ConvertFrom-Json
			$getLI = $json1.projects[0].LeadInvestigators
			$getLI = $getLI.split("(")
			$getLI = $getLI.split(" ")
			$clientName = $getLI[2] + ', ' + $getLI[0]

			$sam = ""
			$sam += Get-ADUser -LDAPFilter "(ObjectClass=User)(anr=$($clientName))" | select samaccountname
			$loginName = $sam.split("=")[1];
			$loginName = $loginName.split("}")[0];
			$targetUser = $theWeb.ensureuser($loginName)
			if($targetUser -ne $Null){
				$newItem = $theList.items.add()
				$newitem["LI"] = $targetUser
				$newitem.update();
				#write-host $targetUser
			}
			write-host $loginName

		}

		$theWeb.Close()
		$theWeb.Dispose()
	}

	Function Update-Projects
	{
		UpdateField "" "Projects2" "Title"
	}

	$theSite = New-Object Microsoft.SharePoint.SPSite($Url)

    Update-Projects


    $theSite.Close()
	$theSite.Dispose()
