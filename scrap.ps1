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
		Write-Output "webName is $webName"
		Write-Output "listName is $listName"
		Write-Output "fieldName is $fieldName"

		[Microsoft.SharePoint.SPList]$theList = $theWeb.Lists.TryGetList($listName)
		if ($theList -eq $null)
		{
			Write-Host "Could not open List : " + $listName -ForegroundColor Red
			return
		}else{
			Write-Output "theList is $theList"
		}

		if($theList.Fields.ContainsField($fieldName) -eq $true)
		{
			$targetUser = "I am a test value"
			if($targetUser -ne $Null){
				$newItem = $theList.items.add()
				$newitem["Title"] = $targetUser
				$newitem.Update();
				write-host $targetUser
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
