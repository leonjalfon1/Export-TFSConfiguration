param
(
    [Parameter(Mandatory=$true)]
    $TFSCollection,
    [Parameter(Mandatory=$true)]
    $Path,
    [Parameter(Mandatory=$true)]
    $User,
    [Parameter(Mandatory=$true)]
    $Password,
    [Parameter(Mandatory=$false)]
    $WitadminLocation = "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
)


function Get-TeamProjects
{
    param
    (
        [Parameter(Mandatory=$true)]
        $CollectionUrl,
        [Parameter(Mandatory=$true)]
        $User,
        [Parameter(Mandatory=$true)]
        $Password
    )

	try
	{
        $apiVersion = "1.0"
        $requestUrl = "$CollectionUrl/_apis/projects" + "?api-version=$apiVersion"
		$credentials = New-Object System.Management.Automation.PSCredential($User, $($Password | ConvertTo-SecureString -AsPlainText -Force))
		$response = Invoke-RestMethod -Method GET -Credential $credentials -ContentType application/json -Uri $requestUrl
		$TeamProjects = $response.value | Select name
		return $TeamProjects
	}

	catch
	{
        Write-Host "Failed to get Team Projects for collection {$CollectionUrl}, Exception: $_" -ForegroundColor Red
		return $null
	}
}


# Check Parameters

if(!(Test-Path $Path))
{
    Write-Host "Error, Invalid Path {$Path}"
    return
}

if(!(Test-Path $WitadminLocation\witadmin.exe))
{
    Write-Host "Invalid WitadminLocation {$WitadminLocation}, trying to find {witadmin.exe}..." -ForegroundColor Yellow

    $WitadminLocation = "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
    if(!(Test-Path $WitadminLocation\witadmin.exe))
    {
        $WitadminLocation = "C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE"
        if(!(Test-Path $WitadminLocation\witadmin.exe))
        {
            $WitadminLocation = "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE"
            if(!(Test-Path $WitadminLocation\witadmin.exe))
            {
                Write-Host "Error, Witadmin.exe not found"
                return
            }
        }
    }
}


# Initialization

Set-Alias -Name witadmin -Value "$WitadminLocation\witadmin.exe"
cd $WitadminLocation
$TargetFolder = Get-Date -format "yyyyMMddTHHmmssffff"

Write-Host "Creating target folder {$TargetFolder}..." -ForegroundColor Cyan
$cmdResponse = mkdir -Path ($Path) -Name $TargetFolder
$Path = $Path + "\" + $TargetFolder


# Retrieve TP's
$TeamProjects = Get-TeamProjects -CollectionUrl $TFSCollection -User $User -Password $Password


# Export Collection XML's

Write-Host "Exporting {Globallist}..." -ForegroundColor Yellow
witadmin exportgloballist /collection:"$TFSCollection" /f:"$Path\_GlobalList.xml"

Write-Host "Exporting {GlobalWorkflow}..." -ForegroundColor Yellow
witadmin exportglobalworkflow  /collection:$TFSCollection /f:"$Path\_GlobalWorkflow.xml"


# Export Team Projects XML's

foreach($Project in $TeamProjects)
{
    $TP = $Project.name

    Write-Host "Creating new folder for {$TP} XML's..." -ForegroundColor Cyan
    $cmdResponse = mkdir -Path ($Path) -Name $TP

    Write-Host "Retrieving existing Work Items in {$TP}..."
    $WorkItemsTypes = witadmin listwitd /collection:$TFSCollection /p:"$TP"


    # Export WorkItems Configurations

    foreach($WIT in $WorkItemsTypes.Split("`r`n"))
    {
        if($WIT)
        {
            Write-Host "Exporting {$WIT} in {$TP}..." -ForegroundColor Yellow
            witadmin exportwitd /collection:$TFSCollection /p:"$TP" /n:"$WIT" /f:"$Path\$TP\$WIT.xml"
        } 
    }


    # Export Project Configurations

    Write-Host "Exporting {Categories} in {$TP}" -ForegroundColor Yellow
    witadmin exportcategories /collection:$TFSCollection /p:"$TP" /f:"$Path\$TP\_Categories.xml"
    
    Write-Host "Exporting {ProcessConfig} in {$TP}" -ForegroundColor Yellow
    witadmin exportprocessconfig /collection:$TFSCollection /p:"$TP" /f:"$Path\$TP\_ProcessConfig.xml"

    Write-Host "Exporting {GlobalWorkflow} in {$TP}" -ForegroundColor Yellow
    witadmin exportglobalworkflow  /collection:$TFSCollection /p:"$TP" /f:"$Path\$TP\_GlobalWorkflow.xml"
}

Write-Host ""
Write-Host " ----------------------------------- " -ForegroundColor Cyan
Write-Host "| TFS Export Configuration Finished |" -ForegroundColor Cyan
Write-Host " ----------------------------------- " -ForegroundColor Cyan