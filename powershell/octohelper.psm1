if(-not $env:octopusAPIKey -or -not $env:OctopusURL) {
    Error "Make sure env:octopusAPIKey and env:OctopusURL are set before using the octohelper module"
    Error "Refer to https://octoposh.readthedocs.io/en/latest/gettingstarted/setting-credentials/ for more info"
}

#Intial Octopus setup where all necessary Octoposh tools are installed
function IntitialOctoSetup  {
    Log "Intitializing Octoposh setup if necessary..."
    #Check if octoposh execution file is available
	$octoExePath = 'C:\Tools\OctopusTools.4.39.2\Octo.exe'
	if(-not (Test-Path -Path $octoExePath)) {
		Log 'Installing Octoposh module . . .'
		Install-Module -Name Octoposh -force
		Set-OctopusToolsFolder -path "C:\Tools"
		Install-OctopusTool -version 4.39.2
	}
	else {
		Log "Octoposh module already exists."
	}
	
	#Install Octoposh as a helpter tool for accessing our Octopus instance
	if (-not (Get-Module -ListAvailable -Name Octoposh)) {
		Log 'Importing Octoposh module . . .'
		Import-Module Octoposh
	}
    
    #Calling Set-OctopusToolsFolder twice due to environment variables needing to get set for multiple cases
    Set-OctopusToolsFolder -path "C:\Tools"
	Set-OctopusToolPath -version 4.39.2
}

<#
    Project functions
#>

function CreateProject {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("p")] [string]$ProjectName,
        [Parameter(Position=2, Mandatory=$false)] [alias("g")] [string]$ProjectGroup,
        [Parameter(Position=3, Mandatory=$false)] [alias("l")] [string]$ProjectLifeCycle
    )
    if(-not $ProjectName) {
        $ProjectName = Read-Host "Enter Project Name"
        if([String]::IsNullOrWhiteSpace($ProjectName)) {
            Error "Invalid Project name, please try again and enter a valid name."
            CreateProject -g $ProjectGroup -l $ProjectLifeCycle
        }
    }
    if(-not $ProjectGroup) {
        $ProjectGroup = UserSelectProjectGroup
    }
    if(-not $ProjectLifeCycle) {
        $ProjectLifeCycle = UserSelectLifeCycle
    }
    $Project = Get-OctopusResourceModel -Resource Project

    $ProjectGroup = GetProjectGroups -ProjectGroupName $ProjectGroup
    $LifeCycle = GetLifeCycles -LifeCycleName $ProjectLifeCycle

    $Project.name = $ProjectName
    $Project.ProjectGroupID = $ProjectGroup.id
    $Project.LifecycleId = $LifeCycle.id

    New-OctopusResource -Resource $Project
}

function DeleteProject {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("p")] [string[]]$ProjectNames
    )
    if(confirmation) {
        GetProjects -ProjectName $ProjectNames | Remove-OctopusResource
    }
}

function GetProjects {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string[]]$ProjectNames,
        [Parameter(Position=2, Mandatory=$false)] [alias("d")] [string[]]$ProjectId,
        [Parameter(Position=3, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=4, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    if($ProjectNames) {
        return Get-OctopusProject -ProjectName $ProjectNames -ResourceOnly
    }
    elseif ($ProjectId) {
        return Get-OctopusProject -ResourceOnly | Where-Object {$_.Id -eq $ProjectId}
    }
    else {
        if($Inclusions) {
            return GetProjects | Where-Object {$_.Id -In $Inclusions}
        }
        elseif($Exclusions) {
            return GetProjects | Where-Object {$_.Id -NotIn $Exclusions}
        }
        else {
            return Get-OctopusProject -ResourceOnly
        }
    }
}

function UserSelectProject {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $projects = GetProjects -Inclusions $Inclusions -Exclusions $Exclusions
	return UserSelectPrompt $projects "Project"
}

function UserSelectProjects {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $projects = GetProjects -Inclusions $Inclusions -Exclusions $Exclusions
	return UserMultiSelectPrompt $projects "Project"
}

function GetProjectObjects {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("e")] [string[]]$ProjectNames,
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$ProjectIds
    )

    [System.Collections.ArrayList]$ProjectObjects = @()
    
    if($ProjectNames) {
        foreach($env in $ProjectNames) {
            $Project = Get-OctopusProject -ProjectName $env -ResourceOnly
            $ProjectObjects.Add($Project)
        }
    }
    elseif ($ProjectIds) {
        foreach($proj in $ProjectIds) {
            $Project = Get-OctopusProject -ResourceOnly | Where-Object {$_.Id -eq $proj}
            $ProjectObjects.Add($Project)
        }
    }

    return $ProjectObjects
}

<#
    Project Group functions
#>

function CreateProjectGroup {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string]$ProjectGroupName
    )
    if(-not $ProjectGroupName) {
        $ProjectGroupName = Read-Host "Enter ProjectGroup Name"
        if([String]::IsNullOrWhiteSpace($ProjectGroupName)) {
            Error "Invalid ProjectGroup name, please try again and enter a valid name."
            CreateProjectGroup
        }
    }
    $ProjectGroup = Get-OctopusResourceModel -Resource ProjectGroup

    $ProjectGroup.Name = $ProjectGroupName

    New-OctopusResource -Resource $ProjectGroup
}

function GetProjectGroups {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string[]]$ProjectGroupName,
        [Parameter(Position=2, Mandatory=$false)] [alias("d")] [string[]]$ProjectGroupId,
        [Parameter(Position=3, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=4, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    if($ProjectGroupName) {
        return Get-OctopusProjectGroup -name $ProjectGroupName -ResourceOnly
    }
    elseif ($ProjectGroupId) {
        return Get-OctopusProjectGroup -ResourceOnly | Where-Object {$_.Id -eq $ProjectGroupId}
    }
    else {
        if($Inclusions) {
            return GetProjectGroups | Where-Object {$_.Id -In $Inclusions}
        }
        elseif($Exclusions) {
            return GetProjectGroups | Where-Object {$_.Id -NotIn $Exclusions}
        }
        else {
            return Get-OctopusProjectGroup -ResourceOnly
        }
    }
}

function UserSelectProjectGroup {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $projectGroups = GetProjectGroups -Inclusions $Inclusions -Exclusions $Exclusions
	return UserSelectPrompt $projectGroups "Project Group"
}

function UserSelectProjectGroups {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $projectGroups = GetProjectGroups -Inclusions $Inclusions -Exclusions $Exclusions
	return UserMultiSelectPrompt $projectGroups "Project Group"
}

<#
    Lifecycle functions
#>

function GetLifeCycles {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string[]]$LifeCycleName,
        [Parameter(Position=2, Mandatory=$false)] [alias("d")] [string[]]$LifeCycleId,
        [Parameter(Position=3, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=4, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    if($LifeCycleName) {
        return Get-OctopusLifecycle -name $LifeCycleName -ResourceOnly
    }
    elseif ($LifeCycleId) {
        return Get-OctopusLifecycle -ResourceOnly | Where-Object {$_.Id -eq $LifeCycleId}
    }
    else {
        if($Inclusions) {
            return GetLifeCycles | Where-Object {$_.Id -In $Inclusions}
        }
        elseif($Exclusions) {
            return GetLifeCycles | Where-Object {$_.Id -NotIn $Exclusions}
        }
        else {
            return Get-OctopusLifecycle -ResourceOnly
        }
    }
}

function UserSelectLifeCycle {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $lifeCycles = GetLifeCycles -Inclusions $Inclusions -Exclusions $Exclusions
	return UserSelectPrompt $lifeCycles "LifeCycle"
}

function UserSelectLifeCycles {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $lifeCycles = GetLifeCycles -Inclusions $Inclusions -Exclusions $Exclusions
	return UserMultiSelectPrompt $lifeCycles "LifeCycle"
}

<#
    Environment functions
#>

function CreateEnvironment {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string]$EnvironmentName
    )
    if(-not $EnvironmentName) {
        $EnvironmentName = Read-Host "Enter Environment Name"
        if([String]::IsNullOrWhiteSpace($EnvironmentName)) {
            Error "Invalid Environment name, please try again and enter a valid name."
            CreateEnvironment
        }
    }
    $Environment = Get-OctopusResourceModel -Resource Environment

    $Environment.name = $EnvironmentName

    New-OctopusResource -Resource $Environment
}

function DeleteEnvironment {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("e")] [string]$EnvironmentName
    )
    if(confirmation) {
        GetEnvironments -EnvironmentName $EnvironmentName | Remove-OctopusResource
    }
}

function GetEnvironments {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string]$EnvironmentName,
        [Parameter(Position=2, Mandatory=$false)] [alias("d")] [string[]]$EnvironmentId,
        [Parameter(Position=3, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=4, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    if($EnvironmentName) {
        return Get-OctopusEnvironment -name $EnvironmentName -ResourceOnly
    }
    elseif ($EnvironmentId) {
        return Get-OctopusEnvironment -ResourceOnly | Where-Object {$_.Id -eq $EnvironmentId}
    }
    else {
        if($Inclusions) {
            return GetEnvironments | Where-Object {$_.Id -In $Inclusions}
        }
        elseif($Exclusions) {
            return GetEnvironments | Where-Object {$_.Id -NotIn $Exclusions}
        }
        else {
            return Get-OctopusEnvironment -ResourceOnly
        }
    }
}

function UserSelectEnvironment {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $environments = GetEnvironments -Inclusions $Inclusions -Exclusions $Exclusions
	return UserSelectPrompt $environments "Environment"
}

function UserSelectEnvironments {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $environments = GetEnvironments -Inclusions $Inclusions -Exclusions $Exclusions
	return UserMultiSelectPrompt $environments "Environment"
}

function GetEnvironmentObjects {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("e")] [string[]]$EnvironmentNames,
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$EnvironmentIds
    )

    if($EnvironmentNames) {
        return Get-OctopusEnvironment -EnvironmentName $EnvironmentNames -ResourceOnly
    }
    elseif ($EnvironmentIds) {
        return Get-OctopusEnvironment -ResourceOnly | Where-Object {$_.Id -in $EnvironmentIds}
    }
    else {
        return Get-OctopusEnvironment -ResourceOnly
    }
}

<#
    Machine Functions
#>

function CreateMachine {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string]$MachineName,
        [Parameter(Position=2, Mandatory=$false)] [alias("h")] [string]$MachineHostname,
        [Parameter(Position=3, Mandatory=$false)] [alias("r")] [string]$MachineRoles,
        [Parameter(Position=4, Mandatory=$false)] [alias("e")] [string]$MachineEnvironments
    )
    if(-not $MachineName) {
        $MachineName = Read-Host "Enter Machine Name"
        if([String]::IsNullOrWhiteSpace($MachineName)) {
            Error "Invalid Machine name, please try again and enter a valid name."
            CreateMachine -h $MachineHostname -r $MachineRoles -e $MachineEnvironments
        }
    }
    if(-not $MachineHostname) {
        $MachineHostname = Read-Host "Enter Machine Host Name (ComputerName or the IP address of the Tentacle machine)"
        if([String]::IsNullOrWhiteSpace($MachineHostname)) {
            Error "Invalid Machine name, please try again and enter a valid name."
            CreateMachine -n $MachineName -r $MachineRoles -e $MachineEnvironments
        }
    }
    if(-not $MachineRoles) {
        $MachineRoles = UserSelectMachineRoles
    }
    if(-not $MachineEnvironments) {
        $MachineEnvironments = UserSelectEnvironments
    }

    $machine = Get-OctopusResourceModel -Resource Machine

    $environments = Get-OctopusEnvironment -EnvironmentName $MachineEnvironments -ResourceOnly
    $machine.name = $MachineName

    foreach($environment in $environments){
        $machine.EnvironmentIds.Add($environment.id)
    }
    foreach ($role in $MachineRoles){
        $machine.Roles.Add($role)    
    }
    #Use the Discover API to get the machine thumbprint.
    $discover = (Invoke-WebRequest "$env:OctopusURL/api/machines/discover?host=$machineHostname&type=TentaclePassive" -Headers (New-OctopusConnection).header).content | ConvertFrom-Json

    $machineEndpoint = New-Object Octopus.Client.Model.Endpoints.ListeningTentacleEndpointResource
    $machine.EndPoint = $machineEndpoint
    $machine.Endpoint.Uri = $discover.Endpoint.Uri
    $machine.Endpoint.Thumbprint = $discover.Endpoint.Thumbprint

    New-OctopusResource -Resource $machine
}

function DeleteMachine {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("e")] [string]$MachineName
    )
    if(confirmation) {
        GetMachines -MachineName $MachineName | Remove-OctopusResource
    }
}

function GetMachines {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string]$MachineName,
        [Parameter(Position=2, Mandatory=$false)] [alias("d")] [string[]]$MachineId,
        [Parameter(Position=3, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=4, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    if($MachineName) {
        return Get-OctopusMachine -name $MachineName -ResourceOnly
    }
    elseif ($MachineId) {
        return Get-OctopusMachine -ResourceOnly | Where-Object {$_.Id -eq $MachineId}
    }
    else {
        if($Inclusions) {
            return GetMachines | Where-Object {$_.Id -In $Inclusions}
        }
        elseif($Exclusions) {
            return GetMachines | Where-Object {$_.Id -NotIn $Exclusions}
        }
        else {
            return Get-OctopusMachine -ResourceOnly
        }
    }
}

function UserSelectMachine {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $machines = GetMachines -Inclusions $Inclusions -Exclusions $Exclusions
	return UserSelectPrompt $machines "Machine"
}

function UserSelectMachines {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $machines = GetMachines -Inclusions $Inclusions -Exclusions $Exclusions
	return UserMultiSelectPrompt $machines "Machine"
}

function GetMachineRoles {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    [System.Collections.ArrayList]$machineRoles = @()

    $machines = GetMachines
    foreach($machine in $machines) {
        $machineRoles.Add($machine.Roles)
    }
    $machineRoles = $machineRoles | Select-Object -uniq
    
    if($Inclusions) {
        return GetMachineRoles | Where-Object {$_ -In $Inclusions}
    }
    elseif($Exclusions) {
        return GetMachineRoles | Where-Object {$_ -NotIn $Exclusions}
    }
    else {
        return $machineRoles
    }
}

function UserSelectMachineRole {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $machineRoles = GetMachineRoles -Inclusions $Inclusions -Exclusions $Exclusions
	return UserSelectPrompt $machineRoles "Machine Role"
}

function UserSelectMachineRoles {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $machineRoles = GetMachineRoles -Inclusions $Inclusions -Exclusions $Exclusions
	return UserMultiSelectPrompt $machineRoles "Machine Role"
}

<#
    Tenant functions
#>

function CreateTenant  {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string]$TenantName,
        [Parameter(Position=2, Mandatory=$false)] [alias("p")] [string[]]$ProjectNames,
        [Parameter(Position=3, Mandatory=$false)] [alias("e")] [string[]]$EnvironmentNames,
        [Parameter(Position=4, Mandatory=$false)] [alias("t")] [string]$Tag
    )
    if(-not $TenantName) {
        $TenantName = Read-Host "Enter Tenant Name"
        if([String]::IsNullOrWhiteSpace($TenantName)) {
            Error "Invalid Tenant name, please try again and enter a valid name."
            CreateTenant -p $ProjectNames -t $Tag
        }
    }

    if(-not $Tag) {
        $Tag = "Internal"
    }

    $env:TenantName = ParseTenantName $TenantName

	$ExistingTenant = Get-OctopusTenant -name $env:TenantName

    [System.Collections.ArrayList]$EnvironmentObjects = @()
	[System.Collections.ArrayList]$ProjectObjects = @()

	if(-not $ExistingTenant){
        Log "Creating new tenant $env:TenantName."
        if(-not $ProjectNames) {
            $env:ProjectNames = UserSelectProjects
        }
        else {
            $env:ProjectNames = $ProjectNames
        }
        if (-not $EnvironmentNames) {
            $env:EnvironmentNames = UserSelectEnvironments
        }
        else {
            $env:EnvironmentNames = $EnvironmentNames
        }

		$Tenant = Get-OctopusResourceModel -Resource Tenant
		Log "Adding projects to tenant $env:TenantName."
		foreach($proj in $env:ProjectNames) {
            Log "Adding $proj"
			$Project = Get-OctopusProject -ProjectName $proj -ResourceOnly
			$ProjectObjects.Add($Project) | Out-Null
		}
		
		Log "Adding environments to tenant $env:TenantName."
		foreach($env in $env:EnvironmentNames) {
            Log "Adding $env"
			$Environment = Get-OctopusEnvironment -EnvironmentName $env -ResourceOnly
			$EnvironmentObjects.Add($Environment) | Out-Null
        }
        
        $tagSet = Get-OctopusTagSet -ResourceOnly
        $tagObject = $tagSet.Tags | Where-Object {$_.Name -eq $Tag}

        $Tenant.Name = $env:TenantName | Out-Null
		foreach($proj in $ProjectObjects) {
			$Tenant.ConnectToProjectAndEnvironments($proj,$EnvironmentObjects)
		}
		
		$Tenant.WithTag($tagObject) | Out-Null
        
		New-OctopusResource $Tenant
        Log "Created tenant $env:TenantName successfully!"
	}
	else {
		Warning "A tenant with the same name '$env:TenantName' already exists.  No new tenant was created."
	}
}

function ParseTenantName {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string]$TenantName
    )
	$pattern = '[^a-zA-Z1-9_]'
    $TenantName = $TenantName -replace $pattern, ''
    return $TenantName
}

function AddTenantToProject {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("t")] [string]$TenantName,
        [Parameter(Position=2, Mandatory=$false)] [alias("p")] [string]$ProjectName
    )
    if(-not $TenantName) {
        $TenantName = UserSelectTenant
    }
    $tenant = GetTenants -n $TenantName

    if(-not $ProjectName) {
        $exclusions = $tenant.ProjectEnvironments.Keys
        $ProjectName = UserSelectProject -Exclusions $exclusions
    }

    $projectEnvironments = $tenant.ProjectEnvironments
    $projectId = $projectEnvironments.Keys | Select-Object -First 1
    $environmentIds = $tenant.ProjectEnvironments[$projectId]

    $project = GetProjects -n $ProjectName
    $environmentObjects = GetEnvironmentObjects -EnvironmentIds $environmentIds

    $tenant.ConnectToProjectAndEnvironments($project,$environmentObjects)
    Update-OctopusResource -Resource $tenant
}

function RemoveTenantFromProject {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("t")] [string]$TenantName,
        [Parameter(Position=2, Mandatory=$false)] [alias("p")] [string]$ProjectName
    )
    if(-not $TenantName) {
        $TenantName = UserSelectTenant
    }
    $tenant = GetTenants -n $TenantName
    $projectEnvironments = $tenant.ProjectEnvironments

    if(-not $ProjectName) {
        $ProjectName = UserSelectProject -Inclusions $projectEnvironments.Keys
    }
    $project = GetProjects -n $ProjectName

    $projectEnvironments.remove($project.Id)
    $tenant.ProjectEnvironments = $projectEnvironments

    Update-OctopusResource -Resource $tenant
}

function AddTenantToEnvironment {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("t")] [string]$TenantName,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string]$EnvironmentName
    )
    if(-not $TenantName) {
        $TenantName = UserSelectTenant
    }
    $tenant = GetTenants -n $TenantName
    $projectEnvironments = $tenant.ProjectEnvironments

    $projectId = $projectEnvironments.Keys | Select-Object -First 1
    $environmentIds = $projectEnvironments[$projectId]

    if(-not $EnvironmentName) {
        $EnvironmentName = UserSelectEnvironment -Exclusions $environmentIds
    }

    $newEnvironment = GetEnvironmentObjects -EnvironmentNames $EnvironmentName
    $environmentIds.Add($newEnvironment.Id)
    $environmentObjects = GetEnvironmentObjects -EnvironmentIds $environmentIds

    $projectObjects = GetProjectObjects -ProjectIds $projectEnvironments.Keys
    foreach($proj in $projectObjects) {
        $tenant.ConnectToProjectAndEnvironments($proj,$environmentObjects)
    }
    
    Update-OctopusResource -Resource $tenant
}

function RemoveTenantFromEnvironment {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("t")] [string]$TenantName,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string]$EnvironmentName
    )
    if(-not $TenantName) {
        $TenantName = UserSelectTenant
    }
    $tenant = GetTenants -n $TenantName
    $projectEnvironments = $tenant.ProjectEnvironments

    $projectId = $projectEnvironments.Keys | Select-Object -First 1
    $environmentIds = $projectEnvironments[$projectId]

    if(-not $EnvironmentName) {
        $EnvironmentName = UserSelectEnvironment -Inclusions $environmentIds
    }

    $newEnvironment = GetEnvironmentObjects -EnvironmentNames $EnvironmentName
    $environmentIds.Remove($newEnvironment.Id)
    $environmentObjects = GetEnvironmentObjects -EnvironmentIds $environmentIds

    $projectObjects = GetProjectObjects -ProjectIds $projectEnvironments.Keys
    foreach($proj in $projectObjects) {
        $tenant.ConnectToProjectAndEnvironments($proj,$environmentObjects)
    }
    
    Update-OctopusResource -Resource $tenant
}

function DeleteTenant {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("t")] [string]$TenantName
    )
    if(confirmation) {
        GetTenants -TenantName $TenantName | Remove-OctopusResource
    }
}

function GetTenants {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("n")] [string]$TenantName,
        [Parameter(Position=2, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=3, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    if($TenantName) {
        return Get-OctopusTenant -name $TenantName -ResourceOnly
    }
    else {
        if($Inclusions) {
            return GetTenants | Where-Object {$_.Name -In $Inclusions}
        }
        elseif($Exclusions) {
            return GetTenants | Where-Object {$_.Name -NotIn $Exclusions}
        }
        else {
            return Get-OctopusTenant -ResourceOnly
        }
    }
}

function UserSelectTenant {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $tenants = GetTenants -Inclusions $Inclusions -Exclusions $Exclusions
    return UserSelectPrompt $tenants "Tenant"
}

function UserSelectTenants {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("i")] [string[]]$Inclusions,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$Exclusions
    )
    $tenants = GetTenants -Inclusions $Inclusions -Exclusions $Exclusions
    return UserMultiSelectPrompt $tenants "Tenant"
}

<#
    Release and Deploy functions
#>

function CreateRelease {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("p")] [string[]]$ProjectNames,
        [Parameter(Position=2, Mandatory=$false)] [alias("e")] [string[]]$EnvironmentName,
        [Parameter(Position=3, Mandatory=$false)] [alias("t")] [string[]]$TenantName,
        [Parameter(Position=4, Mandatory=$false)] [alias("x")] [string]$OctoExe,
        [Parameter(Position=5, Mandatory=$false)] [alias("u")] [string]$OctopusURL,
        [Parameter(Position=6, Mandatory=$false)] [alias("a")] [string]$OctopusAPIKey
    )
    if($OctoExe) {
        $env:OctoExe = $OctoExe
    }
    if($OctopusURL) {
        $env:OctopusURL = $OctopusURL
    }
    if($OctopusAPIKey) {
        $env:octopusAPIKey = $OctopusAPIKey
    }

    if(-not $env:OctoExe -or -not $env:OctopusURL -or -not $env:octopusAPIKey) {
        Write-Host "Make sure to set env:OctoExe, env:OctopusURL, and env:octopusAPIKey variables first"
    }
    
    $OctoExe = $env:OctoExe

    if(-not $ProjectNames) {
        $ProjectNames = UserSelectProjects
    }
	
    if(-not $EnvironmentName) {
        $EnvironmentName = UserSelectEnvironments
    }
	
	Log "Creating release for desired projects ..."
	foreach($proj in $ProjectNames) {
        if($TenantName) {
            foreach($tenant in $TenantName) {
                $tenantArgs += " --tenant $tenant"
            }
            Log ("Creating release for "+($TenantName -join ',')+" on project $proj . . .")
            Invoke-Expression "$OctoExe create-release --project '$proj' --server $env:octopusURL --apikey $env:octopusAPIKey $tenantArgs"
            
            Log ("Deploying release for "+($TenantName -join ',')+" on project $proj . . .")
		    Invoke-Expression "$OctoExe deploy-release --project '$proj' --server $env:octopusURL --apikey $env:octopusAPIKey --version latest --deployto '$EnvironmentName' $tenantArgs"
        }
        else {
            Log "Creating release for all tenants on project $proj . . ."
            Invoke-Expression "$OctoExe create-release --project '$proj' --server $env:octopusURL --apikey $env:octopusAPIKey --tenant *"
            
            Log "Deploying release for all tenants on project $proj . . ."
		    Invoke-Expression "$OctoExe deploy-release --project '$proj' --server $env:octopusURL --apikey $env:octopusAPIKey --version latest --deployto '$EnvironmentName' --tenant *"
        }
	}
}

function DeleteReleases {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("p")] [string]$ProjectName,
        [Parameter(Position=2, Mandatory=$false)] [alias("s")] [string]$StartRelease,
        [Parameter(Position=3, Mandatory=$false)] [alias("e")] [string]$EndRelease,
        [Parameter(Position=4, Mandatory=$false)] [alias("x")] [string]$OctoExe,
        [Parameter(Position=5, Mandatory=$false)] [alias("u")] [string]$OctopusURL,
        [Parameter(Position=6, Mandatory=$false)] [alias("a")] [string]$OctopusAPIKey
    )
    if(-not $OctoExe) {
        $env:OctoExe = $OctoExe
    }
    if(-not $OctopusURL) {
        $env:OctopusURL = $OctopusURL
    }
    if(-not $OctopusAPIKey) {
        $env:octopusAPIKey = $OctopusAPIKey
    }

    if(-not $env:OctoExe -or -not $env:OctopusURL -or -not $env:octopusAPIKey) {
        Write-Host "Make sure to set env:OctoExe, env:OctopusURL, and env:octopusAPIKey variables first"
    }
    else {
        if(-not $ProjectName) {
            #List projects and have the user select one
            $projects = GetProjects
            $projectList = ""

            foreach($proj in $projects) {
                $projectList += "`n " + $index + " : " + $proj.Name
                $index++
            }
            $selectedProject = Read-Host "`nPlease select a Project to delete releases from $projectList`n"
            
            $ProjectName = $projects[$selectedProject].Name
        }

        if(-not $StartRelease -or -not $EndRelease) {
            if(-not $StartRelease) {
                $StartRelease = Read-Host "Enter starting Release number"
            }
            if (-not $EndRelease) {
                $EndRelease = Read-Host "Enter ending Release number"
            }
        }
        
        if($StartRelease -gt $EndRelease) {
            Error "The starting release ($StartRelease) cannot be greater than the ending release ($EndRelease). Please enter correct range values."
            DeleteReleases -p $ProjectName
        }

        $confirmation = Read-Host "Delete releases $StartRelease through $EndRelease for $ProjectName?"
        if($confirmation -eq "y") {
            Log "Deleting releases $StartRelease through $EndRelease for $ProjectName..."
            Invoke-Expression "$env:OctoExe delete-releases --project Web --minversion=$StartRelease --maxversion=$EndRelease --server $env:OctopusURL --apikey $env:octopusAPIKey"
        }
        else {
            DeleteReleases -p $ProjectName
        }
    }
}

<#
    Generic helper functions
#>

function UserSelectPrompt {
    param (
        [Parameter(Position=1, Mandatory=$true)] [alias("l")] [System.Collections.ArrayList]$ObjectList,
        [Parameter(Position=2, Mandatory=$true)] [alias("t")] [string]$Type
    )
    $index = 0
	$list = ""

	foreach($option in $ObjectList) {
		$list += "`n " + $index + " : " + $option.Name
		$index++
	}
	$userSelection = Read-Host "`nPlease select a $Type $list`n"
	
	if($userSelection -in 0..($index-1)) {
		return $ObjectList[$userSelection].Name	
	}
	else {
		Write-Host "Incorrect input, please select an item from the list." -f Red
		UserSelectPrompt $ObjectList $Type
	}
}

function UserMultiSelectPrompt {
    param (
        [Parameter(Position=1, Mandatory=$true)] [alias("l")] [System.Collections.ArrayList]$ObjectList,
        [Parameter(Position=2, Mandatory=$true)] [alias("t")] [string]$Type
    )
    [System.Collections.ArrayList]$SelectedObjects = @()
    $loop = $true

    while($loop) {
        $selection = UserSelectPrompt $ObjectList $Type
		$SelectedObjects.Add($selection) | Out-Null
		$ObjectList.Remove(($ObjectList | Where-Object { $_.Name -eq $selection})) | Out-Null
        $moreObjects = Read-Host "`nWould you like to add another $Type? (Y/N)"
		if($moreObjects -ne "y") {
			$loop = $false
		}
    }

    return $SelectedObjects
}

function Confirmation {
    param (
        [Parameter(Position=1, Mandatory=$false)] [alias("c")] [string]$ConfirmationMessage
    )
    if($ConfirmationMessage) {
        $response = Read-Host $ConfirmationMessage
    }
    else {
        $response = Read-Host "Are you sure this is what you want to do? (y/n)"
    }

    if($response -eq "y") {
        return $true
    }
    else {
        return $false
    }
}