$ProgramFilesPath = "C:\Program Files (x86)"
$global:VSVersion = "14.0"
# Set the Error action so that all errors cause the script to fail
$ErrorActionPreference = "Stop"

function Log($Message) {
	Write-Host "$Message`n" -f cyan
}

function Warning($Message) {
	Write-Host "`n*** $Message ***`n" -f Yellow
}

function Error($Message) {
	Write-Host "`n*** $Message ***`n" -f red
}

function ThrowDetailedException($Exception) {
	$Exception | format-list -force | Out-String | ForEach-Object {Write-Host $_ -f Red}
}

# Check if the script is being run in admin mode, if not exit the program
function CheckAdmin {
	if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
		Error "This script must be run in admin mode. Please open PowerShell in admin mode and rerun."
		exit
	}
}

# Check .NET version
function CheckDotNetVersion {
	param ([Parameter(Position=1, Mandatory=$true)] [double]$Version)
	$DotNetVersion = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -recurse | Get-ItemProperty -name Version,Release -EA 0 | Where-Object {$_.PSChildName -eq "v$Version"} | Select-Object PSChildName
	if (-not ($DotNetVersion)) {
		Error "This script requires that .NET version $Version be installed. Please install .NET version $Version and rerun."
		exit
	}
}

# Search for an executable in a given search path
function Get-Executable {
	param(	[Parameter(Position=1, Mandatory=$true)] [string]$ExecutableName,
			[Parameter(Position=2, Mandatory=$false)] [string]$SearchPath = $ProgramFilesPath)
	return (get-childitem $SearchPath -Recurse -ErrorAction SilentlyContinue | Where-Object {$_.name -eq $ExecutableName} | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}

# Gets the MSBuild executable depending on the version of VS installed
function Get-Msbuild {
	$MsBuildPath = "hklm:\software\Microsoft\MSBuild\ToolsVersions\4.0"
	$vswherePath = "$ProgramFilesPath\Microsoft Visual Studio\Installer\vswhere.exe"
	
	if(Test-Path -Path $vswherePath) {
		$path = Invoke-Expression "& '$vswherePath' -latest -products * -requires Microsoft.Component.MSBuild -property installationPath"
		if ($path) {
			$path = join-path $path 'MSBuild\15.0\Bin'
			if (test-path $path) {
				$MsBuildPath = $path
				Set-Variable -Name "VSVersion" -Value "15.0" -Scope Global
			}
		}
	}
	
	if(-Not (Test-Path -Path $MsBuildPath)) { Throw [System.IO.FileNotFoundException] "Cannot find MSBuild tool path at $MsBuildPath." }
	if($VSVersion -eq "14.0") {
		$MsBuildPath = (Get-ItemProperty $MsBuildPath).MSBuildToolsPath
	}
	
	return $MsBuildPath + (&{if($MsBuildPath -match '\\$') {""} else { "\" }}) + "msbuild.exe"
}

# Gets the MSTest executable depending on the version of VS installed
function Get-Mstest {
	$extension = "\Common7\IDE"
	if($VSVersion -eq "14.0") {
		$msTestPath = "$ProgramFilesPath\Microsoft Visual Studio 14.0" + $extension
	}
	else {
		$vswherePath = "$ProgramFilesPath\Microsoft Visual Studio\Installer\vswhere.exe"
		$path = Invoke-Expression "& '$vswherePath' -latest -products * -requires Microsoft.VisualStudio.PackageGroup.TestTools.Core -property installationPath"
		if ($path) {
			$msTestPath = $path + $extension
		}
	}
	
	if(-Not (Test-Path -Path $msTestPath)) { Throw [System.IO.FileNotFoundException] "Cannot find MSTest tool path at $msTestPath." }
	
	return $msTestPath + "\mstest.exe"
}

# Import the necessary SQL Server assemblies for managing databases
function ImportSqlServerAssemblies {
	Log "Importing necessary assemblies..."
	
	$sqlps = get-executable -ExecutableName "SQLPS.PS1" -SearchPath "$ProgramFilesPath\Microsoft SQL Server"
	if(-Not (Test-Path -Path $sqlps)) { Throw [System.IO.FileNotFoundException] "Connot find SQLPS.PS1 tool path in $sqlps. Make sure you have SQL Server Management Studio installed." }
	Import-Module $sqlps

	$assemblylist =
		"Microsoft.SqlServer.Management.Common",
		"Microsoft.SqlServer.Smo",
		"Microsoft.SqlServer.SmoExtended"

	foreach ($asm in $assemblylist)
	{
		$asm = [Reflection.Assembly]::LoadWithPartialName($asm)
	}
}

# Get a Sql connection string to a server
function GetSqlConnectionString {
    param(
        [Parameter(Position=1, Mandatory=$true)] [string]$ServerName
    )
    return "Data Source=$ServerName;Initial Catalog=msdb;Integrated Security=SSPI;"
}

# Get a Sql connection from a server name
function GetSqlConnection {
    param(
        [Parameter(Position=1, Mandatory=$true)] [string]$ServerName
    )
    $SqlConnectionString = GetSqlConnectionString -ServerName $ServerName
    return New-Object System.Data.SqlClient.SqlConnection $SqlConnectionString
}

# Deploy a given database to a server
function DeployDatabase {
	param(
		[Parameter(Position=1, Mandatory=$true)] [string]$SqlPackage,
		[Parameter(Position=2, Mandatory=$true)] [string]$DacpacPath,
		[Parameter(Position=3, Mandatory=$false)] [string]$Server = "localhost",
		[Parameter(Position=4, Mandatory=$false)] [string]$TargetDatabaseName,
        [Parameter(Position=5, Mandatory=$false)] [bool]$CreateNewDatabase = $false,
        [Parameter(Position=6, Mandatory=$false)] [bool]$BlockOnPossibleDataLoss = $false,
        [Parameter(Position=7, Mandatory=$false)] [bool]$DropObjectNotInSource = $true,
        [Parameter(Position=8, Mandatory=$false)] [bool]$GenerateSmartDefaults = $true,
        [Parameter(Position=9, Mandatory=$false)] [bool]$IncludeTransactionalScripts = $true,
		[Parameter(Position=10, Mandatory=$false)] [string]$Variables)
	
	# Validate parameters
	if(-Not (Test-Path -Path $SqlPackage)) { Throw [System.IO.FileNotFoundException] "Cannot find sqlpackage.exe at $SqlPackage." }
	if(-Not (Test-Path -Path $DacpacPath)) { Throw [System.IO.FileNotFoundException] "Cannot find dacpac at $DacpacPath. Make sure the package was built successfully." }
	else { $DacpacPath = Resolve-Path -Path $DacpacPath }
	
	Log "Deploying $TargetDatabaseName using $DacpacPath..."
	$cmd = "& '$SqlPackage' /Action:Publish /SourceFile:'$DacpacPath' /TargetDatabaseName:$TargetDatabaseName /TargetServerName:$Server /p:CreateNewDatabase=$CreateNewDatabase /p:IncludeTransactionalScripts=$IncludeTransactionalScripts /p:DropObjectsNotInSource=$DropObjectNotInSource /p:GenerateSmartDefaults=$GenerateSmartDefaults /p:BlockOnPossibleDataLoss=$BlockOnPossibleDataLoss"
	if($Variables -ne "") { 
		foreach ($var in $Variables.Split(",")){
			$cmd = "$cmd /v:$var"
		}
	}
	Log $cmd
	Invoke-Expression $cmd
	Log "Deployment complete. $TargetDatabaseName deployed successfully."
}

# Deploy a given database to a server
function DeployDatabaseFromProfile {
	param(
		[Parameter(Position=1, Mandatory=$true)] [string]$SqlPackage,
		[Parameter(Position=2, Mandatory=$true)] [string]$DacpacPath,
		[Parameter(Position=3, Mandatory=$true)] [string]$ProfilePath,
		[Parameter(Position=4, Mandatory=$false)] [string]$Server = "localhost",
		[Parameter(Position=5, Mandatory=$false)] [string]$InstanceName)
	
	# Validate parameters
	if(-Not (Test-Path -Path $SqlPackage)) { Throw [System.IO.FileNotFoundException] "Cannot find sqlpackage.exe at $SqlPackage." }
	if(-Not (Test-Path -Path $DacpacPath)) { Throw [System.IO.FileNotFoundException] "Cannot find dacpac at $DacpacPath. Make sure the package was built successfully." }
	if(-Not (Test-Path -Path $ProfilePath)) { Throw [System.IO.FileNotFoundException] "Cannot find profile at $ProfilePath. Make sure the profile exists." }
	else { $DacpacPath = Resolve-Path -Path $DacpacPath }

	if($InstanceName) {
		$Server = "$Server\$InstanceName"
	}
	
	Log "Deploying database from $ProfilePath using $DacpacPath..."
	$cmd = "& '$SqlPackage' /Action:Publish /SourceFile:'$DacpacPath' /Profile:'$ProfilePath' /TargetServerName:'$Server'"
	Log $cmd
	Invoke-Expression $cmd
	Log "Deployment complete. Database from $ProfilePath deployed successfully."
}

# Drop a database from a server
function DropDatabase {
	param(
		[Parameter(Position=1, Mandatory=$true)] [string]$ServerName,
        [Parameter(Position=2, Mandatory=$true)] [string]$DatabaseName
    )
    Log "Dropping $DatabaseName from $ServerName..."
	
	$server = new-object Microsoft.SqlServer.Management.Smo.Server $ServerName;

	$existingDatabase = $server.Databases[$DatabaseName]
	if($existingDatabase) {
		$server.KillDatabase($DatabaseName)
		Log "Dropped database $DatabaseName from $ServerName successfully."
	}
	else {
		Warning "Database $DatabaseName does not exist on $ServerName"
	}
}

# Restore a database from a file to a server
function RestoreDatabase {
	param(
		[Parameter(Position=1, Mandatory=$true)] [string]$ServerName,
        [Parameter(Position=2, Mandatory=$true)] [string]$DatabaseName,
        [Parameter(Position=3, Mandatory=$true)] [string]$BackupPath
    )
    Log "Restoring $DatabaseName to $ServerName..."
	
	$server = new-object Microsoft.SqlServer.Management.Smo.Server $ServerName;

	$existingDatabase = $server.Databases[$DatabaseName]
	if($existingDatabase) {
		DropDatabase -ServerName $ServerName -DatabaseName $DatabaseName
	}

    $DataFolder = $server.Settings.DefaultFile;
    $LogFolder = $server.Settings.DefaultLog;

    if ($DataFolder.Length -eq 0) {
        $DataFolder = $server.Information.MasterDBPath;
    }

    if ($LogFolder.Length -eq 0) {
        $LogFolder = $server.Information.MasterDBLogPath;
    }
	
    $backupDeviceItem = new-object Microsoft.SqlServer.Management.Smo.BackupDeviceItem $BackupPath, 'File';
    
    $restore = new-object 'Microsoft.SqlServer.Management.Smo.Restore';
    $restore.Database = $DatabaseName;
	$restore.ReplaceDatabase = $true;
    $restore.FileNumber = 1;
    $restore.Devices.Add($backupDeviceItem);
    $dataFileNumber = 0;

    foreach ($file in $restore.ReadFileList($server)) 
    {
        $relocateFile = new-object 'Microsoft.SqlServer.Management.Smo.RelocateFile';
        $relocateFile.LogicalFileName = $file.LogicalName;

        if ($file.Type -eq 'D'){
            if($dataFileNumber -ge 1)
            {
                $suffix = "_$dataFileNumber";
            }
            else
            {
                $suffix = $null;
            }
            $relocateFile.PhysicalFileName = $dataFolder+$DatabaseName+$suffix+"_Primary.mdf";

            $dataFileNumber ++;
        }
        else {
            $relocateFile.PhysicalFileName = $logFolder+$DatabaseName+"_Primary.ldf";
        }
        $restore.RelocateFiles.Add($relocateFile) | Out-Null;
    }
    Try {
        $restore.SqlRestore($server);
    } Catch {
        $message = $_.Exception.Message
        $innerException = $_.Exception.InnerException
        Warning $message
        Warning $innerException
        Warning "If error states 'The operating system returned the error '32(The process cannot access the file because it is being used by another process.)''"
        Warning "This is a known issue as shown here: https://support.microsoft.com/en-us/help/3153836/operating-system-error-32-when-you-restore-a-database-in-sql-server-20"
        Throw
    }
    
    Log "Restore of $DatabaseName to $ServerName completed successfully."
}

# Deploy an SSIS package to the SSIS catalog
function DeploySsis {
	param(	
		[Parameter(Position=1, Mandatory=$true)] [string]$IspacFilePath,
		[Parameter(Position=2, Mandatory=$true)] [string]$IsFolder,
		[Parameter(Position=3, Mandatory=$true)] [string]$ProjectName,
		[Parameter(Position=4, Mandatory=$true)] [string]$ServerName,
		[Parameter(Position=5, Mandatory=$false)] [string]$SsisCatalog = "SSISDB",
		[Parameter(Position=6, Mandatory=$false)] [string]$CatalogPwd = "P@ssw0rd1"
	)
	
	$ssisNamespace = "Microsoft.SqlServer.Management.IntegrationServices"

	# Validate parameters
	if(-Not (Test-Path -Path $IspacFilePath)) { Throw [System.IO.FileNotFoundException] "Cannot find dacpac at $IspacFilePath. Make sure the package was built successfully." }
	
	Log "Deploying Staging ISPac for $ProjectName using $IspacFilePath..."
	
	# Load the IntegrationServices Assembly  
	[Reflection.Assembly]::LoadWithPartialName($ssisNamespace) | Out-Null
	$SqlConnection = GetSqlConnection -ServerName $ServerName
	$integrationServices = New-Object $ssisNamespace".IntegrationServices" $SqlConnection
	# Check if connection succeeded
	if (-not $integrationServices) { Throw [System.Exception] "Failed to connect to server $integrationServices " }
	else { Log "Connected to server $integrationServices" }
	
	# Enable common language runtime (CLR) to create catalog and project folder
	Invoke-Sqlcmd -server $ServerName -database master -query "sp_configure 'clr enabled', 1; RECONFIGURE;"
	
	$catalog = $integrationServices.Catalogs[$SsisCatalog]
	# Check if the SSISDB Catalog exists
	if (-not $catalog) {
		Log "Creating SSIS Catalog '$SsisCatalog'..."  

		# Provision a new SSIS Catalog  
		$catalog = New-Object $ssisNamespace".Catalog" ($integrationServices, $SsisCatalog, $CatalogPwd)
		$catalog.Create()
	}
	
	$folder = $catalog.Folders[$IsFolder]
	# Verify that the catalog folder exists
	if (-not $folder)
	{
		#Create a folder in SSISDB
		Log "Creating Folder '$IsFolder'..."
		$folder = New-Object $ssisNamespace".CatalogFolder" ($catalog, $IsFolder, $IsFolder)            
		$folder.Create()  
	}
		
	# Deploying project to folder
	if($folder.Projects.Contains($ProjectName)) { Log "Deploying $ProjectName to $IsFolder (REPLACE)..." }
	else { Log "Deploying $ProjectName to $IsFolder (NEW)..." }
	
	# Reading ispac file as binary
	[byte[]] $ispacFile = [System.IO.File]::ReadAllBytes($IspacFilePath)
	$folder.DeployProject($ProjectName, $ispacFile)
	$project = $folder.Projects[$ProjectName]
	if (-not $project) { Throw [System.Exception] "Failed to deploy SSIS Project" }
	Log "Staging ISPac deployed successfully."
}

# Execute a Sql Agent job on a given server
function ExecuteJob {
	param (
		[Parameter(Position=1, Mandatory=$true)] [string]$JobName,
		[Parameter(Position=2, Mandatory=$true)] [string]$ServerName
	)
	$sqlConnection = new-object System.Data.SqlClient.SqlConnection 
	$sqlConnection.ConnectionString = GetSqlConnectionString -ServerName $ServerName
	$sqlConnection.Open() 
	$sqlCommand = new-object System.Data.SqlClient.SqlCommand 
	$sqlCommand.CommandTimeout = 120 
	$sqlCommand.Connection = $sqlConnection 
	$sqlCommand.CommandText = "EXEC dbo.sp_start_job '$JobName'"
	Log "Executing job $JobName..." 
	$sqlCommand.ExecuteNonQuery() | Out-Null
	$sqlConnection.Close()
	Warning "$JobName job is running. Please check the Job Activity Monitor for job progress."
}

# Start the SqlAgent of a given server
function StartSqlAgent {
	param (
		[Parameter(Position=1, Mandatory=$true)] [string]$ServerName
	)
	Log "Starting SQL Server Agent..."
	try {
		if($ServerName -ne "localhost") {
			$instance = $ServerName.Split('\')
			if($instance.Length -gt 1) {
				$instance = $instance[1]
            }
            Log "net start 'SQL Server Agent ($instance)'"
			net start "SQL Server Agent ($instance)"
		}
		else {
            Log "net start SQLSERVERAGENT"
			net start SQLSERVERAGENT
		}
		Log "SQL Server Agent started successfully."
	} catch {
        $message = $_.Exception.Message
        if($message -eq "The requested service has already been started." -or $message -eq "System error 5 has occurred.") {
            Log "SQL Server Agent is already running."
        }
        else {
            Error $message
        }
    }
}

function CreateLocalDBInstance {
	param (
		[Parameter(Position=1, Mandatory=$true)] [string]$InstanceName
	)
	Log "Creating localdb instance $InstanceName"
	sqllocaldb create $InstanceName -s
	Log "Successfully created localdb instance $InstanceName"
	sqllocaldb info $InstanceName | format-list | Out-String | ForEach-Object {Write-Host $_ -f cyan}
}

function RemoveLocalDBInstance {
	param (
		[Parameter(Position=1, Mandatory=$true)] [string]$InstanceName
	)
	Log "Stopping localdb instance $InstanceName"
	sqllocaldb stop $InstanceName -k
	Log "Deleting localdb instance $InstanceName"
	sqllocaldb delete $InstanceName
	Log "Cleaning up files from localdb instance $InstanceName"
	$cleanupPath = "$env:USERPROFILE\AppData\Local\Microsoft\Microsoft SQL Server Local DB\Instances\$InstanceName"
	remove-item $cleanupPath -Recurse -Force
	Log "Successfully removed localdb instance $InstanceName"
}

function BuildBiml {
	param(
		[Parameter(Position=1, Mandatory=$true)] [string]$ProjectFile,
		[Parameter(Position=2, Mandatory=$true)] [string]$Options,
		[Parameter(Position=3, Mandatory=$false)] [string]$BimlBuildPath,
		[Parameter(Position=4, Mandatory=$false)] [string]$MsBuildPath
	)	
	
	$ProjectRoot = (Get-Item (Split-Path -Path $ProjectFile)).FullName
	$outputDir = $ProjectRoot + "\output"
	$Options = "LicenseKey=6EX3G-P2JWK-R45US-F2LFG-19E04-DTY9V-V;" + $Options;

	if($BimlBuildPath) {
		$AssemblyPath = $BimlBuildPath
	}
	else {
		$AssemblyPath = (Get-Item $PSScriptRoot).parent.FullName + "\src\lib\biml\5.0"
	}

	if ($MsBuildPath -and (-not ($MsBuildPath -match 'msbuild.exe'))) {
		$MsBuildPath = $MsBuildPath + (&{if($MsBuildPath -match '\\$') {""} else { "\" }}) + "msbuild.exe"
	}
	if(-not $MsBuildPath){
		$MsBuildPath = Get-Msbuild
	}

	$args = @(
		[string]::Format('/p:OutputPath="{0}"', $outputDir)
		, [string]::Format('/p:options="{0}"', $Options)
		, '/p:SqlVersion=SqlServer2016'
		, '/p:SsasVersion=Ssas2016'
		, '/p:SsasTabularVersion=SsasTabular2016'
		, '/p:SsisVersion=Ssis2016'
		, '/p:SsisDeploymentModel=Project'
		, '/p:WarnAsError=False'
		, '/p:Warn=4'
		, '/p:CleanOutputFolder=True'
		, '/p:TaskName=Varigence.Biml.Engine.MSBuild.BimlCompilerTask'
		, [string]::Format('/p:AssemblyFile="{0}\BimlEngine.dll"', $AssemblyPath)
		, [string]::Format('/p:AssemblyPath="{0}"', $AssemblyPath)
		, [string]::Format('/p:ProjectRoot="{0}"', $ProjectRoot)
	)

	Log "Executing MSBuild against $ProjectFile"
	Log "With arguments: $args"
	Log "& $MsBuildPath $ProjectFile $args"

	if(-Not (Test-Path -Path $ProjectFile)) {
		Warning "Project File ($ProjectFile) does not exist"
		return 0
	}

	& $MsBuildPath $ProjectFile $args

	if($lastexitcode) {
		Log "Sending fail trigger to Bamboo..."
		throw "BIML build failure."
	} else {
		Log "No Error reported. BIML build successful."
	}
}

#Convert excel tabs into individual csv's

function ConvertExceltoCSV {

param (
   [Parameter(Position=1,Mandatory=$true)] [string]$InputFolderPath,
   [Parameter(Position=2,Mandatory=$true)] [string]$OrigFileName,
   [Parameter(Position=3,Mandatory=$false)] [array]$TabNames)

    #Create and get Excel Obj
    $excel = New-Object -comobject Excel.Application;
    $excel.visible=$false;
    $excel.DisplayAlerts=$false;
    $WorkbookPath = $InputFolderPath + "\" + $OrigFileName
    $WorkBook = $excel.Workbooks.Open($WorkbookPath);
    $output_type = ".csv";
    $xlCSV = 6;

    # Validate that the file exists
    if(-Not (Test-Path -Path $WorkbookPath)) { Throw [System.IO.FileNotFoundException] "Cannot find $OrigFileName in $InputFolderPath." }

    foreach($ws in $Workbook.Worksheets) {
        if((-not $TabNames) -or ($TabNames -and ($TabNames -contains $ws.Name))) {
            Write-Host "Creating csv for" $ws.name
            $UnprotectSheet = $ws.Unprotect()
            $ws.SaveAs($InputFolderPath + "\" + $ws.Name + $output_type, $xlCSV);
        }
    }
    $Workbook.close()
    $excel.quit()
}
