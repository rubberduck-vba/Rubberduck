# The parameters should be supplied by the Build event of the project
# in order to take macros from Visual Studio to avoid hard-coding
# the paths. To simplify the process, the project should have a 
# reference to the projects that needs to be registered, so that 
# their DLL files will be present in the $(TargetDir) macro. 
#
# Possible syntax for Post Build event of the project to invoke this:
# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe 
#  -command "$(ProjectDir)BuildRegistryScript.ps1 
#  -config '$(ConfigurationName)' 
#  -builderAssemblyPath '$(TargetPath)' 
#  -netToolsDir '$(FrameworkSDKDir)bin\NETFX 4.6.1 Tools\' 
#  -wixToolsDir '$(SolutionDir)packages\WiX.Toolset.3.9.1208.0\tools\wix\' 
#  -sourceDir '$(TargetDir)' 
#  -targetDir '$(TargetDir)' 
#  -projectDir '$(ProjectDir)'
#  -includeDir '$(ProjectDir)InnoSetup\Includes\'
#  -filesToExtract 'Rubberduck.dll'"
param (
	[Parameter(Mandatory=$true)][string]$config,
	[Parameter(Mandatory=$true)][string]$builderAssemblyPath,
	[Parameter(Mandatory=$true)][string]$netToolsDir,
	[Parameter(Mandatory=$true)][string]$wixToolsDir,
	[Parameter(Mandatory=$true)][string]$sourceDir,
	[Parameter(Mandatory=$true)][string]$targetDir,
	[Parameter(Mandatory=$true)][string]$projectDir,
	[Parameter(Mandatory=$true)][string]$includeDir,
	[Parameter(Mandatory=$true)][string]$filesToExtract
)

function Get-ScriptDirectory
{
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value;
  Split-Path $Invocation.MyCommand.Path;
}

# Invokes a Cmd.exe shell script and updates the environment.
function Invoke-CmdScript {
  param(
    [String] $scriptName
  )
  $cmdLine = """$scriptName"" $args & set"
  & $Env:SystemRoot\system32\cmd.exe /c $cmdLine |
  select-string '^([^=]*)=(.*)$' | foreach-object {
    $varName = $_.Matches[0].Groups[1].Value
    $varValue = $_.Matches[0].Groups[2].Value
    set-item Env:$varName $varValue
  }
}

# Returns the current environment.
function Get-Environment {
  get-childitem Env:
}

# Restores the environment to a previous state.
function Restore-Environment {
  param(
    [parameter(Mandatory=$TRUE)]
      [System.Collections.DictionaryEntry[]] $oldEnv
  )
  # Remove any added variables.
  compare-object $oldEnv $(Get-Environment) -property Key -passthru |
  where-object { $_.SideIndicator -eq "=>" } |
  foreach-object { remove-item Env:$($_.Name) }
  # Revert any changed variables to original values.
  compare-object $oldEnv $(Get-Environment) -property Value -passthru |
  where-object { $_.SideIndicator -eq "<=" } |
  foreach-object { set-item Env:$($_.Name) $_.Value }
}

# Remove older imported registry scripts for debug builds.
function Clean-OldImports
{
	param(
		[String] $dir
	)
	$i = 0;
	Get-ChildItem $dir -Filter DebugRegistryEntries.reg.imported_*.txt | 
	Sort-Object Name -Descending |
	Foreach-Object {
		if($i -ge 10) {
			$_.Delete();
		}
		$i++;
	}
}

Set-StrictMode -Version latest;
$ErrorActionPreference = "Stop";
$DebugUnregisterRun = $false;

try
{
	# Clean imports older than 10 builds
	Clean-OldImports ((Get-ScriptDirectory) + "\LocalRegistryEntries");;

	# Allow multiple DLL files to be registered if necessary
	$separator = "|";
	$option = [System.StringSplitOptions]::RemoveEmptyEntries;
	$files = $filesToExtract.Split($separator, $option);
	
	# Load the Deployment DLL
	[System.Reflection.Assembly]::LoadFrom($builderAssemblyPath);

	# Determine if MIDL is available for building
	$devPath = $Env:ProgramFiles + "*\Microsoft Visual Studio\*\*\Common*\Tools\VsDevCmd.bat";
	$devPath = Resolve-Path -Path $devPath;
	if($devPath)
	{
		# Additional verifications as some versions of VsDevCmd.bat might not initialize the environment for C++ build tools
		$result = Get-Module -ListAvailable -Name "VSSetup" -ErrorAction SilentlyContinue;
		if(!$result)
		{
			Write-Warning "VSSetup not installed; extracting...";
			$moduleDirPath = "$([Environment]::GetFolderPath("MyDocuments"))\WindowsPowerShell";
			if(!(Test-Path -Path $moduleDirPath -PathType Container))
			{
				Write-Warning "WindowsPowerShell directory not found in user's documents. Creating.";
				New-Item -Path $moduleDirPath -ItemType Directory;
			}
			$moduleDirPath += "\Modules";
			if(!(Test-Path -Path $moduleDirPath -PathType Container))
			{
				Write-Warning "WindowsPowerShell\Modules directory not found in user's documents. Creating.";
				New-Item -Path $moduleDirPath -ItemType Directory;
			}
			$moduleDirPath += "\VSSetup";
			if(!(Test-Path -Path $moduleDirPath -PathType Container))
			{
				Write-Warning "WindowsPowerShell\Modules\VSSetup directory not found in user's documents. Creating.";
				New-Item -Path $moduleDirPath -ItemType Directory;
			}
			# Sanity check
			if(!(Test-Path -Path $moduleDirPath -PathType Container))
			{
				Write-Error "WindowsPowerShell\Modules\VSSetup directory still not found in user's documents after attempt to create it. Cannot continue";
				throw [System.IO.DirectoryNotFoundException] "Cannot create or locate the directory at path '$moduleDirPath'";
			}
			Expand-Archive "$projectDir\OleWoo\VSSetup.zip" $moduleDirPath -Force;
		}

		try {
			Import-Module VSSetup -Force:$true;
			$result = Get-VSSetupInstance | Select-VSSetupInstance -Latest -Require Microsoft.VisualStudio.Component.VC.Tools.x86.x64;
		} catch {
			$result = $null;
			Write-Warning "Error occurred with using VSSetup module";
			Write-Error ($_);
		}

		if(!$result)
		{
			$devPath = $null;
			Write-Warning "Cannot locate the VS Setup instance capable of building with C++ build tools";
		}
	}

	if(!$devPath)
	{
		Write-Warning "Cannot locate the VsDevCmd.bat to initialize C++ build tools; falling back to tlbexp.exe....";
	}

	Write-Host "";

	foreach($file in $files)
	{
		Write-Host "Processing '$file'";
		Write-Host "";

		$dllFile = [System.String]$file;
		$idlFile = [System.String]($file -replace ".dll", ".idl");
		$tlb32File = [System.String]($file -replace ".dll", ".x32.tlb");
		$tlb64File = [System.String]($file -replace ".dll", ".x64.tlb");

		$sourceDll = $sourceDir + $file;
		$targetDll = $targetDir + $file;
		$sourceTlb32 = $sourceDir + $tlb32File;
		$targetTlb32 = $targetDir + $tlb32File;
		$sourceTlb64 = $sourceDir + $tlb64File;
		$targetTlb64 = $targetDir + $tlb64File;
		$dllXml = $targetDll + ".xml";
		$tlbXml = $targetTlb32 + ".xml";

		# Write-Host "Variable printout:"
		# Write-Host "dllFile = $dllFile";
		# Write-Host "idlFile = $idlFile";
		# Write-Host "tlb32File = $tlb32File";
		# Write-Host "tlb64File = $tlb64File";
		# Write-Host "sourceDll = $sourceDll";
		# Write-Host "targetDll = $targetDll";
		# Write-Host "sourceTlb32 = $sourceTlb32";
		# Write-Host "targetTlb32 = $targetTlb32";
		# Write-Host "sourceTlb64 = $sourceTlb64";
		# Write-Host "targetTlb64 = $targetTlb64";
		# Write-Host "dllXml = $dllXml";
		# Write-Host "tlbXml = $tlbXml";
		# Write-Host "targetDir = $targetDir";
		# Write-Host "";

		# Use for debugging issues with passing parameters to the external programs
		# Note that it is not legal to have syntax like `& $cmdIncludingArguments` or `& $cmd $args`
		# For simplicity, the arguments are pass in literally.
		# & "C:\GitHub\Rubberduck\Rubberduck\Rubberduck.Deployment\echoargs.exe" ""$sourceDll"" /win32 /out:""$sourceTlb"";
		
		# Compile TLB files using MIDL
		if($devPath)
		{
			$idlGenerator = New-Object Rubberduck.Deployment.IdlGeneration.IdlGenerator;
		
			$idl = $idlGenerator.GenerateIdl($sourceDll);
			$encoding = New-Object System.Text.UTF8Encoding $true;
			[System.IO.File]::WriteAllLines($idlFile, $idl, $encoding);
		
			$origEnv = Get-Environment;
			try {
				Invoke-CmdScript "$devPath";
				
				if($targetDir.EndsWith("\"))
				{
					$targetDirWithoutSlash = $targetDir.Substring(0,$targetDir.Length-1);
				}
				else
				{
					$targetDirWithoutSlash = $targetDir;
				}

				& midl.exe /win32 /tlb ""$tlb32File"" ""$idlFile"" /out ""$targetDirWithoutSlash"";
				& midl.exe /amd64 /tlb ""$tlb64File"" ""$idlFile"" /out ""$targetDirWithoutSlash"";
			} catch {
				throw;
			} finally {
				Restore-Environment $origEnv;
			}
		}

		# Compile TLB files using tlbexp.exe
		if(!$devPath)
		{
			$cmd = "{0}tlbexp.exe" -f $netToolsDir;
			& $cmd ""$sourceDll"" /win32 /out:""$sourceTlb32"";
			& $cmd ""$sourceDll"" /win64 /out:""$sourceTlb64"";
		}

		# Harvest both DLL and TLB files using WiX's heat.exe, generating XML files
		$cmd = "{0}heat.exe" -f $wixToolsDir;
		& $cmd file ""$sourceDll"" -out ""$dllXml"";
		& $cmd file ""$sourceTlb32"" -out ""$tlbXml"";
		
		# Initialize the registry builder with the provided XML files
		$builder = New-Object Rubberduck.Deployment.Builders.RegistryEntryBuilder;
		$entries = $builder.Parse($tlbXml, $dllXml);

		# For debugging
		# $entries | Format-Table | Out-String |% {Write-Host $_};
		
		$writer = New-Object Rubberduck.Deployment.Writers.InnoSetupRegistryWriter;
		$content = $writer.Write($entries, $dllFile, $tlb32File, $tlb64File);
		
		# The file must be encoded in UTF-8 BOM
		$regFile = ($includeDir + ($file -replace ".dll", ".reg.iss"));
		$encoding = New-Object System.Text.UTF8Encoding $true;
		[System.IO.File]::WriteAllLines($regFile, $content, $encoding);
		$content = $null;

		# Register the debug build on the local machine
		if($config -eq "Debug")
		{
			if(!$DebugUnregisterRun) 
			{
				# First see if there are registry script from the previous build
				# If so, execute them to delete previous build's keys (which may
				# no longer exist for the current build and thus won't be overwritten)
				$dir = ((Get-ScriptDirectory) + "\LocalRegistryEntries");
				$regFileDebug = $dir + "\DebugRegistryEntries.reg";

				if (Test-Path -Path $dir -PathType Container)
				{
					if (Test-Path -Path $regFileDebug -PathType Leaf)
					{
						$datetime = Get-Date;
						if ([Environment]::Is64BitOperatingSystem)
						{
							& reg.exe import $regFileDebug /reg:32;
							& reg.exe import $regFileDebug /reg:64;
						}
						else 
						{
							& reg.exe import $regFileDebug;
						}
						& reg.exe import ($dir + "\RubberduckAddinRegistry.reg");
						Move-Item -Path $regFileDebug -Destination ($regFileDebug + ".imported_" + $datetime.ToUniversalTime().ToString("yyyyMMddHHmmss") + ".txt" );
					}
				}
				else
				{
					New-Item $dir -ItemType Directory;
				}

				$DebugUnregisterRun = $true;
			}

			# NOTE: The local writer will perform the actual registry changes; the return
			# is a registry script with deletion instructions for the keys to be deleted
			# in the next build.
			$writer = New-Object Rubberduck.Deployment.Writers.LocalDebugRegistryWriter;
			$content = $writer.Write($entries, $dllFile, $tlb32File, $tlb64File);

			$encoding = New-Object System.Text.ASCIIEncoding;
			[System.IO.File]::AppendAllText($regFileDebug, $content, $encoding);
		}

		Write-Host "Finished processing '$file'";
		Write-Host "";
	}
	
	Write-Host "Finished processing all files";
}
catch
{
	Write-Error ($_);
	# Cause the build to fail
	throw;
}

# for debugging locally
# Write-Host "Press any key to continue...";
# Read-Host -Prompt "Press Enter to continue";

