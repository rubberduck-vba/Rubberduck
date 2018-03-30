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
#  -includeDir '$(ProjectDir)InnoSetup\Includes\'
#  -filesToExtract 'Rubberduck.dll'"
param (
	[Parameter(Mandatory=$true)][string]$config,
	[Parameter(Mandatory=$true)][string]$builderAssemblyPath,
	[Parameter(Mandatory=$true)][string]$netToolsDir,
	[Parameter(Mandatory=$true)][string]$wixToolsDir,
	[Parameter(Mandatory=$true)][string]$sourceDir,
	[Parameter(Mandatory=$true)][string]$targetDir,
	[Parameter(Mandatory=$true)][string]$includeDir,
	[Parameter(Mandatory=$true)][string]$filesToExtract
)

function Get-ScriptDirectory
{
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value
  Split-Path $Invocation.MyCommand.Path
}

Set-StrictMode -Version latest
$ErrorActionPreference = "Stop";

try
{
	# Allow multiple DLL files to be registered if necessary
	$separator = "|"
	$option = [System.StringSplitOptions]::RemoveEmptyEntries
	$files = $filesToExtract.Split($separator, $option);

	foreach($file in $files)
	{
		$sourceDll = $sourceDir + $file;
		$targetDll = $targetDir + $file;
		$sourceTlb32 = $sourceDir + ($file -replace ".dll", ".x32.tlb");
		$targetTlb32 = $targetDir + ($file -replace ".dll", ".x32.tlb");
		$sourceTlb64 = $sourceDir + ($file -replace ".dll", ".x64.tlb");
		$targetTlb64 = $targetDir + ($file -replace ".dll", ".x64.tlb");
		$dllXml = $targetDll + ".xml"
		$tlbXml = $targetTlb32 + ".xml"

		# Use for debugging issues with passing parameters to the external programs
		# Note that it is not legal to have syntax like `& $cmdIncludingArguments` or `& $cmd $args`
		# For simplicity, the arguments are pass in literally.
		# & "C:\GitHub\Rubberduck\Rubberduck\Rubberduck.Deployment\echoargs.exe" ""$sourceDll"" /win32 /out:""$sourceTlb"";
		
		$cmd = "{0}tlbexp.exe" -f $netToolsDir;
		& $cmd ""$sourceDll"" /win32 /out:""$sourceTlb32"";
		& $cmd ""$sourceDll"" /win64 /out:""$sourceTlb64"";

		$cmd = "{0}heat.exe" -f $wixToolsDir;
		& $cmd file ""$sourceDll"" -out ""$dllXml"";
		& $cmd file ""$sourceTlb32"" -out ""$tlbXml"";
			
		[System.Reflection.Assembly]::LoadFrom($builderAssemblyPath);
		$builder = New-Object Rubberduck.Deployment.Builders.RegistryEntryBuilder
	
		$entries = $builder.Parse($tlbXml, $dllXml);

		# For debugging
		# $entries | Format-Table | Out-String |% {Write-Host $_};
		
		$writer = New-Object Rubberduck.Deployment.Writers.InnoSetupRegistryWriter
		$content = $writer.Write($entries);
		 
		$regFile = ($includeDir + ($file -replace ".dll", ".reg.iss"))
		$encoding = New-Object System.Text.UTF8Encoding $False
		[System.IO.File]::WriteAllLines($regFile, $content, $encoding)

		# Register the debug build on the local machine
		if($config -eq "Debug")
		{
			# First see if there are registry script from the previous build
			# If so, execute them to delete previous build's keys (which may
		    # no longer exist for the current build and thus won't be overwritten)
			$dir = ((Get-ScriptDirectory) + "\LocalRegistryEntries");
			$regFile = $dir + "\DebugRegistryEntries.reg";

			if (Test-Path -Path $dir -PathType Container)
			{
				if (Test-Path -Path $regFile -PathType Leaf)
				{
					$datetime = Get-Date;
					if ([Environment]::Is64BitOperatingSystem)
					{
						& reg.exe import $regFile /reg:32;
						& reg.exe import $regFile /reg:64;
					}
					else 
					{
						& reg.exe import $regFile;
					}
					& reg.exe import ($dir + "\RubberduckAddinRegistry.reg");
					Move-Item -Path $regFile -Destination ($regFile + ".imported_" + $datetime.ToUniversalTime().ToString("yyyyMMddHHmmss") + ".txt" )
				}
			}
			else
			{
				New-Item $dir -ItemType Directory
			}
			
			# NOTE: The local writer will perform the actual registry changes; the return
			# is a registry script with deletion instructions for the keys to be deleted
			# in the next build.
			$writer = New-Object Rubberduck.Deployment.Writers.LocalDebugRegistryWriter
			$content = $writer.Write($entries);

			$encoding = New-Object System.Text.ASCIIEncoding;
			[System.IO.File]::WriteAllLines($regFile, $content, $encoding);
		}
	}
}
catch
{
	Write-Host -Foreground Red -Background Black ($_);
	# Cause the build to fail
	throw;
}

# for debugging locally
# Write-Host "Press any key to continue...";
# Read-Host -Prompt "Press Enter to continue";

