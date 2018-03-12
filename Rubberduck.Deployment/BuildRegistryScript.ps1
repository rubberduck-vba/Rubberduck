param (
	[Parameter(Mandatory=$true)][string]$builderAssemblyPath,
	[Parameter(Mandatory=$true)][string]$netToolsDir,
	[Parameter(Mandatory=$true)][string]$wixToolsDir,
	[Parameter(Mandatory=$true)][string]$sourceDir,
	[Parameter(Mandatory=$true)][string]$targetDir,
	[Parameter(Mandatory=$true)][string]$filesToExtract
)

Set-StrictMode -Version latest
$ErrorActionPreference = "Stop";

try
{
	$separator = "|"
	$option = [System.StringSplitOptions]::RemoveEmptyEntries
	$files = $filesToExtract.Split($separator, $option);

	foreach($file in $files)
	{
		$sourceDll = $sourceDir + $file;
		$targetDll = $targetDir + $file;
		$sourceTlb = $sourceDir + ($file -replace ".dll", ".tlb");
		$targetTlb = $targetDir + ($file -replace ".dll", ".tlb");
		$dllXml = $targetDll + ".xml"
		$tlbXml = $targetTlb + ".xml"

		$cmd = "{0}tlbexp.exe" -f $netToolsDir;
		# & "C:\GitHub\Rubberduck\Rubberduck\Rubberduck.Deployment\echoargs.exe" ""$sourceDll"" /win32 /out:""$sourceTlb"";
		& $cmd ""$sourceDll"" /win32 /out:""$sourceTlb"";

		$cmd = "{0}heat.exe" -f $wixToolsDir;
		& $cmd file ""$sourceDll"" -out ""$dllXml"";
		& $cmd file ""$sourceTlb"" -out ""$tlbXml"";

		$bitness = 0;
	
		[System.Reflection.Assembly]::LoadFrom($builderAssemblyPath);
		$builder = New-Object Rubberduck.Deployment.RegistryEntryBuilder
	
		$out = $builder.Parse($tlbXml, $dllXml, $bitness);

		$out | Format-Table | Out-String |% {Write-Host $_};
	}
}
catch
{
	Write-Host -Foreground Red -Background Black ($_);
}

# for debugging locally
# Write-Host "Press any key to continue...";
# [void][System.Console]::ReadKey($true);