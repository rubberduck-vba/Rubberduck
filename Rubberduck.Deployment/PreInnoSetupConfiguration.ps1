param (
	[Parameter(Mandatory=$true)][string]$workingDir 
)

$source = $workingDir + '\Licenses\license.rtf';
$license = $workingDir + '\InnoSetup\Includes\license.rtf';
$endYear = Get-Date -Format yyyy;
(Get-Content $source).replace('$(YEAR$)', $endYear) | Set-Content $license;
