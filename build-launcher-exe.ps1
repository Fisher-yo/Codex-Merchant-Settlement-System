$ErrorActionPreference = 'Stop'

$projectRoot = $PSScriptRoot
$sourcePath = Join-Path $projectRoot 'tools\SettlementConsoleLauncher.cs'
$exeName = (-join ([char[]](0x5468, 0x7ed3, 0x7b97, 0x64cd, 0x4f5c, 0x53f0))) + '.exe'
$portableName = (-join ([char[]](0x5468, 0x7ed3, 0x7b97, 0x64cd, 0x4f5c, 0x53f0, 0x005f, 0x4fbf, 0x643a, 0x7248)))
$rootExePath = Join-Path $projectRoot $exeName
$portableDir = Join-Path (Join-Path $projectRoot 'dist') $portableName
$portableExePath = Join-Path $portableDir $exeName
$compilerPath = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'

if (-not (Test-Path -LiteralPath $compilerPath)) {
  $compilerPath = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
}

if (-not (Test-Path -LiteralPath $compilerPath)) {
  throw 'Cannot find .NET Framework C# compiler csc.exe.'
}

if (-not (Test-Path -LiteralPath $sourcePath)) {
  throw "Cannot find launcher source: $sourcePath"
}

& $compilerPath /nologo /target:winexe /optimize+ /platform:anycpu /reference:System.Windows.Forms.dll /out:$rootExePath $sourcePath

if (Test-Path -LiteralPath $portableDir) {
  Copy-Item -LiteralPath $rootExePath -Destination $portableExePath -Force
}

Write-Host "Created: $rootExePath"
if (Test-Path -LiteralPath $portableExePath) {
  Write-Host "Copied: $portableExePath"
}
