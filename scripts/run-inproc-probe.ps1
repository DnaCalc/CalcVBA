# Runs a probe module in one owned Excel three ways (real VBA, CalcVBA on VM3, CalcVBA on the
# JIT) through the VBA loader of tools/calcvba-excel, or with -Xll through its XLL loader.
# Builds the DLL, copies it to a fresh path (Excel keeps a loaded DLL locked), fills in the
# loader and comparison modules, and runs them with run-excel-probe-session.ps1 on a private
# desktop. Results: <Directory>\results.json. With -PerModule each module is its own case in its
# own workbook (Entry in each), so a line VBA refuses to compile hides no other line.
#
#   pwsh -File scripts/run-inproc-probe.ps1 -Directory temp/inproc/run1 -ModulePath probe.bas -Entry RunAll
param(
    [Parameter(Mandatory=$true)][string]$Directory,
    [Parameter(Mandatory=$true)][string[]]$ModulePath,
    [Parameter(Mandatory=$true)][string]$Entry,
    [string]$Module,
    [switch]$Xll,
    [switch]$PerModule,
    [switch]$SkipBuild
)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$Directory=[IO.Path]::GetFullPath($Directory)
if (Test-Path -LiteralPath $Directory) { throw 'Use a new run directory' }
[void][IO.Directory]::CreateDirectory($Directory)
if (-not $SkipBuild) {
    & cargo build -p calcvba-excel --manifest-path (Join-Path $repo 'Cargo.toml')
    if ($LASTEXITCODE) { throw 'cargo build failed' }
}
$dll=Join-Path $Directory 'calcvba_excel.dll'
Copy-Item -LiteralPath (Join-Path $repo 'target\debug\calcvba_excel.dll') -Destination $dll
$modules=@()
foreach ($path in $ModulePath) {
    $full=[IO.Path]::GetFullPath($path)
    $name=[IO.Path]::GetFileNameWithoutExtension($full)
    $kind=if ($full.EndsWith('.cls')) { 2 } else { 1 }
    $modules+=[ordered]@{name=$name;kind=$kind;path=$full}
}
if (-not $Module) { $Module=$modules[0].name }
$loaders=Join-Path $repo 'tools\calcvba-excel\loaders'
$loader=(Get-Content -Raw -LiteralPath (Join-Path $loaders 'CalcVBALoader.bas')).Replace('<DLL>',$dll)
$compareTemplate=if ($Xll) { 'CalcVBAXllCompare.bas' } else { 'CalcVBACompare.bas' }
$compareText=Get-Content -Raw -LiteralPath (Join-Path $loaders $compareTemplate)
$loaderPath=Join-Path $Directory 'CalcVBALoader.bas'; Set-Content -LiteralPath $loaderPath -Value $loader -Encoding utf8NoBOM
# One case: its modules, the loader, and a comparison module running $moduleName.$Entry.
function New-Case($id, $caseModules, $moduleName) {
    $comparePath=Join-Path $Directory "CalcVBACompare-$id.bas"
    Set-Content -LiteralPath $comparePath -Value $compareText.Replace('<DLL>',$dll).Replace('<MODULE>',$moduleName).Replace('<ENTRY>',$Entry) -Encoding utf8NoBOM
    $all=@($caseModules)+@([ordered]@{name='CalcVBALoader';kind=1;path=$loaderPath},[ordered]@{name='CalcVBACompare';kind=1;path=$comparePath})
    @{id=$id;modules=$all;entry='CalcVBACompare.CompareAll'}
}
$cases=if ($PerModule) { @($modules | ForEach-Object { New-Case $_.name @($_) $_.name }) } else { @(New-Case 'inproc' $modules $Module) }
$spec=Join-Path $Directory 'spec.json'
@{cases=$cases} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $spec -Encoding utf8NoBOM
$desktop=Join-Path (Split-Path $repo -Parent) 'DnaCalc\tools\excel-session\on-private-desktop.ps1'
& pwsh -NoProfile -File $desktop -Command "pwsh -NoProfile -File $(Join-Path $PSScriptRoot 'run-excel-probe-session.ps1') -Directory $(Join-Path $Directory 'session') -SpecPath $spec" `
    -Log (Join-Path $Directory 'desk.log') -GuardOut (Join-Path $Directory 'guard.json') -SessionRoot (Join-Path $Directory 'desk') | Out-Null
$results=Join-Path $Directory 'session\results.json'
if (-not (Test-Path -LiteralPath $results)) { Get-Content -LiteralPath (Join-Path $Directory 'desk.log') -Tail 30; throw 'no results' }
Copy-Item -LiteralPath $results -Destination (Join-Path $Directory 'results.json')
foreach ($r in (Get-Content -Raw -LiteralPath $results | ConvertFrom-Json)) {
    "$($r.id) compile=$($r.compile_status) error=$($r.run_error)"
    if ($r.compile_dialogs) { $r.compile_dialogs.dialog_text; $r.compile_dialogs.expanded_line }
    if ($r.run_dialogs) { $r.run_dialogs.dialog_text }
    $r.run_value
}
