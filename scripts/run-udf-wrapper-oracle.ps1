param(
    [Parameter(Mandatory=$true)][string]$OracleSession,
    [Parameter(Mandatory=$true)][string]$Directory
)
$ErrorActionPreference='Stop'
$PSNativeCommandUseErrorActionPreference=$true
$root=Split-Path $PSScriptRoot
$Directory=[IO.Path]::GetFullPath($Directory)
if (Test-Path -LiteralPath $Directory) { throw 'Use a new UDF comparison directory' }
$fixture=Join-Path $root 'tools/calcvba-udf/tests/projects'
$generated=Join-Path $Directory 'generated'
Push-Location $root
try {
    cargo run -p calcvba-udf -- --project (Join-Path $fixture 'Functions.basproj') --out $generated --prefix Native_ --paramarray-slots 4
    $xll=Join-Path $generated 'FunctionsUdf.xll'
    $source=(Get-Content -Raw -LiteralPath (Join-Path $fixture 'Oracle.bas')).Replace('@XLL_PATH@',$xll)
    $oracle=Join-Path $Directory 'Oracle.bas'
    $source | Set-Content -LiteralPath $oracle -Encoding utf8NoBOM
    $modules=Join-Path $Directory 'modules.json'
    $runningObject=Join-Path $Directory 'owned-excel.object'
    $backendSource=(Get-Content -Raw -LiteralPath (Join-Path $fixture 'BackendProbe.bas')).Replace('@RUNNING_OBJECT_PATH@',$runningObject)
    $backendSource | Set-Content -LiteralPath (Join-Path $Directory 'BackendProbe.bas') -Encoding utf8NoBOM
    $backendProject=Join-Path $Directory 'Backend.basproj'
    $backendXml=[xml](Get-Content -Raw -LiteralPath (Join-Path $fixture 'Functions.basproj'))
    $backendXml.Project.ItemGroup.Module.SetAttribute('Include',(Join-Path $fixture 'Functions.bas'))
    $probe=$backendXml.CreateElement('Module'); $probe.SetAttribute('Include','BackendProbe.bas'); [void]$backendXml.Project.ItemGroup.AppendChild($probe)
    $backendXml.Save($backendProject)
    @(
        @{name='Functions';kind=1;source_path=(Join-Path $fixture 'Functions.bas')},
        @{name='OracleSelfTest';kind=1;source_path=$oracle}
    ) | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $modules -Encoding utf8NoBOM
    & (Join-Path $PSScriptRoot 'native-export-oracle-project.ps1') -OracleSession $OracleSession -Directory (Join-Path $Directory 'office') -ModulesFile $modules -ReferencePaths (Join-Path ([Environment]::GetFolderPath('System')) 'scrrun.dll') -RunnerProject $backendProject -RunningObjectPath $runningObject
    $result=Get-Content -Raw -LiteralPath (Join-Path $Directory 'office/result.json') | ConvertFrom-Json
    if ($result.compile_status -ne 'ok' -or $result.run_status -ne 'ok') { throw "UDF comparison failed: $($result.error)" }
    foreach ($backend in @('vm3','jit')) {
        $values=Get-Content -Raw -LiteralPath (Join-Path $Directory ('office/'+$backend+'.txt'))
        if (-not $values.Contains('owned-range=5:Range:2')) { throw "$backend did not return the owned Range comparison" }
    }
}
finally { Pop-Location }
