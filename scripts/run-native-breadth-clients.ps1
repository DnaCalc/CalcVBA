param(
    [Parameter(Mandatory=$true)][string]$CliPath,
    [Parameter(Mandatory=$true)][string]$OracleSession,
    [Parameter(Mandatory=$true)][string]$OutDir
)
$ErrorActionPreference='Stop'
$PSNativeCommandUseErrorActionPreference=$true
$repo=(Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$OracleSession=(Resolve-Path -LiteralPath $OracleSession).Path
$ready=Get-Content -Raw -LiteralPath (Join-Path $OracleSession 'ready.json') | ConvertFrom-Json
$identity=Get-Content -Raw -LiteralPath (Join-Path $OracleSession 'excel-identity.json') | ConvertFrom-Json
$owned=Get-Process -Id ([int]$ready.pid)
if ($owned.Id -eq 3144 -or $ready.launches -ne 1 -or $identity.ownership -ne 'owned-new-instance' -or
    $owned.Id -ne $identity.pid -or $ready.run_id -cne $identity.run_id -or
    $owned.StartTime.ToUniversalTime() -ne ([datetime]$identity.process_start_utc).ToUniversalTime() -or
    $owned.Path -cne $identity.executable_path) { throw 'Owned Excel identity mismatch' }
$OutDir=[IO.Path]::GetFullPath($OutDir)
if (Test-Path -LiteralPath $OutDir) { throw 'Use a new output directory' }
& $CliPath build (Join-Path $repo 'examples/native/breadth/NativeBreadth.basproj') --target NativeLibrary --out-dir $OutDir
$dll=Join-Path $OutDir 'NativeBreadth.dll'
$client=Get-Content -Raw -LiteralPath (Join-Path $repo 'crates/calcvba-build/tests/support/native_breadth_client.bas')
$client=$client.Replace('__DLL__',$dll.Replace('"','""')) -replace '(?m)^Option Explicit\r?\n',''
$body=Get-Content -Raw -LiteralPath (Join-Path $repo 'examples/native/breadth/Breadth.bas')
$body=$body -replace '(?m)^Attribute .*\r?\n',''
$body=$body -replace '(?m)^Option Explicit\r?\n',''
$source=Join-Path $OutDir 'OracleSelfTest.bas'
"Option Explicit`r`n"+$client+"`r`n"+$body | Set-Content -LiteralPath $source -Encoding utf8NoBOM
$id='native-breadth-'+[Guid]::NewGuid().ToString('N')
$job=Join-Path $OracleSession ('jobs/'+$id+'.json')
$temporary=Join-Path $OutDir 'job.json'
[ordered]@{id=$id;source_path=$source;expected='13 native exports|Dictionary activation|foreign Dictionary|owned Excel and Range|shutdown'} |
    ConvertTo-Json | Set-Content -LiteralPath $temporary -Encoding utf8NoBOM
Move-Item -LiteralPath $temporary -Destination $job
$result=Join-Path $OracleSession ('results/'+$id+'.json')
$deadline=[DateTime]::UtcNow.AddSeconds(180)
while (-not (Test-Path -LiteralPath $result)) {
    if ([DateTime]::UtcNow -gt $deadline) { throw 'Owned Excel breadth probe timed out' }
    Start-Sleep -Milliseconds 200
}
$record=Get-Content -Raw -LiteralPath $result | ConvertFrom-Json
$record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutDir 'excel-result.json') -Encoding utf8NoBOM
if ($record.error -or $record.compile_status -ne 'ok' -or $record.run_status -ne 'ok') { throw ($record | ConvertTo-Json -Compress) }
$record | ConvertTo-Json -Compress
