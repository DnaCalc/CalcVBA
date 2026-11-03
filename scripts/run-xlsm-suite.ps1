param(
    [Parameter(Mandatory=$true)][string]$RuntimeTestExecutable,
    [string[]]$Cases=@(),
    [string]$RunId=('xlsm_suite_{0:yyyyMMddTHHmmssZ}' -f [DateTime]::UtcNow)
)
$ErrorActionPreference='Stop'
$repo=(Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$root=Join-Path $repo ('temp/calcvba15/suites/'+$RunId)
New-Item -ItemType Directory -Path $root -ErrorAction Stop | Out-Null
# Keep the exact binary immutable while later Cargo builds run.
$engine=Join-Path $root 'xlsm_projects.exe'
Copy-Item -LiteralPath (Resolve-Path $RuntimeTestExecutable).Path -Destination $engine
if ($Cases.Count -eq 0) {
    $Cases=@('chains','events','mixed','documents','arrays','bounds','application')
    $Cases+=@((Get-Content -Raw (Join-Path $repo 'crates/calcvba-host/tests/support/xlsm/cases.json') | ConvertFrom-Json).id)
}
$results=@()
foreach ($case in $Cases) {
    for ($attempt=1; $attempt -le 3; $attempt++) {
        $caseRun=$RunId+'-'+$case+'-'+$attempt
        $log=Join-Path $root ($case+'-'+$attempt+'.log')
        try {
            & (Join-Path $PSScriptRoot 'run-xlsm-project-oracle.ps1') -CaseId $case -RuntimeTestExecutable $engine -RunId $caseRun *> $log
            $results+=,[ordered]@{case=$case;run_id=$caseRun;attempt=$attempt;passed=$true;infrastructure_retry=$false;error=$null}
            Write-Output ($case+': passed')
            break
        } catch {
            $infrastructure=$_.Exception.Message -like '*first selected case failed before durable ownership*'
            $retry=$infrastructure -and $attempt -lt 3
            $results+=,[ordered]@{case=$case;run_id=$caseRun;attempt=$attempt;passed=$false;infrastructure_retry=$retry;error=$_.Exception.Message}
            $_ | Out-String | Add-Content -LiteralPath $log
            Write-Output ($case+': failed; '+$_.Exception.Message)
            if (-not $retry) { break }
        }
        ConvertTo-Json -InputObject $results -Depth 8 | Set-Content -LiteralPath (Join-Path $root 'results.json') -Encoding utf8NoBOM
    }
    ConvertTo-Json -InputObject $results -Depth 8 | Set-Content -LiteralPath (Join-Path $root 'results.json') -Encoding utf8NoBOM
}
if (@($results | Where-Object { -not $_.passed -and -not $_.infrastructure_retry }).Count -gt 0) { throw "Saved XLSM suite has failures; $root/results.json" }
