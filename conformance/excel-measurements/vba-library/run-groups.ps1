# Runs the VBA library probes in Excel VBA, one owned Excel session per group of modules (so a
# compile error in one group loses only that group), on a private desktop, and saves the answers.
#   pwsh -File run-groups.ps1 -RunDir temp\vba-library\x2 [-Group 11] [-Only P_A,P_B]
param(
    [Parameter(Mandatory = $true)][string]$RunDir,
    [int]$Group = 11,
    [string]$Only = ''
)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$RunDir = [IO.Path]::GetFullPath((Join-Path $repo $RunDir))
$desk = Join-Path (Split-Path $repo -Parent) 'DnaCalc\tools\excel-session\on-private-desktop.ps1'
$session = Join-Path $repo 'scripts\run-excel-probe-session.ps1'
& py -3 -I (Join-Path $PSScriptRoot 'run_probes.py') spec $RunDir $Group $Only
$all = (Get-Content -Raw (Join-Path $RunDir 'spec.json') | ConvertFrom-Json).cases
foreach ($c in $all) {
    $d = Join-Path $RunDir $c.id
    [void][IO.Directory]::CreateDirectory($d)
    @{ cases = @($c) } | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $d 'spec.json') -Encoding utf8NoBOM
    & pwsh -NoProfile -File $desk -Command "pwsh -NoProfile -File $session -Directory $d\session -SpecPath $d\spec.json" `
        -Log "$d\desk.log" -GuardOut "$d\guard.json" -SessionRoot "$d\desk" | Out-Null
    $res = Join-Path $d 'session\results.json'
    if (Test-Path $res) {
        & py -3 -I (Join-Path $PSScriptRoot 'run_probes.py') save $res
    } else {
        $dlg = Get-Content (Join-Path $d 'session\guardian-events-1.jsonl') -ErrorAction SilentlyContinue |
            ForEach-Object { $_ | ConvertFrom-Json } | Where-Object event_type -eq 'dialog-observation' | Select-Object -Last 1
        "$($c.id) FAILED: $(($dlg.dialog_text) -join ' ') [$($c.members -join ',')]"
    }
}
