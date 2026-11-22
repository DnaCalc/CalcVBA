param(
    [Parameter(Mandatory)][string]$RunDirectory,
    [Parameter(Mandatory)][string]$PrivateDesktopScript,
    [string[]]$Only,
    [int]$Size
)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
$desktopScript = (Resolve-Path -LiteralPath $PrivateDesktopScript).Path
$run = [IO.Path]::GetFullPath($RunDirectory)
if (Test-Path -LiteralPath $run) { throw 'Use a new run directory.' }
if ($Size -lt 0 -or ($Size -and $Only.Count -ne 1)) { throw 'A size override requires one workload.' }
$selected = @(Import-Csv (Join-Path $PSScriptRoot cases.tsv) -Delimiter "`t" | Where-Object { !$Only -or $_.workload -in $Only })
if (!$selected.Count -or ($Only -and $selected.Count -ne $Only.Count)) { throw 'Unknown or repeated workload.' }
[void][IO.Directory]::CreateDirectory($run)
$file = (Join-Path $run lines.txt).Replace('"', '""')
$driver = Join-Path $run BenchExcel.bas
$body = [Collections.Generic.List[string]]::new()
$body.Add((Get-Content -Raw (Join-Path $PSScriptRoot BenchExcel.bas)))
$body.Add('Public Function RunAll() As String')
foreach ($case in $selected) {
    $n = if ($Size) { $Size } else { [int]$case.size }
    $body.Add(('    RunAll = RunAll & Bench("{0}", {1}, "{2}")' -f $case.workload, $n, $file))
}
$body.Add('End Function')
$body | Set-Content -LiteralPath $driver -Encoding utf8NoBOM
$specPath = Join-Path $run spec.json
@{ cases = @(@{ id = 'VbaBench'; entry = 'BenchExcel.RunAll'; modules = @(
    @{ name = 'B_Work'; kind = 1; path = (Join-Path $PSScriptRoot B_Work.bas) }
    @{ name = 'B_Acc'; kind = 2; path = (Join-Path $PSScriptRoot B_Acc.cls) }
    @{ name = 'B_Sheet'; kind = 1; path = (Join-Path $PSScriptRoot B_Sheet.bas) }
    @{ name = 'ProbeFmt'; kind = 1; path = (Join-Path $repo 'conformance/excel-measurements/vba-library/probes/ProbeFmt.bas') }
    @{ name = 'BenchExcel'; kind = 1; path = $driver }
) }) } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $specPath -Encoding utf8NoBOM
$session = Join-Path $run session
$script = Join-Path $repo scripts/run-excel-probe-session.ps1
$command = "pwsh -NoProfile -File '$($script.Replace("'", "''"))' -Directory '$($session.Replace("'", "''"))' -SpecPath '$($specPath.Replace("'", "''"))'"
$desktop = 'CalcVbaBench-' + [Guid]::NewGuid().ToString('N').Substring(0,12)
$desktopArgs = @('-NoProfile', '-File', "`"$desktopScript`"", '-Command', "`"$command`"", '-Log', "`"$run\desktop.log`"",
    '-GuardOut', "`"$run\guard.json`"", '-SessionRoot', "`"$session`"", '-DesktopName', $desktop)
$wrapper = Start-Process -FilePath (Join-Path $PSHOME pwsh.exe) -ArgumentList $desktopArgs -WindowStyle Hidden -PassThru
if (!$wrapper.WaitForExit(895000)) {
    $launch = Join-Path $session launch.json
    if (Test-Path -LiteralPath $launch) {
        $identity = Get-Content -Raw -LiteralPath $launch | ConvertFrom-Json
        $owned = Get-Process -Id $identity.pid -ErrorAction SilentlyContinue
        if ($owned -and $owned.Path -eq $identity.path -and
            $owned.StartTime.ToUniversalTime() -eq ([DateTimeOffset]$identity.start_utc).UtcDateTime) {
            $owned.Kill()
        }
    }
    if (!$wrapper.WaitForExit(5000)) { $wrapper.Kill() }
    throw "Excel benchmark timed out. See $run."
}
if ($wrapper.ExitCode -ne 0) { throw "Private Excel session failed. See $run."
}
$result = Get-Content -Raw -LiteralPath (Join-Path $session results.json) | ConvertFrom-Json
if ($result.compile_status -ne 'ok' -or $result.run_error -or !$result.run_value -or $result.run_dialogs.Count) {
    throw "Excel benchmark did not complete. See $session\results.json."
}
$header = "workload`tsize`tchecksum_bits`tmedian_ms`trun1_ms`trun2_ms`trun3_ms`trun4_ms`trun5_ms"
$header, $result.run_value.TrimEnd() | Set-Content -LiteralPath (Join-Path $run excel.tsv) -Encoding utf8NoBOM
