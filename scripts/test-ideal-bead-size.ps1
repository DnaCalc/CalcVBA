$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib-ideal-program-validation.ps1')

foreach ($size in @('small', 'medium', 'large')) {
    Assert-IdealBeadSize -Issue ([pscustomobject]@{ labels = @("size-$size") }) -Owner 'relative-size'
}
Assert-IdealBeadSize -Issue ([pscustomobject]@{ labels = @(); estimated_minutes = 120 }) -Owner 'historical'
$negativeCases = @(
    [pscustomobject]@{ labels = @() },
    [pscustomobject]@{ labels = @(); estimated_minutes = 0 },
    [pscustomobject]@{ labels = @(); estimated_minutes = 481 },
    [pscustomobject]@{ labels = @('size-huge'); estimated_minutes = 120 },
    [pscustomobject]@{ labels = @('size-small', 'size-large'); estimated_minutes = 120 }
)
foreach ($case in $negativeCases) {
    $rejected = $false
    try { Assert-IdealBeadSize -Issue $case -Owner 'invalid-size' }
    catch {
        if ($_.Exception.Message -notmatch '^validate-workset-rollout: invalid-size ') { throw }
        $rejected = $true
    }
    if (-not $rejected) { throw 'Invalid or missing relative size unexpectedly passed' }
}
Write-Host 'test-ideal-bead-size: ok (4 positive, 5 negative cases)'
