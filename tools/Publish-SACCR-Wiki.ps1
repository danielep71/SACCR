param(
    [string]$RepoRoot = "",
    [string]$WikiPath = ""
)

$ErrorActionPreference = "Stop"

function Invoke-Git {
    param([string]$WorkingDirectory, [Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)
    & git -C $WorkingDirectory @Arguments
    if ($LASTEXITCODE -ne 0) { throw "git $($Arguments -join ' ') failed in $WorkingDirectory" }
}

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}
$RepoRoot = [IO.Path]::GetFullPath($RepoRoot)
if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot ".git"))) { throw "Repository not found: $RepoRoot" }

if ([string]::IsNullOrWhiteSpace($WikiPath)) {
    $WikiPath = Join-Path (Split-Path -Parent $RepoRoot) ((Split-Path -Leaf $RepoRoot) + ".wiki")
}
$WikiPath = [IO.Path]::GetFullPath($WikiPath)

$origin = (& git -C $RepoRoot remote get-url origin).Trim()
if ($LASTEXITCODE -ne 0) { throw "Cannot read origin." }
$match = [regex]::Match($origin, "github\.com[:/](?<slug>[^/]+/[^/]+?)(?:\.git)?$")
if (-not $match.Success) { throw "Unsupported GitHub origin: $origin" }
$wikiUrl = "https://github.com/$($match.Groups['slug'].Value).wiki.git"

if ((& git -C $RepoRoot status --porcelain)) {
    throw "Wiki publication requires a clean committed source tree."
}
$sourceSha = (& git -C $RepoRoot rev-parse HEAD).Trim()
$token = [Guid]::NewGuid().ToString("N")
$export = Join-Path ([IO.Path]::GetTempPath()) "saccr-wiki-export-$token"
$readback = Join-Path ([IO.Path]::GetTempPath()) "saccr-wiki-readback-$token"

try {
    Write-Host "Source: $RepoRoot"
    Write-Host "Commit: $sourceSha"
    Write-Host "Wiki:   $wikiUrl"
    Write-Host "Local:  $WikiPath"
    Write-Host ""

    & python (Join-Path $RepoRoot "tools\check_wiki.py") --root $RepoRoot --export-dir $export
    if ($LASTEXITCODE -ne 0) { throw "Wiki source/export validation failed." }

    if (-not (Test-Path -LiteralPath (Join-Path $WikiPath ".git"))) {
        if (Test-Path -LiteralPath $WikiPath) {
            if (@(Get-ChildItem -LiteralPath $WikiPath -Force).Count -ne 0) {
                throw "Wiki path exists but is not an empty Git checkout: $WikiPath"
            }
            Remove-Item -LiteralPath $WikiPath -Force
        }
        & git clone $wikiUrl $WikiPath
        if ($LASTEXITCODE -ne 0) { throw "Wiki clone failed." }
    } else {
        Invoke-Git $WikiPath remote set-url origin $wikiUrl
        Invoke-Git $WikiPath pull --ff-only
    }

    $new = Get-Content (Join-Path $export "Wiki-Source.json") -Raw | ConvertFrom-Json
    $newManaged = @($new.pages.PSObject.Properties.Name) + @("Wiki-Source.json")
    $oldManaged = @()
    $oldManifest = Join-Path $WikiPath "Wiki-Source.json"
    if (Test-Path $oldManifest) {
        $old = Get-Content $oldManifest -Raw | ConvertFrom-Json
        if ($null -ne $old.pages) { $oldManaged = @($old.pages.PSObject.Properties.Name) + @("Wiki-Source.json") }
    }

    $allowed = @($newManaged + $oldManaged | Sort-Object -Unique)
    $unmanaged = @(Get-ChildItem $WikiPath -File -Force |
        Where-Object { $_.Name -notin $allowed -and $_.Name -notin @(".gitignore", ".gitattributes") } |
        Select-Object -ExpandProperty Name)
    if ($unmanaged.Count -gt 0) {
        throw "Unmanaged Wiki files must be migrated into docs/wiki first: $($unmanaged -join ', ')"
    }

    foreach ($name in $oldManaged) {
        if ($name -notin $newManaged) {
            $path = Join-Path $WikiPath $name
            if (Test-Path $path) { Remove-Item $path -Force }
        }
    }
    foreach ($name in $newManaged) {
        Copy-Item (Join-Path $export $name) (Join-Path $WikiPath $name) -Force
    }

    Invoke-Git $WikiPath add -A
    & git -C $WikiPath diff --cached --quiet
    $diffCode = $LASTEXITCODE
    if ($diffCode -eq 1) {
        Invoke-Git $WikiPath diff --cached --check
        Invoke-Git $WikiPath commit -m "Publish wiki from $sourceSha"
        Invoke-Git $WikiPath push origin HEAD
        Write-Host "Wiki published."
    } elseif ($diffCode -eq 0) {
        Write-Host "Wiki already current; no commit needed."
    } else {
        throw "Unable to inspect staged Wiki changes."
    }

    & git clone --quiet $wikiUrl $readback
    if ($LASTEXITCODE -ne 0) { throw "Fresh Wiki read-back clone failed." }
    & python (Join-Path $RepoRoot "tools\check_wiki.py") --root $RepoRoot --published-dir $readback
    if ($LASTEXITCODE -ne 0) { throw "Published Wiki does not match source." }

    Write-Host ""
    Write-Host "SUCCESS: Wiki matches source commit $sourceSha."
}
finally {
    foreach ($path in @($export, $readback)) {
        if (Test-Path $path) { Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
