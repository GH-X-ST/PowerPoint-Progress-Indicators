$ErrorActionPreference = 'Stop'
$Project = Split-Path -Parent $PSScriptRoot
$Utf8 = [Text.UTF8Encoding]::new($false)

function New-TestDeck($App, [string[]] $Titles, [int[]] $Starts, [int] $Count,
                      [int] $Width = 960, [int] $Height = 540, [switch] $Dark) {
    $deck = $App.Presentations.Add(-1)
    $deck.PageSetup.SlideWidth = $Width
    $deck.PageSetup.SlideHeight = $Height
    for ($i = 1; $i -le $Count; $i++) {
        $slide = $deck.Slides.Add($i, 12)
        $body = $slide.Shapes.AddTextbox(1, $Width * 0.25, $Height * 0.44,
                                          $Width * 0.5, 80)
        $body.Name = 'PI_BODY_SENTINEL'
        $body.TextFrame.TextRange.Text = "Progress Indicators  /  $i"
        $body.TextFrame.TextRange.Font.Name = 'Arial'
        $body.TextFrame.TextRange.Font.Size = 24
        if ($Dark) {
            $slide.FollowMasterBackground = 0
            $slide.Background.Fill.Solid()
            $slide.Background.Fill.ForeColor.RGB = 0x30251B
            $body.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
        }
    }
    for ($i = 0; $i -lt $Starts.Count; $i++) {
        [void]$deck.SectionProperties.AddBeforeSlide($Starts[$i], $Titles[$i])
    }
    return $deck
}

function Close-TestDeck($Deck) {
    $Deck.Saved = -1
    $Deck.Close()
}

function Write-CaseLog([string] $Path, $Record) {
    $line = ConvertTo-Json -InputObject $Record -Depth 8 -Compress
    [IO.File]::AppendAllText($Path, $line + [Environment]::NewLine, $Utf8)
}

function Export-NativeSlides($Deck, [int[]] $Slides, [string] $Folder) {
    [void][IO.Directory]::CreateDirectory($Folder)
    $height = [int](1440 * $Deck.PageSetup.SlideHeight / $Deck.PageSetup.SlideWidth)
    foreach ($i in $Slides) {
        $Deck.Slides.Item($i).Export((Join-Path $Folder "$i.png"), 'PNG', 1440, $height)
    }
}

function Open-TestAddIn($App, [string] $Path) {
    $resolved = (Resolve-Path -LiteralPath $Path).Path
    foreach ($existing in $App.AddIns) {
        if ($existing.FullName -eq $resolved -and $existing.Loaded) {
            return @{ Item = $existing; Added = $false; LoadedHere = $false }
        }
        if ($existing.FullName -eq $resolved) {
            $existing.Loaded = -1
            return @{ Item = $existing; Added = $false; LoadedHere = $true }
        }
    }
    $item = $App.AddIns.Add($resolved)
    $item.Loaded = -1
    return @{ Item = $item; Added = $true; LoadedHere = $true }
}

function Close-TestAddIn($App, $Handle) {
    if ($Handle.LoadedHere) { $Handle.Item.Loaded = 0 }
    if ($Handle.Added) { $App.AddIns.Remove($Handle.Item.Name) }
}


function Get-AddInHash([string] $Path) {
    $stream = [IO.File]::Open((Resolve-Path -LiteralPath $Path).Path,
        [IO.FileMode]::Open, [IO.FileAccess]::Read,
        ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try { return (Get-FileHash -InputStream $stream).Hash }
    finally { $stream.Dispose() }
}

function Restore-OriginalWindow($App, $Window) {
    foreach ($remaining in $App.Windows) {
        if ($remaining -eq $Window) {
            $remaining.Activate()
            return
        }
    }
}
