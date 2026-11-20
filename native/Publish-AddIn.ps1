param([Parameter(Mandatory)][string]$SourcePresentation,
      [Parameter(Mandatory)][string]$Destination)
. "$PSScriptRoot/PowerPoint.ps1"
$target = [IO.Path]::GetFullPath($Destination)
if (Test-Path -LiteralPath $target) { throw "Destination already exists: $target" }
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
$app = New-Object -ComObject PowerPoint.Application
$original = $null
if ($app.Windows.Count) { $original = $app.ActiveWindow }
$source = (Resolve-Path -LiteralPath $SourcePresentation).Path
$deck = $app.Presentations.Open($source, -1, 0, -1)
try {
    $check = $app.Run(([IO.Path]::GetFileName($source) + '!CharCount'), 'x|y|z', '|')
    if ($check -ne 2) { throw 'Unexpected return from the source VBA project.' }
    # ppSaveAsOpenXMLAddin is the native macro-enabled add-in file format.
    $deck.SaveAs($target, 30)
} finally {
    Close-TestDeck $deck
    Restore-OriginalWindow $app $original
}
[ordered]@{source=$source;sourceSha256=(Get-AddInHash $source);output=$target
    sha256=(Get-AddInHash $target);bytes=(Get-Item -LiteralPath $target).Length
    officeVersion=$app.Version;officeBuild=$app.Build;charCount=$check;saveAsFormat=30} |
    ConvertTo-Json
