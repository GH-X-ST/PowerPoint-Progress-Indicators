# PowerPoint Progress Indicators

[简体中文](README.zh-CN.md)

PowerPoint Progress Indicators is a VBA add-in that turns your slide sections into a clickable navigation bar. During a slide show, click a section title to jump to its first slide, or click a slide marker to jump to that slide. This project is an independently maintained edition of Olaf Nöhring's Progress Indicator 1.960. [Upstream project](https://datenbank-projekt.de/projekte/progress-indicator-powerpoint-2010-2013-2016-32-and-64-bit-windows-mac)

## Download

| Edition | Release package | Add-in to install |
| --- | --- | --- |
| English | [Download EN.11.2](https://github.com/GH-X-ST/PowerPoint-Progress-Indicators/releases/download/EN.11.2/ProgressIndicators-EN.11.2.zip) | `ProgressIndicators_v1960_3264b_SB_EN11_2.ppam` |
| Simplified Chinese | [Download CN.11.2](https://github.com/GH-X-ST/PowerPoint-Progress-Indicators/releases/download/CN.11.2/ProgressIndicators-CN.11.2.zip) | `ProgressIndicators_v1960_3264b_SB_CN11_2.ppam` |

Extract the ZIP for the language you want. The `.ppam` is the add-in. The matching `.pptm` is a presentation you can open in PowerPoint to inspect or edit the VBA project. [Earlier releases](https://github.com/GH-X-ST/PowerPoint-Progress-Indicators/releases)

## Install in PowerPoint for Windows

1. In desktop PowerPoint, choose **File > Options > Add-ins**. Under **Manage**, select **PowerPoint Add-ins** and click **Go**.
2. Enable one edition at a time. If an older version or the other language edition is active, clear its checkbox.
3. Click **Add New**, select the extracted `.ppam` for your language, and follow your organization's usual macro approval process.

## Generate progress indicators

1. Open a presentation with sections. On the **Progress Indicators** tab, open **Design Presets** and choose a style.
2. Choose the top, bottom, left, or right edge of the slide. **Equal Sections** gives each nonempty section the same length. **Equal Slides** allocates section lengths in proportion to the number of slides that count toward progress.
3. Click **Update Current Slide** to see the result on the current slide. Click **Generate** to apply it across the configured slide range.

The 40 presets include pill-shaped sections, circular markers, ruler ticks, square segments, and continuous tracks. In **Crimson Layers**, the current section joins the crimson strip in a single shape. EN.11.2 places closer section panels in front of more distant ones, including panels more than three visible sections away from the current section. Section shadows use a 10 pt blur and 62% transparency for a softer edge. Past and upcoming slide markers use 70% transparency. The current marker stays opaque and is the only one that glows. To apply these settings to an existing deck, select **Crimson Layers** again and click **Generate**.

## Build the English edition

Run these commands in PowerShell from the repository root. They create a Python 3.12 environment with `uv`, install the pinned dependencies, and build the English maintenance presentation.

```powershell
uv venv --python 3.12 .venv
uv pip install --python .venv\Scripts\python.exe -r requirements-build.txt
.\.venv\Scripts\python.exe build_package.py --language en --output-dir rebuild --report rebuild/package-validation.json
```

The English VBA and Ribbon XML are generated in `rebuild/en/source/`. The English preset catalog is `rebuild/en/styles.json`. The build uses the shared VBA and preset definitions in `source/` and `styles.json`, the translations in `locales/en/`, and the upstream PowerPoint templates in `original/`.

To produce the release add-in, use desktop PowerPoint to save the generated maintenance presentation as a `.ppam`. The script requires PowerPoint to be able to run macros in that presentation and writes to a new destination. Run it from the same PowerShell window.

```powershell
$name = 'ProgressIndicators_v1960_3264b_SB_EN11_2'
$sourceDeck = (Resolve-Path "rebuild\$name.pptm").Path
$releaseDir = Join-Path (Get-Location) 'rebuild\distribution'
$ppam = Join-Path $releaseDir "$name.ppam"
.\native\Publish-AddIn.ps1 -SourcePresentation $sourceDeck -Destination $ppam
Copy-Item -LiteralPath $sourceDeck -Destination (Join-Path $releaseDir "$name.pptm")
```

## License and credits

The software is distributed under the MIT License and retains the copyright notices for Edward Loper, Adam M. Wilson, and Olaf Nöhring. The upstream templates also include FatCow Farm-Fresh Web Icons, which have separate attribution and licensing terms. [MIT License](LICENSE) | [Third-party notices](THIRD_PARTY_NOTICES.en.md)
