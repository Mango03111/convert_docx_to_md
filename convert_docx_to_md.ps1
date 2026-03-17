param(
    [string]$DocxPath = "",
    [string]$MarkdownPath = "",
    [string]$ImageDir = ""
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-RunStyle {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$Node,
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$Ns
    )

    $isBold = $false
    $isItalic = $false

    $runProps = $Node.SelectSingleNode("./w:rPr", $Ns)
    if ($runProps) {
        $boldNode = $runProps.SelectSingleNode("./w:b", $Ns)
        $italicNode = $runProps.SelectSingleNode("./w:i", $Ns)
        if ($boldNode -and $boldNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main") -ne "false") {
            $isBold = $true
        }
        elseif ($boldNode -and $boldNode.Attributes.Count -eq 0) {
            $isBold = $true
        }

        if ($italicNode -and $italicNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main") -ne "false") {
            $isItalic = $true
        }
        elseif ($italicNode -and $italicNode.Attributes.Count -eq 0) {
            $isItalic = $true
        }
    }

    return @{
        Bold = $isBold
        Italic = $isItalic
    }
}

function Format-InlineStyle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [Parameter(Mandatory = $true)]
        [bool]$Bold,
        [Parameter(Mandatory = $true)]
        [bool]$Italic
    )

    if (-not $Text) {
        return ""
    }

    if ($Text -match '^\[\[IMAGE:[^\]]+\]\]$') {
        return $Text
    }

    $trimmed = $Text.Trim()
    if (-not $trimmed) {
        return $Text
    }

    $leadingLength = $Text.Length - $Text.TrimStart().Length
    $trailingLength = $Text.Length - $Text.TrimEnd().Length
    $leading = $Text.Substring(0, $leadingLength)
    $core = $Text.Substring($leadingLength, $Text.Length - $leadingLength - $trailingLength)
    $trailing = $Text.Substring($Text.Length - $trailingLength)

    if ($Bold -and $Italic) {
        return $leading + "***" + $core + "***" + $trailing
    }
    if ($Bold) {
        return $leading + "**" + $core + "**" + $trailing
    }
    if ($Italic) {
        return $leading + "*" + $core + "*" + $trailing
    }

    return $Text
}

function Get-NodeMarkdown {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$Node,
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$Ns
    )

    $builder = New-Object System.Text.StringBuilder
    foreach ($child in $Node.ChildNodes) {
        switch ($child.LocalName) {
            "t" {
                [void]$builder.Append($child.InnerText)
            }
            "tab" {
                [void]$builder.Append("    ")
            }
            "br" {
                [void]$builder.Append("`n")
            }
            "cr" {
                [void]$builder.Append("`n")
            }
            "drawing" {
                $blips = $child.SelectNodes(".//a:blip", $Ns)
                foreach ($blip in $blips) {
                    $embed = $blip.GetAttribute("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                    if ($embed) {
                        [void]$builder.Append("[[IMAGE:$embed]]")
                    }
                }
            }
            "r" {
                $style = Get-RunStyle -Node $child -Ns $Ns
                $runText = Get-NodeMarkdown -Node $child -Ns $Ns
                [void]$builder.Append((Format-InlineStyle -Text $runText -Bold $style.Bold -Italic $style.Italic))
            }
            "hyperlink" {
                [void]$builder.Append((Get-NodeMarkdown -Node $child -Ns $Ns))
            }
            default {
                if ($child.HasChildNodes) {
                    [void]$builder.Append((Get-NodeMarkdown -Node $child -Ns $Ns))
                }
            }
        }
    }

    return $builder.ToString()
}

function Convert-UrlText {
    param([string]$Text)

    $pattern = '(https?://[^\s\)]+)'
    return [System.Text.RegularExpressions.Regex]::Replace(
        $Text,
        $pattern,
        { param($m) "[{0}]({0})" -f $m.Groups[1].Value }
    )
}

function Remove-MarkdownEmphasis {
    param([string]$Text)

    if (-not $Text) {
        return ""
    }

    return ($Text -replace '(\*\*\*|\*\*|\*)', '')
}

function Get-ParagraphStyle {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$Paragraph,
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$Ns
    )

    $styleNode = $Paragraph.SelectSingleNode("./w:pPr/w:pStyle", $Ns)
    if (-not $styleNode) {
        return ""
    }

    if ($styleNode.Attributes["w:val"]) {
        return $styleNode.Attributes["w:val"].Value
    }

    foreach ($attr in $styleNode.Attributes) {
        if ($attr.LocalName -eq "val") {
            return $attr.Value
        }
    }

    return ""
}

function Get-NumberingMap {
    param([System.IO.Compression.ZipArchive]$Zip)

    $numberingMap = @{}
    $numberingEntry = $Zip.GetEntry("word/numbering.xml")
    if (-not $numberingEntry) {
        return $numberingMap
    }

    $reader = New-Object System.IO.StreamReader($numberingEntry.Open())
    $xmlText = $reader.ReadToEnd()
    $reader.Close()
    [xml]$numberingXml = $xmlText

    $ns = New-Object System.Xml.XmlNamespaceManager($numberingXml.NameTable)
    $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

    $abstractMap = @{}
    foreach ($abstract in $numberingXml.SelectNodes("//w:abstractNum", $ns)) {
        $abstractId = $abstract.GetAttribute("abstractNumId", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        if (-not $abstractId) {
            continue
        }

        $levelMap = @{}
        foreach ($lvl in $abstract.SelectNodes("./w:lvl", $ns)) {
            $ilvl = $lvl.GetAttribute("ilvl", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            $numFmt = $lvl.SelectSingleNode("./w:numFmt", $ns)
            $fmt = "bullet"
            if ($numFmt) {
                $fmtValue = $numFmt.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                if ($fmtValue) {
                    $fmt = $fmtValue
                }
            }
            $levelMap[$ilvl] = $fmt
        }
        $abstractMap[$abstractId] = $levelMap
    }

    foreach ($num in $numberingXml.SelectNodes("//w:num", $ns)) {
        $numId = $num.GetAttribute("numId", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        $abstractRef = $num.SelectSingleNode("./w:abstractNumId", $ns)
        if (-not $numId -or -not $abstractRef) {
            continue
        }
        $abstractId = $abstractRef.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        if ($abstractMap.ContainsKey($abstractId)) {
            $numberingMap[$numId] = $abstractMap[$abstractId]
        }
    }

    return $numberingMap
}

function Get-ListInfo {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$Paragraph,
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$Ns,
        [Parameter(Mandatory = $true)]
        [hashtable]$NumberingMap
    )

    $numPr = $Paragraph.SelectSingleNode("./w:pPr/w:numPr", $Ns)
    if (-not $numPr) {
        $style = Get-ParagraphStyle -Paragraph $Paragraph -Ns $Ns
        if ($style -eq "ListParagraph") {
            return @{
                IsList = $true
                Level = 0
                Marker = "-"
            }
        }

        return @{
            IsList = $false
            Level = 0
            Marker = ""
        }
    }

    $ilvlNode = $numPr.SelectSingleNode("./w:ilvl", $Ns)
    $numIdNode = $numPr.SelectSingleNode("./w:numId", $Ns)
    $level = 0
    $numId = ""
    if ($ilvlNode) {
        $ilvl = $ilvlNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        if ($ilvl -ne "") {
            $level = [int]$ilvl
        }
    }
    if ($numIdNode) {
        $numId = $numIdNode.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    }

    $fmt = "bullet"
    if ($numId -and $NumberingMap.ContainsKey($numId)) {
        $levelMap = $NumberingMap[$numId]
        $levelKey = [string]$level
        if ($levelMap.ContainsKey($levelKey)) {
            $fmt = $levelMap[$levelKey]
        }
    }

    $marker = "-"
    if ($fmt -and $fmt -notin @("bullet", "none")) {
        $marker = "1."
    }

    return @{
        IsList = $true
        Level = $level
        Marker = $marker
    }
}

function Convert-InlineImageMarkers {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [Parameter(Mandatory = $true)]
        [hashtable]$ImageMap
    )

    return [System.Text.RegularExpressions.Regex]::Replace($Text, '\[\[IMAGE:([^\]]+)\]\]', {
        param($m)
        $rid = $m.Groups[1].Value
        if ($ImageMap.ContainsKey($rid)) {
            $relativePath = $ImageMap[$rid] -replace "\\", "/"
            return "![]($relativePath)"
        }
        return ""
    })
}

function Convert-TableCellToMarkdown {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$Cell,
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$Ns,
        [Parameter(Mandatory = $true)]
        [hashtable]$ImageMap
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($paragraph in $Cell.SelectNodes("./w:p", $Ns)) {
        $raw = Get-NodeMarkdown -Node $paragraph -Ns $Ns
        $text = Convert-InlineImageMarkers -Text $raw.Trim() -ImageMap $ImageMap
        if ($text) {
            $parts.Add($text)
        }
    }

    $cellText = ($parts -join "<br>")
    $cellText = $cellText -replace '\|', '\|'
    return $cellText
}

function Convert-TableToMarkdown {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$Table,
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$Ns,
        [Parameter(Mandatory = $true)]
        [hashtable]$ImageMap
    )

    $rows = @($Table.SelectNodes("./w:tr", $Ns))
    if ($rows.Count -eq 0) {
        return @()
    }

    $renderedRows = New-Object System.Collections.Generic.List[object]
    $maxCols = 0
    foreach ($row in $rows) {
        $cells = @($row.SelectNodes("./w:tc", $Ns))
        $values = New-Object System.Collections.Generic.List[string]
        foreach ($cell in $cells) {
            $values.Add((Convert-TableCellToMarkdown -Cell $cell -Ns $Ns -ImageMap $ImageMap))
        }
        if ($values.Count -gt $maxCols) {
            $maxCols = $values.Count
        }
        $renderedRows.Add($values.ToArray())
    }

    if ($maxCols -eq 0) {
        return @()
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $header = @($renderedRows[0])
    while ($header.Count -lt $maxCols) {
        $header += ""
    }
    $lines.Add("| " + ($header -join " | ") + " |")
    $lines.Add("| " + ((1..$maxCols | ForEach-Object { "---" }) -join " | ") + " |")

    for ($i = 1; $i -lt $renderedRows.Count; $i++) {
        $rowValues = @($renderedRows[$i])
        while ($rowValues.Count -lt $maxCols) {
            $rowValues += ""
        }
        $lines.Add("| " + ($rowValues -join " | ") + " |")
    }

    return ,$lines.ToArray()
}

function Convert-ParagraphToMarkdown {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$Paragraph,
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$Ns,
        [Parameter(Mandatory = $true)]
        [hashtable]$ImageMap,
        [Parameter(Mandatory = $true)]
        [hashtable]$NumberingMap,
        [Parameter(Mandatory = $true)]
        [bool]$IsFirstParagraph
    )

    $raw = Get-NodeMarkdown -Node $Paragraph -Ns $Ns
    $raw = $raw -replace "\r", ""
    $raw = $raw.Trim()
    $style = Get-ParagraphStyle -Paragraph $Paragraph -Ns $Ns

    $lines = New-Object System.Collections.Generic.List[string]

    $imagePattern = '\[\[IMAGE:([^\]]+)\]\]'
    $textWithoutImages = ([System.Text.RegularExpressions.Regex]::Replace($raw, $imagePattern, "")).Trim()
    if ($textWithoutImages) {
        $textWithoutImages = Convert-UrlText $textWithoutImages
        $listInfo = Get-ListInfo -Paragraph $Paragraph -Ns $Ns -NumberingMap $NumberingMap
        if ($IsFirstParagraph) {
            $lines.Add("# $(Remove-MarkdownEmphasis $textWithoutImages)")
        }
        elseif ($listInfo.IsList) {
            $indent = "  " * $listInfo.Level
            $lines.Add(("{0}{1} {2}" -f $indent, $listInfo.Marker, $textWithoutImages))
        }
        elseif ($style -match '^Heading[1-6]$') {
            $level = [int]($style.Substring(7))
            if ($level -lt 1 -or $level -gt 6) { $level = 2 }
            $lines.Add(("{0} {1}" -f ("#" * $level), (Remove-MarkdownEmphasis $textWithoutImages)))
        }
        elseif (
            $textWithoutImages.Length -le 20 -and (
                $textWithoutImages.EndsWith(":") -or
                [int][char]$textWithoutImages[$textWithoutImages.Length - 1] -eq 0xFF1A
            )
        ) {
            $lines.Add(("## {0}" -f (Remove-MarkdownEmphasis $textWithoutImages)))
        }
        else {
            $lines.Add($textWithoutImages)
        }
    }

    $imageMatches = [System.Text.RegularExpressions.Regex]::Matches($raw, $imagePattern)
    foreach ($match in $imageMatches) {
        $rid = $match.Groups[1].Value
        if ($ImageMap.ContainsKey($rid)) {
            $relativePath = $ImageMap[$rid] -replace "\\", "/"
            $lines.Add("![]($relativePath)")
        }
    }

    return ,$lines.ToArray()
}

$workspace = $PWD.Path

if (-not $DocxPath) {
    $docxFiles = @(Get-ChildItem -LiteralPath $workspace -Filter *.docx)
    if ($docxFiles.Count -ne 1) {
        throw "Expected exactly one .docx file in the working directory when -DocxPath is omitted."
    }
    $DocxPath = $docxFiles[0].Name
}

if (-not $MarkdownPath) {
    $MarkdownPath = ([System.IO.Path]::GetFileNameWithoutExtension($DocxPath) + ".md")
}

if (-not $ImageDir) {
    $ImageDir = ([System.IO.Path]::GetFileNameWithoutExtension($DocxPath) + "_images")
}

$docxFullPath = Join-Path $workspace $DocxPath
$mdFullPath = Join-Path $workspace $MarkdownPath
$imageFullDir = Join-Path $workspace $ImageDir

if (-not (Test-Path -LiteralPath $docxFullPath)) {
    throw "Source docx not found: $docxFullPath"
}

if (Test-Path -LiteralPath $imageFullDir) {
    Remove-Item -LiteralPath $imageFullDir -Recurse -Force
}
New-Item -ItemType Directory -Path $imageFullDir | Out-Null

$zip = [System.IO.Compression.ZipFile]::OpenRead($docxFullPath)
try {
    $relsEntry = $zip.GetEntry("word/_rels/document.xml.rels")
    $docEntry = $zip.GetEntry("word/document.xml")
    if (-not $relsEntry -or -not $docEntry) {
        throw "Invalid docx: missing document.xml or relationships."
    }

    $relsReader = New-Object System.IO.StreamReader($relsEntry.Open())
    $relsText = $relsReader.ReadToEnd()
    $relsReader.Close()
    [xml]$relsXml = $relsText

    $docReader = New-Object System.IO.StreamReader($docEntry.Open())
    $docText = $docReader.ReadToEnd()
    $docReader.Close()
    [xml]$docXml = $docText

    $ns = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
    $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    $ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    $ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $relNs = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
    $relNs.AddNamespace("pr", "http://schemas.openxmlformats.org/package/2006/relationships")
    $numberingMap = Get-NumberingMap -Zip $zip

    $imageMap = @{}
    foreach ($rel in $relsXml.SelectNodes("//*[local-name()='Relationship']")) {
        $target = $rel.GetAttribute("Target")
        $id = $rel.GetAttribute("Id")
        if ($target -like "media/*") {
            $entry = $zip.GetEntry(("word/" + $target.Replace("\", "/")))
            if ($entry) {
                $fileName = [System.IO.Path]::GetFileName($target)
                $destPath = Join-Path $imageFullDir $fileName
                $inStream = $entry.Open()
                $outStream = [System.IO.File]::Open($destPath, [System.IO.FileMode]::Create)
                try {
                    $inStream.CopyTo($outStream)
                }
                finally {
                    $outStream.Dispose()
                    $inStream.Dispose()
                }
                $imageMap[$id] = Join-Path $ImageDir $fileName
            }
        }
    }

    $body = $docXml.SelectSingleNode("//w:body", $ns)
    $output = New-Object System.Collections.Generic.List[string]
    $seenFirstParagraph = $false
    foreach ($child in $body.ChildNodes) {
        $lines = @()
        if ($child.LocalName -eq "p") {
            $lines = Convert-ParagraphToMarkdown -Paragraph $child -Ns $ns -ImageMap $imageMap -NumberingMap $numberingMap -IsFirstParagraph:(-not $seenFirstParagraph)
        }
        elseif ($child.LocalName -eq "tbl") {
            $lines = Convert-TableToMarkdown -Table $child -Ns $ns -ImageMap $imageMap
        }

        if (@($lines).Count -gt 0) {
            foreach ($line in $lines) {
                if ($line -ne "") {
                    $output.Add($line)
                }
            }
            $output.Add("")
            if (-not $seenFirstParagraph) {
                $seenFirstParagraph = $true
            }
        }
    }

    while ($output.Count -gt 0 -and [string]::IsNullOrWhiteSpace($output[$output.Count - 1])) {
        $output.RemoveAt($output.Count - 1)
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($mdFullPath, $output, $utf8NoBom)

    "Markdown: $mdFullPath"
    "Images: $imageFullDir"
    "Image count: $($imageMap.Count)"
}
finally {
    $zip.Dispose()
}
