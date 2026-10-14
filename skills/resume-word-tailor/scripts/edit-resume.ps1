param(
    [Parameter(Mandatory=$true)][ValidateSet('Inspect','Apply')][string]$Mode,
    [Parameter(Mandatory=$true)][string]$TemplatePath,
    [string]$InspectionPath,
    [string]$PlanPath,
    [string]$OutputPath,
    [string]$PreviewPdf,
    [int]$MaxPages = 0
)
$ErrorActionPreference = 'Stop'

function Write-Utf8([string]$Path, [string]$Text) {
    $parent = Split-Path -Parent $Path
    if ($parent) { [void][System.IO.Directory]::CreateDirectory($parent) }
    [System.IO.File]::WriteAllText($Path, $Text, (New-Object System.Text.UTF8Encoding($false)))
}
function Paragraph-Text($Paragraph) {
    return $Paragraph.Range.Text.TrimEnd([char]13, [char]7)
}
function Snapshot($Document) {
    $Document.Repaginate()
    $paragraphs = @()
    for ($i=1; $i -le $Document.Paragraphs.Count; $i++) {
        $p = $Document.Paragraphs.Item($i)
        $r = $p.Range
        $format = [ordered]@{
            size = $r.Font.Size; eastAsia = $r.Font.NameFarEast; ascii = $r.Font.NameAscii
            before = $p.Format.SpaceBefore; after = $p.Format.SpaceAfter
            spacing = $p.Format.LineSpacing; spacingRule = $p.Format.LineSpacingRule
            alignment = $p.Format.Alignment; left = $p.Format.LeftIndent
            right = $p.Format.RightIndent; firstLine = $p.Format.FirstLineIndent
            listType = $r.ListFormat.ListType; keepNext = $p.Format.KeepWithNext
        }
        $paragraphs += [pscustomobject]@{index=$i; text=(Paragraph-Text $p); format=$format}
    }
    $sections = @()
    for ($i=1; $i -le $Document.Sections.Count; $i++) {
        $s=$Document.Sections.Item($i).PageSetup
        $sections += [ordered]@{width=$s.PageWidth; height=$s.PageHeight; top=$s.TopMargin;
            bottom=$s.BottomMargin; left=$s.LeftMargin; right=$s.RightMargin;
            orientation=$s.Orientation; header=$s.HeaderDistance; footer=$s.FooterDistance}
    }
    return [pscustomobject]@{
        pages=$Document.ComputeStatistics(2); paragraphs=$paragraphs
        sections=$sections; shapes=$Document.Shapes.Count; inlineShapes=$Document.InlineShapes.Count
    }
}
function Stable-Json($Value) { return ConvertTo-Json -InputObject $Value -Depth 12 -Compress }
function Shape-Geometry($Document) {
    $items=@()
    for($i=1; $i -le $Document.Shapes.Count; $i++) {
        $s=$Document.Shapes.Item($i)
        # Word may rename legacy .doc shapes when saving as .docx; compare geometry.
        $items += [ordered]@{type=$s.Type; width=$s.Width; height=$s.Height;
            left=$s.Left; top=$s.Top; rotation=$s.Rotation;
            horizontal=$s.RelativeHorizontalPosition; vertical=$s.RelativeVerticalPosition;
            wrap=$s.WrapFormat.Type}
    }
    return ,$items
}
function Assert-Uniform($Range, [int]$Index) {
    foreach($property in @('Bold','Italic','Underline','Size','Color')) {
        if ($Range.Font.$property -eq 9999999) {
            throw "Mixed $property formatting at paragraph $Index; preserve individual runs with a tailored edit."
        }
    }
}

$template=(Resolve-Path -LiteralPath $TemplatePath).Path
$sourceHash=(Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash
$changes=@()
if ($Mode -eq 'Apply') {
    if (-not $PlanPath -or -not $OutputPath) { throw 'Apply requires PlanPath and OutputPath.' }
    $OutputPath=[System.IO.Path]::GetFullPath($OutputPath)
    if ([System.IO.Path]::GetExtension($OutputPath) -ne '.docx') { throw 'Output must be .docx.' }
    if (Test-Path -LiteralPath $OutputPath) { throw 'Output exists; choose a new path.' }
    if ($PreviewPdf) {
        $PreviewPdf=[System.IO.Path]::GetFullPath($PreviewPdf)
        if ([System.IO.Path]::GetExtension($PreviewPdf) -ne '.pdf') { throw 'Preview must be .pdf.' }
        if (Test-Path -LiteralPath $PreviewPdf) { throw 'Preview exists; choose a new path.' }
    }
    $plan=Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $changes=@($plan.replacements)
    if ($changes.Count -eq 0) { throw 'No replacements in plan.' }
    $seen=@{}
    foreach($c in $changes) {
        foreach($field in @('index','expectedText','text')) {
            if ($null -eq $c.PSObject.Properties[$field]) { throw "Missing $field in replacement." }
        }
        $index=[int]$c.index
        if ($seen.ContainsKey($index)) { throw 'Duplicate paragraph index.' }
        $seen[$index]=$true
        if ([string]$c.text -match "[`r`n`a]") { throw 'Replacement must not contain paragraph or cell markers.' }
    }
}

$app=$null; $doc=$null; $created=$false; $previewCreated=$false; $ok=$false
try {
    $app=New-Object -ComObject Word.Application
    $app.Visible=$false
    $app.DisplayAlerts=0
    $doc=$app.Documents.Open($template,$false,$true)
    $before=Snapshot $doc
    $shapeBefore=Shape-Geometry $doc
    if ($Mode -eq 'Inspect') {
        $json=ConvertTo-Json -InputObject $before -Depth 12
        if ($InspectionPath) { Write-Utf8 ([System.IO.Path]::GetFullPath($InspectionPath)) $json }
        else { Write-Output $json }
        $ok=$true
    } else {
        if ($MaxPages -le 0) { $MaxPages=$before.pages }
        foreach($c in $changes) {
            $i=[int]$c.index
            if ($i -lt 1 -or $i -gt $before.paragraphs.Count) { throw 'Paragraph index outside template.' }
            if ((Paragraph-Text $doc.Paragraphs.Item($i)) -cne [string]$c.expectedText) {
                throw "Template text mismatch at paragraph $i; inspect again."
            }
            $pr=$doc.Paragraphs.Item($i).Range
            if ($pr.Information(12)) { throw "Table paragraph $i needs a tailored edit; generic replacement is unsupported." }
            $r=$pr.Duplicate; $r.End=$r.End-1
            $old=$r.Text; $new=[string]$c.text
            $colon=$old.IndexOf('：'); $newColon=$new.IndexOf('：')
            $hasLabel=$false
            if ($colon -ge 0 -and $colon -lt $old.Length-1 -and $newColon -ge 0) {
                $label=$doc.Range($r.Start,$r.Start+$colon+1)
                $hasLabel=($label.Font.Bold -eq -1)
            }
            $c | Add-Member -NotePropertyName hasLabel -NotePropertyValue $hasLabel
            if($hasLabel) {
                Assert-Uniform ($doc.Range($r.Start,$r.Start+$colon+1)) $i
                $body=$doc.Range($r.Start+$colon+1,$r.End)
                $tailBold=$null
                if($body.Font.Bold -eq 9999999 -and $body.Text.Length -gt 1) {
                    $last=$body.Text.Substring($body.Text.Length-1)
                    if($last -match '^[。；，！？.!?;]$' -and $new.EndsWith($last)) {
                        $core=$doc.Range($body.Start,$body.End-1)
                        Assert-Uniform $core $i
                        $tailBold=$doc.Range($body.End-1,$body.End).Font.Bold
                        $body=$core
                    }
                }
                Assert-Uniform $body $i
                $c | Add-Member -NotePropertyName bodyBold -NotePropertyValue $body.Font.Bold
                $c | Add-Member -NotePropertyName tailBold -NotePropertyValue $tailBold
            } else { Assert-Uniform $r $i }
        }
        foreach($c in ($changes | Sort-Object { [int]$_.index } -Descending)) {
            $r=$doc.Paragraphs.Item([int]$c.index).Range.Duplicate
            $r.End=$r.End-1
            $old=$r.Text; $new=[string]$c.text
            $colon=$old.IndexOf('：'); $newColon=$new.IndexOf('：')
            if ($c.hasLabel) {
                $doc.Range($r.Start+$colon+1,$r.End).Text=$new.Substring($newColon+1)
                $doc.Range($r.Start,$r.Start+$colon+1).Text=$new.Substring(0,$newColon+1)
                $updated=$doc.Paragraphs.Item([int]$c.index).Range
                $doc.Range($updated.Start,$updated.Start+$newColon+1).Font.Bold=-1
                $doc.Range($updated.Start+$newColon+1,$updated.End-1).Font.Bold=$c.bodyBold
                if($null -ne $c.tailBold) { $doc.Range($updated.End-2,$updated.End-1).Font.Bold=$c.tailBold }
            } else { $r.Text=$new }
        }
        $after=Snapshot $doc
        if ((Stable-Json (Shape-Geometry $doc)) -cne (Stable-Json $shapeBefore)) { throw 'Shape geometry changed.' }
        if ($after.pages -gt $MaxPages) { throw "Page overflow: $($after.pages) > $MaxPages. Do not shorten finalized text or alter formatting silently." }
        if ($before.paragraphs.Count -ne $after.paragraphs.Count -or $before.shapes -ne $after.shapes -or
            $before.inlineShapes -ne $after.inlineShapes -or (Stable-Json $before.sections) -cne (Stable-Json $after.sections)) {
            throw 'Document structure or page setup changed.'
        }
        for($i=0; $i -lt $before.paragraphs.Count; $i++) {
            if ((Stable-Json $before.paragraphs[$i].format) -cne (Stable-Json $after.paragraphs[$i].format)) {
                throw "Formatting changed at paragraph $($i+1); inspect source runs."
            }
            if (-not $seen.ContainsKey($i+1) -and $before.paragraphs[$i].text -cne $after.paragraphs[$i].text) {
                throw "Untargeted text changed at paragraph $($i+1)."
            }
        }
        [void][System.IO.Directory]::CreateDirectory((Split-Path -Parent $OutputPath))
        $created=$true
        $doc.SaveAs2($OutputPath,16)
        $doc.Close(0); $doc=$null
        $doc=$app.Documents.Open($OutputPath,$false,$true)
        foreach($c in $changes) {
            if ((Paragraph-Text $doc.Paragraphs.Item([int]$c.index)) -cne [string]$c.text) {
                throw "Saved text mismatch at paragraph $($c.index)."
            }
        }
        $saved=Snapshot $doc
        if ((Stable-Json (Shape-Geometry $doc)) -cne (Stable-Json $shapeBefore)) { throw 'Saved shape geometry changed.' }
        if ((Stable-Json $saved) -cne (Stable-Json $after)) { throw 'Saved document changed text, layout, or formatting.' }
        if ($PreviewPdf) {
            [void][System.IO.Directory]::CreateDirectory((Split-Path -Parent $PreviewPdf))
            $previewCreated=$true
            $doc.ExportAsFixedFormat($PreviewPdf,17)
        }
        if ((Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash -cne $sourceHash) { throw 'Source file hash changed.' }
        $ok=$true
        Write-Output (ConvertTo-Json -InputObject ([ordered]@{output=$OutputPath; pages=$saved.pages;
            replacements=$changes.Count; paragraphs=$saved.paragraphs.Count; shapes=$saved.shapes;
            sourceUnchanged=$true; exactTextVerified=$true; formattingVerified=$true}) -Compress)
    }
} finally {
    try { if ($null -ne $doc) { $doc.Close(0) } }
    finally {
        try { if ($null -ne $app) { $app.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } }
        finally {
            if (-not $ok) {
                if ($created -and [System.IO.File]::Exists($OutputPath)) { [System.IO.File]::Delete($OutputPath) }
                if ($previewCreated -and [System.IO.File]::Exists($PreviewPdf)) { [System.IO.File]::Delete($PreviewPdf) }
            }
        }
    }
}
