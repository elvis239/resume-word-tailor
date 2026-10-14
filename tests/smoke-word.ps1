param([string]$WorkspaceRoot = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath($WorkspaceRoot)
$tmp=Join-Path $root ('.tmp\word-smoke-'+[guid]::NewGuid().ToString('N'))
if (-not ([IO.Path]::GetFullPath($tmp).StartsWith((Join-Path $root '.tmp')+[IO.Path]::DirectorySeparatorChar))) {
    throw 'Test directory must stay under workspace .tmp.'
}
[void][IO.Directory]::CreateDirectory($tmp)
$editor=Join-Path $root 'skills\resume-word-tailor\scripts\edit-resume.ps1'
$template=Join-Path $tmp 'template.docx'
$app=$null; $doc=$null
function Write-Plan($Changes,[string]$Path) {
    $text=@{replacements=@($Changes)} | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText($Path,$text,(New-Object Text.UTF8Encoding($false)))
}
try {
    $app=New-Object -ComObject Word.Application
    $app.Visible=$false; $app.DisplayAlerts=0
    $doc=$app.Documents.Add()
    $doc.Content.Text="Synthetic Resume`rExample Education`r技术栈：SQL`r背景与目标：`rSynthetic project description."
    $doc.Content.Font.Size=10
    $p=$doc.Paragraphs.Item(3).Range
    $doc.Range($p.Start,$p.Start+4).Font.Bold=-1
    $doc.Paragraphs.Item(4).Range.Font.Bold=-1
    $doc.SaveAs2($template,16)
    $doc.Close(0); $doc=$null
    $app.Quit(); $app=$null
    $hash=(Get-FileHash -LiteralPath $template).Hash
    $inspection=Join-Path $tmp 'inspection.json'
    & $editor -Mode Inspect -TemplatePath $template -InspectionPath $inspection
    $state=Get-Content -LiteralPath $inspection -Raw -Encoding UTF8 | ConvertFrom-Json
    $old=$state.paragraphs[2].text
    $plan=Join-Path $tmp 'plan.json'
    Write-Plan @{index=3; expectedText=$old; text='技术栈：SQL、Python'} $plan
    $output=Join-Path $tmp 'success.docx'
    & $editor -Mode Apply -TemplatePath $template -PlanPath $plan -OutputPath $output
    if (-not (Test-Path -LiteralPath $output)) { throw 'Expected successful output.' }
    if ((Get-FileHash -LiteralPath $template).Hash -cne $hash) { throw 'Source changed.' }
    Write-Output 'PASS: exact replacement, native formatting and source hash verification'
    Write-Plan @{index=3; expectedText='Wrong template text'; text='Replacement'} $plan
    $output=Join-Path $tmp 'mismatch.docx'
    try {
        & $editor -Mode Apply -TemplatePath $template -PlanPath $plan -OutputPath $output
        throw 'Expected mismatch rejection.'
    } catch {
        if ($_.Exception.Message -notlike 'Template text mismatch*') { throw }
    }
    if (Test-Path -LiteralPath $output) { throw 'Mismatch left an output.' }
    Write-Output 'PASS: expected-text mismatch rejects with no output'
    Write-Plan @{index=5; expectedText=$state.paragraphs[4].text; text=('overflow test '*2500)} $plan
    $output=Join-Path $tmp 'overflow.docx'
    try {
        & $editor -Mode Apply -TemplatePath $template -PlanPath $plan -OutputPath $output
        throw 'Expected page-overflow rejection.'
    } catch {
        if ($_.Exception.Message -notlike 'Page overflow*') { throw }
    }
    if (Test-Path -LiteralPath $output) { throw 'Overflow left an output.' }
    if ((Get-FileHash -LiteralPath $template).Hash -cne $hash) { throw 'Source changed after guards.' }
    Write-Output 'PASS: page overflow rejects with no output; source unchanged'
} finally {
    try { if ($null -ne $doc) { $doc.Close(0) } }
    finally {
        try { if ($null -ne $app) { $app.Quit() } }
        finally {
            foreach($file in [IO.Directory]::GetFiles($tmp)) { [IO.File]::Delete($file) }
            [IO.Directory]::Delete($tmp)
        }
    }
}
