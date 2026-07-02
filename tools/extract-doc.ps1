param(
  [Parameter(Mandatory=$true)][string]$InputDoc,
  [Parameter(Mandatory=$true)][string]$OutputText
)
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $resolvedInput = (Resolve-Path $InputDoc).Path
  $doc = $word.Documents.Open($resolvedInput, $false, $true)
  $text = $doc.Content.Text -replace "`v", "" -replace "`f", "`r"
  [System.IO.File]::WriteAllText(
    [System.IO.Path]::GetFullPath($OutputText),
    $text,
    [System.Text.UTF8Encoding]::new($false)
  )
} finally {
  if ($doc) { $doc.Close($false) }
  if ($word) { $word.Quit() }
  if ($doc) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
  if ($word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
