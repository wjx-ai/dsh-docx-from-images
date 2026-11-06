# Windows.Media.Ocr wrapper. Reads an image, prints one TSV row per text line:
#   y0 <TAB> y1 <TAB> x0 <TAB> x1 <TAB> text
# The bounding box is the union of the line's word rects, in the image's pixel space.
#
# Kept ASCII-only: Windows PowerShell 5.1 reads a .ps1 without a BOM as ANSI.
param(
  [Parameter(Mandatory = $true)][string]$src,
  [Parameter(Mandatory = $true)][string]$out,
  [string]$lang = 'zh-Hans-CN'
)

$ErrorActionPreference = 'Stop'

[void][Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
[void][Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and
  $_.GetParameters().Count -eq 1 -and
  $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]

function Await($op, $type) {
  $task = $asTask.MakeGenericMethod($type).Invoke($null, @($op))
  [void]$task.Wait(-1)
  $task.Result
}

$file    = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($src)) ([Windows.Storage.StorageFile])
$stream  = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap  = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

$language = New-Object Windows.Globalization.Language $lang
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($language)
if ($null -eq $engine) {
  $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
}
if ($null -eq $engine) { throw "no OCR engine available for '$lang'" }

$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

$sb = New-Object System.Text.StringBuilder
foreach ($line in $result.Lines) {
  $words = @($line.Words)
  if ($words.Count -eq 0) { continue }
  $y0 = [double]::MaxValue; $y1 = 0.0; $x0 = [double]::MaxValue; $x1 = 0.0
  foreach ($w in $words) {
    $r = $w.BoundingRect
    if ($r.Y -lt $y0) { $y0 = $r.Y }
    if (($r.Y + $r.Height) -gt $y1) { $y1 = $r.Y + $r.Height }
    if ($r.X -lt $x0) { $x0 = $r.X }
    if (($r.X + $r.Width) -gt $x1) { $x1 = $r.X + $r.Width }
  }
  $text = $line.Text -replace "`t", ' ' -replace "`r", ' ' -replace "`n", ' '
  [void]$sb.AppendLine(('{0:F0}{1}{2:F0}{1}{3:F0}{1}{4:F0}{1}{5}' -f $y0, [char]9, $y1, $x0, $x1, $text))
}

[System.IO.File]::WriteAllText($out, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
Write-Output ('lines=' + @($result.Lines).Count + ' size=' + $decoder.PixelWidth + 'x' + $decoder.PixelHeight)