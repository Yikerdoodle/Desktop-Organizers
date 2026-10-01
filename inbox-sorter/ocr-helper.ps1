# Reads text inside pictures (Windows' built-in OCR) + photo details (camera, date taken).
# Used by inbox_sorter.py. Input: a UTF-8 JSON file with a list of image paths.
# Output: one JSON line per image: {path, text, taken, make, model, width, height, error}
# Everything happens locally with Windows' own OCR engine - nothing is uploaded anywhere.
param([Parameter(Mandatory)][string]$ListFile)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.FileProperties.ImageProperties, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]

# WinRT calls are async; this turns them into normal blocking calls
$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
function Await($op, [Type]$type) {
    $task = $asTaskGeneric.MakeGenericMethod($type).Invoke($null, @($op))
    [void]$task.Wait(20000)
    $task.Result
}

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
$max = [Windows.Media.Ocr.OcrEngine]::MaxImageDimension
$paths = Get-Content -LiteralPath $ListFile -Raw -Encoding UTF8 | ConvertFrom-Json

foreach ($path in $paths) {
    $out = [ordered]@{ path = $path; text = ''; taken = ''; make = ''; model = ''; width = 0; height = 0; error = '' }
    try {
        $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile])
        try {
            $p = Await ($file.Properties.GetImagePropertiesAsync()) ([Windows.Storage.FileProperties.ImageProperties])
            if ($p.DateTaken.Year -gt 1601) { $out.taken = $p.DateTaken.ToString('yyyy-MM-dd') }
            $out.make = "$($p.CameraManufacturer)".Trim(); $out.model = "$($p.CameraModel)".Trim()
        } catch { }
        $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        try {
            $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
            $w = [int]$dec.OrientedPixelWidth; $h = [int]$dec.OrientedPixelHeight
            $out.width = $w; $out.height = $h
            $t = New-Object Windows.Graphics.Imaging.BitmapTransform
            $big = [Math]::Max($w, $h)
            if ($big -gt $max) {  # OCR has a size limit: shrink big pictures first
                $scale = $max / $big
                $t.ScaledWidth = [uint32][Math]::Floor($dec.PixelWidth * $scale)
                $t.ScaledHeight = [uint32][Math]::Floor($dec.PixelHeight * $scale)
            }
            $bmp = Await ($dec.GetSoftwareBitmapAsync([Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
                    [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied, $t,
                    [Windows.Graphics.Imaging.ExifOrientationMode]::RespectExifOrientation,
                    [Windows.Graphics.Imaging.ColorManagementMode]::DoNotColorManage)) ([Windows.Graphics.Imaging.SoftwareBitmap])
            if ($engine -and [Math]::Min($w, $h) -ge 40) {
                $r = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
                $out.text = (($r.Lines | ForEach-Object { $_.Text }) -join "`n")
                if ($out.text.Length -gt 4000) { $out.text = $out.text.Substring(0, 4000) }
            }
        } finally { $stream.Dispose() }
    } catch {
        $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }; $out.error = $e.Message
    }
    [Console]::Out.WriteLine(($out | ConvertTo-Json -Compress))
    [Console]::Out.Flush()
}
