[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function New-RoundedPath {
    param(
        [System.Drawing.RectangleF]$Rectangle,
        [float]$Radius
    )

    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $diameter = $Radius * 2
    $path.AddArc($Rectangle.Left, $Rectangle.Top, $diameter, $diameter, 180, 90)
    $path.AddArc($Rectangle.Right - $diameter, $Rectangle.Top, $diameter, $diameter, 270, 90)
    $path.AddArc($Rectangle.Right - $diameter, $Rectangle.Bottom - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($Rectangle.Left, $Rectangle.Bottom - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-IconPng {
    param([int]$Size)

    $scale = $Size / 256.0
    $bitmap = [System.Drawing.Bitmap]::new($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $background = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(40, 122, 104))
    $paper = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(249, 247, 241))
    $blue = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(45, 132, 170))
    $darkPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(23, 61, 54), [Math]::Max(1.0, 10 * $scale))
    $whitePen = [System.Drawing.Pen]::new([System.Drawing.Color]::White, [Math]::Max(1.0, 10 * $scale))
    $tealPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(40, 122, 104), [Math]::Max(1.0, 10 * $scale))
    $stream = [System.IO.MemoryStream]::new()

    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::Transparent)

        $backgroundPath = New-RoundedPath -Rectangle ([System.Drawing.RectangleF]::new(0, 0, $Size, $Size)) -Radius ([float](58 * $scale))
        try { $graphics.FillPath($background, $backgroundPath) } finally { $backgroundPath.Dispose() }

        $page = [System.Drawing.RectangleF]::new([float](61 * $scale), [float](40 * $scale), [float](134 * $scale), [float](174 * $scale))
        $graphics.FillRectangle($paper, $page)
        $graphics.DrawRectangle($darkPen, $page.X, $page.Y, $page.Width, $page.Height)
        $tealPen.StartCap = $tealPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
        $graphics.DrawLine($tealPen, 86 * $scale, 103 * $scale, 152 * $scale, 103 * $scale)
        $graphics.DrawLine($tealPen, 86 * $scale, 133 * $scale, 138 * $scale, 133 * $scale)
        $graphics.DrawLine($tealPen, 86 * $scale, 163 * $scale, 125 * $scale, 163 * $scale)

        $graphics.FillEllipse($blue, 123 * $scale, 123 * $scale, 96 * $scale, 96 * $scale)
        $graphics.DrawEllipse($darkPen, 123 * $scale, 123 * $scale, 96 * $scale, 96 * $scale)
        $whitePen.StartCap = $whitePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
        $graphics.DrawLine($whitePen, 171 * $scale, 143 * $scale, 171 * $scale, 173 * $scale)
        $graphics.DrawLine($whitePen, 171 * $scale, 173 * $scale, 192 * $scale, 186 * $scale)

        $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Output -NoEnumerate $stream.ToArray()
    }
    finally {
        $stream.Dispose()
        $darkPen.Dispose()
        $whitePen.Dispose()
        $tealPen.Dispose()
        $background.Dispose()
        $paper.Dispose()
        $blue.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

$sizes = @(16, 24, 32, 48, 64, 128, 256)
$images = [System.Collections.Generic.List[byte[]]]::new()
foreach ($size in $sizes) {
    $images.Add((New-IconPng -Size $size))
}
$parent = Split-Path -Parent $OutputPath
if ($parent) {
    [System.IO.Directory]::CreateDirectory($parent) | Out-Null
}

$file = [System.IO.File]::Create($OutputPath)
$writer = [System.IO.BinaryWriter]::new($file)
try {
    $writer.Write([UInt16]0)
    $writer.Write([UInt16]1)
    $writer.Write([UInt16]$images.Count)
    [UInt32]$offset = 6 + (16 * $images.Count)
    for ($index = 0; $index -lt $images.Count; $index++) {
        $size = $sizes[$index]
        $writer.Write([byte]$(if ($size -eq 256) { 0 } else { $size }))
        $writer.Write([byte]$(if ($size -eq 256) { 0 } else { $size }))
        $writer.Write([byte]0)
        $writer.Write([byte]0)
        $writer.Write([UInt16]1)
        $writer.Write([UInt16]32)
        $writer.Write([UInt32]$images[$index].Length)
        $writer.Write($offset)
        $offset += [UInt32]$images[$index].Length
    }
    foreach ($image in $images) {
        $writer.Write($image)
    }
}
finally {
    $writer.Dispose()
    $file.Dispose()
}

Write-Host "Created $OutputPath"
