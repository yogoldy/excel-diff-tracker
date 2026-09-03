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
    $white = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(247, 251, 249))
    $whitePen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(247, 251, 249), [Math]::Max(1.0, 11 * $scale))
    $stream = [System.IO.MemoryStream]::new()

    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::Transparent)

        $backgroundPath = New-RoundedPath -Rectangle ([System.Drawing.RectangleF]::new(0, 0, $Size, $Size)) -Radius ([float](58 * $scale))
        try { $graphics.FillPath($background, $backgroundPath) } finally { $backgroundPath.Dispose() }

        $whitePen.StartCap = $whitePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
        $whitePen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
        $grid = [System.Drawing.RectangleF]::new([float](42 * $scale), [float](47 * $scale), [float](92 * $scale), [float](162 * $scale))
        $gridPath = New-RoundedPath -Rectangle $grid -Radius ([float](9 * $scale))
        try { $graphics.DrawPath($whitePen, $gridPath) } finally { $gridPath.Dispose() }
        $graphics.DrawLine($whitePen, 42 * $scale, 91 * $scale, 134 * $scale, 91 * $scale)
        $graphics.DrawLine($whitePen, 42 * $scale, 135 * $scale, 134 * $scale, 135 * $scale)
        $graphics.DrawLine($whitePen, 42 * $scale, 177 * $scale, 134 * $scale, 177 * $scale)
        $graphics.DrawLine($whitePen, 88 * $scale, 47 * $scale, 88 * $scale, 209 * $scale)
        $graphics.DrawLine($whitePen, 111 * $scale, 135 * $scale, 139 * $scale, 135 * $scale)
        $graphics.DrawLine($whitePen, 139 * $scale, 135 * $scale, 190 * $scale, 92 * $scale)
        $graphics.DrawLine($whitePen, 139 * $scale, 135 * $scale, 216 * $scale, 135 * $scale)
        $graphics.DrawLine($whitePen, 139 * $scale, 135 * $scale, 190 * $scale, 178 * $scale)
        foreach ($point in @(@(139,135,10), @(216,92,12), @(216,135,12), @(216,178,12))) {
            $graphics.FillEllipse($white, ($point[0]-$point[2])*$scale, ($point[1]-$point[2])*$scale, (2*$point[2])*$scale, (2*$point[2])*$scale)
        }

        $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Output -NoEnumerate $stream.ToArray()
    }
    finally {
        $stream.Dispose()
        $whitePen.Dispose()
        $background.Dispose()
        $white.Dispose()
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
