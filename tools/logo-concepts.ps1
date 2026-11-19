# Renders ILLXL logo concepts side by side at 256 / 64 / 32 / 16 px -> tools\concepts.png
Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

function Tile([Drawing.Graphics]$g, [float]$s, $c1, $c2) {
    $r = $s * 0.22
    $p = New-Object Drawing.Drawing2D.GraphicsPath
    $p.AddArc(0, 0, 2 * $r, 2 * $r, 180, 90); $p.AddArc($s - 2 * $r, 0, 2 * $r, 2 * $r, 270, 90)
    $p.AddArc($s - 2 * $r, $s - 2 * $r, 2 * $r, 2 * $r, 0, 90); $p.AddArc(0, $s - 2 * $r, 2 * $r, 2 * $r, 90, 90)
    $p.CloseFigure()
    $b = New-Object Drawing.Drawing2D.LinearGradientBrush ([Drawing.PointF]::new(0, 0)), ([Drawing.PointF]::new($s, $s)), $c1, $c2
    $g.FillPath($b, $p)
}
function Txt([Drawing.Graphics]$g, $text, $font, [float]$size, $color, [float]$cx, [float]$cy, $style = 'Bold') {
    $f = New-Object Drawing.Font $font, $size, ([Drawing.FontStyle]$style), ([Drawing.GraphicsUnit]::Pixel)
    $sf = New-Object Drawing.StringFormat; $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
    $g.DrawString($text, $f, (New-Object Drawing.SolidBrush $color), ([Drawing.RectangleF]::new($cx - 500, $cy - 500, 1000, 1000)), $sf)
}
$navy = [Drawing.Color]::FromArgb(15, 23, 42); $blue = [Drawing.Color]::FromArgb(37, 99, 235)
$cyan = [Drawing.Color]::FromArgb(56, 189, 248); $white = [Drawing.Color]::White
$green = [Drawing.Color]::FromArgb(52, 211, 153)

# Each concept draws itself into a square of size s
$concepts = [ordered]@{
    'A  XL wordmark' = {
        param($g, $s)
        Tile $g $s $navy $blue
        Txt $g 'ILL' 'Bahnschrift' ($s * 0.20) $cyan ($s / 2) ($s * 0.27) 'Bold'
        Txt $g 'XL' 'Bahnschrift' ($s * 0.50) $white ($s / 2) ($s * 0.62) 'Bold'
    }
    'B  bars + X' = {
        param($g, $s)
        Tile $g $s $navy $navy
        $w = $s * 0.09; $b = New-Object Drawing.SolidBrush $white; $a = New-Object Drawing.SolidBrush $cyan
        $g.FillRectangle($b, $s * 0.16, $s * 0.22, $w, $s * 0.56)                        # I
        $g.FillRectangle($b, $s * 0.31, $s * 0.22, $w, $s * 0.56); $g.FillRectangle($b, $s * 0.31, $s * 0.69, $s * 0.15, $w)   # L
        $pen = New-Object Drawing.Pen $cyan, ($s * 0.11); $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
        $g.DrawLine($pen, $s * 0.56, $s * 0.27, $s * 0.82, $s * 0.73); $g.DrawLine($pen, $s * 0.82, $s * 0.27, $s * 0.56, $s * 0.73)
    }
    'C  big X' = {
        param($g, $s)
        Tile $g $s $blue $navy
        $pen = New-Object Drawing.Pen $white, ($s * 0.15); $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
        $g.DrawLine($pen, $s * 0.30, $s * 0.28, $s * 0.70, $s * 0.72)
        $pen2 = New-Object Drawing.Pen $cyan, ($s * 0.15); $pen2.StartCap = 'Round'; $pen2.EndCap = 'Round'
        $g.DrawLine($pen2, $s * 0.70, $s * 0.28, $s * 0.30, $s * 0.72)
    }
    'D  ILL/XL stacked' = {
        param($g, $s)
        Tile $g $s $navy $blue
        Txt $g 'ILL' 'Bahnschrift' ($s * 0.36) $white ($s / 2) ($s * 0.32) 'Bold'
        Txt $g 'XL' 'Bahnschrift' ($s * 0.36) $cyan ($s / 2) ($s * 0.68) 'Bold'
    }
}

$sizes = 256, 64, 32, 16
$rowH = 290; $W = 30 + ($sizes | Measure-Object -Sum).Sum + 40 * $sizes.Count + 220
$sheet = New-Object Drawing.Bitmap ([int]$W), ([int]($rowH * $concepts.Count + 20))
$SG = [Drawing.Graphics]::FromImage($sheet); $SG.Clear([Drawing.Color]::FromArgb(245, 246, 248))
$y = 20
foreach ($name in $concepts.Keys) {
    Txt $SG $name 'Segoe UI' 18 $navy 110 ($y + 128) 'Bold'
    $x = 230
    foreach ($s in $sizes) {
        $bmp = New-Object Drawing.Bitmap ([int]$s), ([int]$s)
        $g = [Drawing.Graphics]::FromImage($bmp); $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'
        & $concepts[$name] $g ([float]$s)
        $SG.DrawImage($bmp, $x, $y + (256 - $s) / 2, $s, $s); $x += $s + 40
    }
    $y += $rowH
}
$sheet.Save("$PSScriptRoot\concepts.png")
"saved $PSScriptRoot\concepts.png"
