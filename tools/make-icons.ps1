# ILLXL brand + ribbon icons
#   pwsh -File tools\make-icons.ps1
# Draws every icon on a 128px canvas and saves downscaled PNGs to assets\icons\
#   <name>.png      32 px  (normal ribbon buttons)
#   <name>_L.png    64 px  (large ribbon buttons)
# plus assets\logo-512.png / logo-256.png for the README and releases.
# Logo = concept D: "ILL" over "XL" on a navy-to-blue tile.
Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot '..\assets\icons'
New-Item -ItemType Directory -Force $out | Out-Null

#---- palette ---------------------------------------------------------------------
function C([string]$hex) { [Drawing.ColorTranslator]::FromHtml($hex) }
$INK = C '#8A94A6'                     # mid-gray: readable on light AND dark Office themes
$BLUE = C '#2563EB'; $NAVY = C '#0F172A'; $CYAN = C '#38BDF8'
$GREEN = C '#16A34A'; $RED = C '#DC2626'; $AMBER = C '#F59E0B'; $WHITE = C '#FFFFFF'

#---- drawing helpers (all on the 128px canvas) -------------------------------------
function RoundRect([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $p = New-Object Drawing.Drawing2D.GraphicsPath
    $p.AddArc($x, $y, 2 * $r, 2 * $r, 180, 90); $p.AddArc($x + $w - 2 * $r, $y, 2 * $r, 2 * $r, 270, 90)
    $p.AddArc($x + $w - 2 * $r, $y + $h - 2 * $r, 2 * $r, 2 * $r, 0, 90); $p.AddArc($x, $y + $h - 2 * $r, 2 * $r, 2 * $r, 90, 90)
    $p.CloseFigure(); $p
}
function Fill($cv, $color, [float]$x, [float]$y, [float]$w, [float]$h, [float]$r = 10) {
    $cv.FillPath((New-Object Drawing.SolidBrush $color), (RoundRect $x $y $w $h $r))
}
function Stroke($cv, $color, [float]$width, [float]$x, [float]$y, [float]$w, [float]$h, [float]$r = 10, [switch]$Dotted) {
    $pen = New-Object Drawing.Pen $color, $width
    if ($Dotted) { $pen.DashStyle = 'Dot' }
    $cv.DrawPath($pen, (RoundRect $x $y $w $h $r))
}
function Txt($cv, [string]$t, [float]$size, $color, [float]$cx, [float]$cy, [string]$font = 'Bahnschrift', [string]$style = 'Bold') {
    $f = New-Object Drawing.Font $font, $size, ([Drawing.FontStyle]$style), ([Drawing.GraphicsUnit]::Pixel)
    $sf = New-Object Drawing.StringFormat; $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
    $cv.DrawString($t, $f, (New-Object Drawing.SolidBrush $color), ([Drawing.RectangleF]::new($cx - 200, $cy - 200, 400, 400)), $sf)
}
function Line($cv, $color, [float]$w, [float[]]$pts, [switch]$Arrow) {
    $pen = New-Object Drawing.Pen $color, $w; $pen.StartCap = 'Round'; $pen.EndCap = 'Round'; $pen.LineJoin = 'Round'
    if ($Arrow) { $pen.CustomEndCap = New-Object Drawing.Drawing2D.AdjustableArrowCap 3, 3 }
    $p = for ($i = 0; $i -lt $pts.Count; $i += 2) { [Drawing.PointF]::new($pts[$i], $pts[$i + 1]) }
    $cv.DrawLines($pen, [Drawing.PointF[]]$p)
}
function Poly($cv, $color, [float[]]$pts) {
    $p = for ($i = 0; $i -lt $pts.Count; $i += 2) { [Drawing.PointF]::new($pts[$i], $pts[$i + 1]) }
    $cv.FillPolygon((New-Object Drawing.SolidBrush $color), [Drawing.PointF[]]$p)
}
function LogoTile($cv) {
    $b = New-Object Drawing.Drawing2D.LinearGradientBrush ([Drawing.PointF]::new(0, 0)), ([Drawing.PointF]::new(128, 128)), $NAVY, $BLUE
    $cv.FillPath($b, (RoundRect 0 0 128 128 28))
}

#---- the icons --------------------------------------------------------------------
$icons = [ordered]@{
    logo       = { param($cv) LogoTile $cv; Txt $cv 'ILL' 46 $WHITE 64 41; Txt $cv 'XL' 46 $CYAN 64 87 }
    presets    = { param($cv) Fill $cv $CYAN 40 10 76 56 12; Fill $cv $BLUE 26 36 76 56 12; Fill $cv $NAVY 12 62 76 56 12
                   Txt $cv 'Aa' 30 $WHITE 50 90 }
    perf       = { param($cv) Poly $cv $AMBER @(74, 6, 26, 72, 58, 72, 46, 122, 102, 48, 68, 48, 84, 6) }
    refabs     = { param($cv) Txt $cv '$' 118 $BLUE 64 66 }
    refrel     = { param($cv) Txt $cv 'A1' 74 $BLUE 64 66 }
    otherfmt   = { param($cv) Txt $cv 'A' 70 $INK 38 66; Txt $cv 'E' 70 $BLUE 92 66 }
    scalediv   = { param($cv) Txt $cv ([string][char]0xF7) 64 $INK 34 38; Txt $cv 'k' 92 $BLUE 80 76 }
    scalemul   = { param($cv) Txt $cv ([string][char]0xD7) 64 $INK 34 38; Txt $cv 'k' 92 $BLUE 80 76 }
    sign       = { param($cv) Txt $cv ([string][char]0xB1) 118 $BLUE 64 62 }
    div100     = { param($cv) Txt $cv ([string][char]0xF7) 60 $INK 32 38; Txt $cv '%' 80 $BLUE 80 76 }
    mul100     = { param($cv) Txt $cv ([string][char]0xD7) 60 $INK 32 38; Txt $cv '%' 80 $BLUE 80 76 }
    cagr       = { param($cv) Line $cv $INK 8 @(14, 10, 14, 114, 118, 114)
                   Line $cv $BLUE 10 @(26, 100, 52, 92, 78, 74, 104, 26) -Arrow }
    pctchg     = { param($cv) Txt $cv ([string][char]0x394) 92 $INK 44 62; Txt $cv '%' 56 $BLUE 98 86 }
    eqleft     = { param($cv) Line $cv $BLUE 10 @(64, 64, 18, 64) -Arrow; Txt $cv '=' 84 $INK 94 60 }
    growth     = { param($cv) Line $cv $GREEN 12 @(12, 104, 46, 70, 70, 88, 112, 34) -Arrow }
    autocolor  = { param($cv) Fill $cv $BLUE 8 8 52 52 10; Fill $cv (C '#111827') 68 8 52 52 10
                   Fill $cv $GREEN 8 68 52 52 10; Fill $cv $RED 68 68 52 52 10 }
    inputstyle = { param($cv) Fill $cv (C '#FFF2CC') 10 22 108 84 8; Stroke $cv $INK 5 10 22 108 84 8 -Dotted
                   Txt $cv '1.0' 52 $BLUE 64 66 }
    headerstyle= { param($cv) Fill $cv $NAVY 6 30 116 68 8; Fill $cv $WHITE 24 56 80 16 6 }
    headers    = { param($cv) foreach ($x in 4, 46, 88) { Fill $cv $NAVY $x 38 36 52 6; Fill $cv $WHITE ($x + 8) 58 20 10 4 } }
    variance   = { param($cv) Fill $cv $NAVY 6 18 116 92 12; Txt $cv ([string][char]0x394) 72 $WHITE 46 64; Txt $cv '%' 44 $CYAN 94 74 }
    zerocheck  = { param($cv) Fill $cv (C '#DCFCE7') 10 10 108 108 16; Stroke $cv $GREEN 6 10 10 108 108 16
                   Line $cv $GREEN 13 @(34, 66, 56, 88, 96, 40) }
    zeroclear  = { param($cv) Fill $cv (C '#FEE2E2') 10 10 108 108 16; Stroke $cv $RED 6 10 10 108 108 16
                   Line $cv $RED 12 @(40, 40, 88, 88); Line $cv $RED 12 @(88, 40, 40, 88) }
    tabcolor   = { param($cv) Poly $cv $INK @(8, 104, 22, 34, 106, 34, 120, 104); Poly $cv $WHITE @(18, 98, 29, 42, 99, 42, 110, 98)
                   Fill $cv $AMBER 22 82 84 14 3 }
    tagvalue   = { param($cv) Txt $cv '[#]' 70 $BLUE 64 64 }
    tagduration= { param($cv) Txt $cv '[d]' 70 $BLUE 64 64 }
    tagrate    = { param($cv) Txt $cv '[%]' 66 $BLUE 64 64 }
    tagremove  = { param($cv) Txt $cv '[ ]' 70 $INK 64 64; Line $cv $RED 11 @(26, 104, 102, 24) }
    cleanstyles= { param($cv) Txt $cv 'Aa' 70 $INK 58 56; Line $cv $RED 10 @(70, 118, 118, 70); Line $cv $RED 10 @(70, 70, 118, 118) }
}

#---- render ---------------------------------------------------------------------------
function Render($draw, [int]$size, [string]$path) {
    $big = New-Object Drawing.Bitmap 128, 128
    $cv = [Drawing.Graphics]::FromImage($big)
    $cv.SmoothingMode = 'AntiAlias'; $cv.TextRenderingHint = 'AntiAlias'; $cv.Clear([Drawing.Color]::Transparent)
    & $draw $cv
    $small = New-Object Drawing.Bitmap $size, $size
    $sg = [Drawing.Graphics]::FromImage($small)
    $sg.InterpolationMode = 'HighQualityBicubic'; $sg.PixelOffsetMode = 'HighQuality'; $sg.CompositingQuality = 'HighQuality'
    $sg.DrawImage($big, 0, 0, $size, $size)
    $small.Save($path, [Drawing.Imaging.ImageFormat]::Png)
    $sg.Dispose(); $cv.Dispose(); $small.Dispose(); $big.Dispose()
}
foreach ($name in $icons.Keys) {
    Render $icons[$name] 32 "$out\$name.png"
    Render $icons[$name] 64 "$out\${name}_L.png"
}
Render $icons['logo'] 512 (Join-Path $PSScriptRoot '..\assets\logo-512.png')
Render $icons['logo'] 256 (Join-Path $PSScriptRoot '..\assets\logo-256.png')

#---- contact sheet for review (tools\icons-sheet.png, not shipped) -------------------
$names = @($icons.Keys); $cols = 8; $cell = 120
$sheet = New-Object Drawing.Bitmap ($cols * $cell), ([math]::Ceiling($names.Count / $cols) * $cell)
$sh = [Drawing.Graphics]::FromImage($sheet); $sh.Clear([Drawing.Color]::FromArgb(243, 243, 243)); $sh.TextRenderingHint = 'AntiAlias'
for ($i = 0; $i -lt $names.Count; $i++) {
    $x = ($i % $cols) * $cell; $y = [math]::Floor($i / $cols) * $cell
    $L = [Drawing.Image]::FromFile((Resolve-Path "$out\$($names[$i])_L.png")); $S = [Drawing.Image]::FromFile((Resolve-Path "$out\$($names[$i]).png"))
    $sh.DrawImage($L, $x + 8, $y + 10, 64, 64); $sh.DrawImage($S, $x + 80, $y + 26, 32, 32)
    Txt $sh $names[$i] 12 $INK ($x + 60) ($y + 96) 'Segoe UI' 'Regular'
    $L.Dispose(); $S.Dispose()
}
$sheet.Save((Join-Path $PSScriptRoot 'icons-sheet.png'))
"icons: $($names.Count)  -> $out"
