# FileCollector V3.0 - Icon Generator
# Generates Resources/app.ico (multi-resolution) with brand palette (POSTEX purple + orange)
# Concept: orange folder with purple documents being collected into it

[CmdletBinding()]
param(
    [string]$OutputPath = "$PSScriptRoot\..\Resources\app.ico"
)

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Drawing.Common -ErrorAction SilentlyContinue

# Brand colors (from POSTEX logo)
$colorOrange       = [System.Drawing.Color]::FromArgb(255, 237, 111, 48)
$colorOrangeDark   = [System.Drawing.Color]::FromArgb(255, 196,  80, 24)
$colorOrangeLight  = [System.Drawing.Color]::FromArgb(255, 252, 142,  82)
$colorPurple       = [System.Drawing.Color]::FromArgb(255,  92,  45, 140)
$colorPurpleDark   = [System.Drawing.Color]::FromArgb(255,  64,  28, 110)
$colorWhite        = [System.Drawing.Color]::White
$colorShadow       = [System.Drawing.Color]::FromArgb(70, 0, 0, 0)

function New-RoundedRectPath {
    param(
        [single]$x, [single]$y, [single]$w, [single]$h, [single]$r
    )
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $r * 2
    $path.AddArc($x,            $y,            $d, $d, 180, 90)  | Out-Null
    $path.AddArc($x + $w - $d,  $y,            $d, $d, 270, 90)  | Out-Null
    $path.AddArc($x + $w - $d,  $y + $h - $d,  $d, $d,   0, 90)  | Out-Null
    $path.AddArc($x,            $y + $h - $d,  $d, $d,  90, 90)  | Out-Null
    $path.CloseFigure()
    return $path
}

function New-IconBitmap {
    param([int]$Size)

    $bmp = New-Object System.Drawing.Bitmap($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)

    # Coordinates relative to a 256x256 design canvas
    $scale = $Size / 256.0

    # ---------- DOCUMENTS (3 stacked, slightly tilted) ----------
    # We draw them BEFORE the folder so the folder front overlaps them.
    function Draw-Doc {
        param([single]$cx, [single]$cy, [single]$angle, [System.Drawing.Color]$accent)
        $w = 92 * $scale
        $h = 116 * $scale
        $state = $g.Save()
        $g.TranslateTransform($cx * $scale, $cy * $scale)
        $g.RotateTransform($angle)
        $rect = New-Object System.Drawing.RectangleF(([single](-$w / 2)), ([single](-$h / 2)), $w, $h)
        # Shadow
        $shadowBrush = New-Object System.Drawing.SolidBrush($colorShadow)
        $shadowRect = New-Object System.Drawing.RectangleF(([single]($rect.X + 3 * $scale)), ([single]($rect.Y + 4 * $scale)), $rect.Width, $rect.Height)
        $g.FillRectangle($shadowBrush, $shadowRect)
        $shadowBrush.Dispose()
        # Paper
        $paperBrush = New-Object System.Drawing.SolidBrush($colorWhite)
        $g.FillRectangle($paperBrush, $rect)
        $paperBrush.Dispose()
        # Top accent stripe
        $stripeRect = New-Object System.Drawing.RectangleF($rect.X, $rect.Y, $rect.Width, ([single](14 * $scale)))
        $stripeBrush = New-Object System.Drawing.SolidBrush($accent)
        $g.FillRectangle($stripeBrush, $stripeRect)
        $stripeBrush.Dispose()
        # Body lines (the "text")
        $lineBrush = New-Object System.Drawing.SolidBrush($colorPurple)
        $lineH = 6 * $scale
        $lineX = $rect.X + 10 * $scale
        $maxLineW = $rect.Width - 20 * $scale
        $lineY = $rect.Y + 28 * $scale
        $widths = @(1.0, 0.8, 0.95, 0.6, 0.85)
        foreach ($wf in $widths) {
            $lr = New-Object System.Drawing.RectangleF($lineX, $lineY, ([single]($maxLineW * $wf)), $lineH)
            $g.FillRectangle($lineBrush, $lr)
            $lineY += 14 * $scale
        }
        $lineBrush.Dispose()
        $g.Restore($state)
    }

    if ($Size -ge 32) {
        # Three documents fanned out behind the folder
        Draw-Doc -cx 70  -cy 90  -angle -14 -accent $colorPurple
        Draw-Doc -cx 186 -cy 86  -angle  14 -accent $colorOrange
        Draw-Doc -cx 128 -cy 78  -angle   0 -accent $colorPurpleDark
    } elseif ($Size -ge 24) {
        # Just one document for medium sizes
        Draw-Doc -cx 128 -cy 80  -angle 0 -accent $colorPurple
    }

    # ---------- FOLDER ----------
    # Tab (top-left)
    $tabPath = New-RoundedRectPath -x (24 * $scale) -y (110 * $scale) -w (96 * $scale) -h (28 * $scale) -r (8 * $scale)
    $tabBrush = New-Object System.Drawing.SolidBrush($colorOrangeDark)
    $g.FillPath($tabBrush, $tabPath)
    $tabBrush.Dispose()
    $tabPath.Dispose()

    # Folder body (rounded rectangle)
    $folderPath = New-RoundedRectPath -x (20 * $scale) -y (130 * $scale) -w (216 * $scale) -h (110 * $scale) -r (14 * $scale)
    # Gradient fill
    $folderRect = New-Object System.Drawing.RectangleF(([single](20 * $scale)), ([single](130 * $scale)), ([single](216 * $scale)), ([single](110 * $scale)))
    $folderBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $folderRect, $colorOrangeLight, $colorOrange,
        [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
    $g.FillPath($folderBrush, $folderPath)
    $folderBrush.Dispose()

    # Folder front lip (slightly darker band at top)
    $lipPath = New-RoundedRectPath -x (20 * $scale) -y (130 * $scale) -w (216 * $scale) -h (24 * $scale) -r (12 * $scale)
    $lipBrush = New-Object System.Drawing.SolidBrush($colorOrangeDark)
    $g.FillPath($lipBrush, $lipPath)
    $lipBrush.Dispose()
    $lipPath.Dispose()

    # Subtle outline
    if ($Size -ge 48) {
        $borderPen = New-Object System.Drawing.Pen($colorOrangeDark, ([single](2 * $scale)))
        $g.DrawPath($borderPen, $folderPath)
        $borderPen.Dispose()
    }
    $folderPath.Dispose()

    # Down-arrow on folder front (only at large sizes) - hints at "collection"
    if ($Size -ge 64) {
        $arrowPath = New-Object System.Drawing.Drawing2D.GraphicsPath
        $cx = 128 * $scale
        $cy = 195 * $scale
        $aw = 30 * $scale
        $ah = 30 * $scale
        # Arrow shaft
        $arrowPath.AddRectangle((New-Object System.Drawing.RectangleF(
            ([single]($cx - $aw * 0.18)), ([single]($cy - $ah * 0.5)),
            ([single]($aw * 0.36)), ([single]($ah * 0.55))))) | Out-Null
        # Arrow head (triangle)
        $tri = New-Object 'System.Drawing.PointF[]' 3
        $tri[0] = New-Object System.Drawing.PointF(([single]($cx - $aw * 0.5)), ([single]($cy + $ah * 0.05)))
        $tri[1] = New-Object System.Drawing.PointF(([single]($cx + $aw * 0.5)), ([single]($cy + $ah * 0.05)))
        $tri[2] = New-Object System.Drawing.PointF(([single]$cx),                ([single]($cy + $ah * 0.55)))
        $arrowPath.AddPolygon($tri) | Out-Null

        $arrowBrush = New-Object System.Drawing.SolidBrush($colorWhite)
        $g.FillPath($arrowBrush, $arrowPath)
        $arrowBrush.Dispose()
        $arrowPath.Dispose()
    }

    $g.Dispose()
    return $bmp
}

function Save-IcoMultiSize {
    param(
        [int[]]$Sizes,
        [string]$Path
    )

    $pngs = @()
    foreach ($s in $Sizes) {
        $bmp = New-IconBitmap -Size $s
        $ms = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $pngs += [PSCustomObject]@{ Size = $s; Bytes = $ms.ToArray() }
        $ms.Dispose()
        $bmp.Dispose()
    }

    $headerSize = 6
    $entrySize  = 16
    $offset = $headerSize + $entrySize * $pngs.Count

    $out = New-Object System.IO.MemoryStream
    $w   = New-Object System.IO.BinaryWriter($out)
    # ICONDIR
    $w.Write([UInt16]0)              # reserved
    $w.Write([UInt16]1)              # type = icon
    $w.Write([UInt16]$pngs.Count)    # number of images

    # ICONDIRENTRY for each
    $entries = @()
    foreach ($p in $pngs) {
        $entry = [PSCustomObject]@{ Size = $p.Size; Length = $p.Bytes.Length; Offset = $offset }
        $entries += $entry

        $bw = if ($p.Size -ge 256) { [byte]0 } else { [byte]$p.Size }
        $bh = $bw

        $w.Write([byte]$bw)              # width
        $w.Write([byte]$bh)              # height
        $w.Write([byte]0)                # color count (0 for >=256 colors)
        $w.Write([byte]0)                # reserved
        $w.Write([UInt16]1)              # planes
        $w.Write([UInt16]32)             # bits per pixel
        $w.Write([UInt32]$p.Bytes.Length)  # bytes in resource
        $w.Write([UInt32]$offset)        # image offset

        $offset += $p.Bytes.Length
    }

    foreach ($p in $pngs) {
        $w.Write($p.Bytes)
    }

    $w.Flush()
    $bytes = $out.ToArray()
    $w.Dispose()
    $out.Dispose()

    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    [System.IO.File]::WriteAllBytes($Path, $bytes)

    Write-Host "ICO written: $Path ($($bytes.Length) bytes, $($pngs.Count) sizes)"
    foreach ($e in $entries) {
        Write-Host ("  - {0,3}x{0,-3}  {1,7} bytes  @ offset {2}" -f $e.Size, $e.Length, $e.Offset)
    }
}

Save-IcoMultiSize -Sizes @(16, 24, 32, 48, 64, 128, 256) -Path $OutputPath
