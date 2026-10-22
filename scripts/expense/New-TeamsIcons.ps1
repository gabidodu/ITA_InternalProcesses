<#
.SYNOPSIS
    Generates appPackage/color.png (192x192) and appPackage/outline.png (32x32)
    for the Decontări (expense reports) Teams app: a receipt/document glyph
    with a folded corner, a few "line item" bars, and — on the color icon
    only — a green checkmark badge for "approved/reimbursed".

.DESCRIPTION
    Replaces the placeholder icons borrowed from the sibling
    LeaveRequestForTeams app. Drawn programmatically via System.Drawing
    (no external asset), matching Teams' exact size requirements:
    color.png must be 192x192, outline.png must be 32x32, solid white on a
    transparent background (no color/gradients), per Teams manifest schema
    validation rules.

.EXAMPLE
    ./New-TeamsIcons.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$appPackageDir = Join-Path $PSScriptRoot "..\appPackage"
$accent = [System.Drawing.Color]::FromArgb(255, 70, 71, 117)      # #464775 — manifest.json accentColor
$foldTint = [System.Drawing.Color]::FromArgb(255, 214, 213, 232)  # light lavender, folded-corner shading
$lineColor = $accent
$totalColor = [System.Drawing.Color]::FromArgb(255, 224, 164, 88) # warm amber — the "total" bar
$badgeGreen = [System.Drawing.Color]::FromArgb(255, 47, 168, 79)

function New-RoundedRectPath {
    param([float]$X, [float]$Y, [float]$W, [float]$H, [float]$R)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $R * 2
    $path.AddArc($X, $Y, $d, $d, 180, 90)
    $path.AddArc($X + $W - $d, $Y, $d, $d, 270, 90)
    $path.AddArc($X + $W - $d, $Y + $H - $d, $d, $d, 0, 90)
    $path.AddArc($X, $Y + $H - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

# ---------- color.png (192x192) ----------
$bmp = New-Object System.Drawing.Bitmap(192, 192)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

# Full-bleed accent background — Teams applies its own corner rounding when
# displaying the icon, so a plain filled square is the standard/safe choice.
$g.Clear($accent)

# The receipt/document, with a folded top-right corner.
$docPath = New-Object System.Drawing.Drawing2D.GraphicsPath
$docPoints = @(
    [System.Drawing.PointF]::new(48, 32),
    [System.Drawing.PointF]::new(126, 32),
    [System.Drawing.PointF]::new(144, 50),
    [System.Drawing.PointF]::new(144, 160),
    [System.Drawing.PointF]::new(48, 160)
)
$docPath.AddPolygon([System.Drawing.PointF[]]$docPoints)
$whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$g.FillPath($whiteBrush, $docPath)

# Folded-corner flap shading.
$foldPath = New-Object System.Drawing.Drawing2D.GraphicsPath
$foldPoints = @(
    [System.Drawing.PointF]::new(126, 32),
    [System.Drawing.PointF]::new(144, 32),
    [System.Drawing.PointF]::new(144, 50)
)
$foldPath.AddPolygon([System.Drawing.PointF[]]$foldPoints)
$foldBrush = New-Object System.Drawing.SolidBrush($foldTint)
$g.FillPath($foldBrush, $foldPath)

# Three "line item" bars — the third (total) in a distinct amber tone.
$lineBrush = New-Object System.Drawing.SolidBrush($lineColor)
$totalBrush = New-Object System.Drawing.SolidBrush($totalColor)
$bar1 = New-RoundedRectPath -X 66 -Y 64 -W 62 -H 8 -R 4
$bar2 = New-RoundedRectPath -X 66 -Y 84 -W 62 -H 8 -R 4
$bar3 = New-RoundedRectPath -X 66 -Y 104 -W 42 -H 10 -R 5
$g.FillPath($lineBrush, $bar1)
$g.FillPath($lineBrush, $bar2)
$g.FillPath($totalBrush, $bar3)

# Green "approved / reimbursed" checkmark badge, bottom-right.
$badgeBrush = New-Object System.Drawing.SolidBrush($badgeGreen)
$g.FillEllipse($badgeBrush, 118, 128, 40, 40)
$checkPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 5)
$checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$checkPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
$checkPoints = @(
    [System.Drawing.Point]::new(127, 148),
    [System.Drawing.Point]::new(135, 156),
    [System.Drawing.Point]::new(150, 137)
)
$g.DrawLines($checkPen, [System.Drawing.Point[]]$checkPoints)

$colorPath = Join-Path $appPackageDir "color.png"
$bmp.Save($colorPath, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()
Write-Host "Wrote $colorPath (192x192)" -ForegroundColor Green

# ---------- outline.png (32x32) ----------
# Teams requires a single-tone white glyph on a transparent background here —
# no accent fill, no second color. Same silhouette, scaled down, with two
# thin transparent gaps punched out to hint at the line-item texture.
$obmp = New-Object System.Drawing.Bitmap(32, 32)
$og = [System.Drawing.Graphics]::FromImage($obmp)
$og.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$og.Clear([System.Drawing.Color]::Transparent)

$oDocPath = New-Object System.Drawing.Drawing2D.GraphicsPath
$oDocPoints = @(
    [System.Drawing.PointF]::new(7, 4),
    [System.Drawing.PointF]::new(20, 4),
    [System.Drawing.PointF]::new(26, 10),
    [System.Drawing.PointF]::new(26, 28),
    [System.Drawing.PointF]::new(7, 28)
)
$oDocPath.AddPolygon([System.Drawing.PointF[]]$oDocPoints)
$oWhiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$og.FillPath($oWhiteBrush, $oDocPath)

# Punch transparent gaps (SourceCopy so alpha is actually replaced, not blended).
$og.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
$clearBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Transparent)
$og.FillRectangle($clearBrush, 10, 12, 13, 2)
$og.FillRectangle($clearBrush, 10, 17, 13, 2)
$og.FillRectangle($clearBrush, 10, 22, 9, 2)

$outlinePath = Join-Path $appPackageDir "outline.png"
$obmp.Save($outlinePath, [System.Drawing.Imaging.ImageFormat]::Png)
$og.Dispose()
$obmp.Dispose()
Write-Host "Wrote $outlinePath (32x32)" -ForegroundColor Green
