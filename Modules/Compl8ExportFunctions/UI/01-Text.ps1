#region Text UI Helpers

# ANSI color map for dashboard rendering (matches PowerShell ConsoleColor names)
$script:AnsiColorMap = @{
    'Black'       = "`e[30m"
    'DarkRed'     = "`e[31m"
    'DarkGreen'   = "`e[32m"
    'DarkYellow'  = "`e[33m"
    'DarkBlue'    = "`e[34m"
    'DarkMagenta' = "`e[35m"
    'DarkCyan'    = "`e[36m"
    'Gray'        = "`e[37m"
    'DarkGray'    = "`e[90m"
    'Red'         = "`e[91m"
    'Green'       = "`e[92m"
    'Yellow'      = "`e[93m"
    'Blue'        = "`e[94m"
    'Magenta'     = "`e[95m"
    'Cyan'        = "`e[96m"
    'White'       = "`e[97m"
}
$script:AnsiReset = "`e[0m"

function Get-TerminalSize {
    <#
    .SYNOPSIS
        Returns the current terminal viewport dimensions.
    .DESCRIPTION
        Single source of truth for terminal width/height. Reads on every call
        (no caching) so resizes are picked up. Clamps to sensible minimums and
        subtracts 1 from width to avoid last-column line-wrap artifacts.
    #>
    try {
        @{
            Width  = [Math]::Max(40, [Console]::WindowWidth - 1)
            Height = [Math]::Max(10, [Console]::WindowHeight)
        }
    }
    catch {
        @{ Width = 119; Height = 30 }
    }
}

function Format-ProgressBar {
    <#
    .SYNOPSIS
        Returns a width-proportional progress bar string like [####........].
    .PARAMETER Percent
        Completion percentage (0-100).
    .PARAMETER MaxWidth
        Maximum character width for the bar including brackets. Defaults to
        a fraction of terminal width (clamped between 20 and 50).
    #>
    [CmdletBinding()]
    param(
        [double]$Percent = 0,
        [int]$MaxWidth = 0
    )

    if ($MaxWidth -le 0) {
        $termWidth = (Get-TerminalSize).Width
        $MaxWidth = [Math]::Min(50, [Math]::Max(20, [int]($termWidth * 0.35)))
    }

    $inner = $MaxWidth - 2  # subtract brackets
    $filled = [Math]::Min($inner, [int]($Percent / 100 * $inner))
    $empty = $inner - $filled
    return "[" + ("#" * $filled) + ("." * $empty) + "]"
}

function Write-BoxTop {
    <#
    .SYNOPSIS
        Writes the top border of a box: ╔═══════╗ or ┌───────┐
    #>
    param([int]$InnerWidth, [string]$Color = 'Cyan', [switch]$Double)
    $h = if ($Double) { [char]0x2550 } else { [char]0x2500 }  # ═ or ─
    $tl = if ($Double) { [char]0x2554 } else { [char]0x250C }  # ╔ or ┌
    $tr = if ($Double) { [char]0x2557 } else { [char]0x2510 }  # ╗ or ┐
    Write-Host ("  {0}{1}{2}" -f $tl, ([string]$h * $InnerWidth), $tr) -ForegroundColor $Color
}

function Write-BoxBottom {
    <#
    .SYNOPSIS
        Writes the bottom border of a box: ╚═══════╝ or └───────┘
    #>
    param([int]$InnerWidth, [string]$Color = 'Cyan', [switch]$Double)
    $h = if ($Double) { [char]0x2550 } else { [char]0x2500 }
    $bl = if ($Double) { [char]0x255A } else { [char]0x2514 }  # ╚ or └
    $br = if ($Double) { [char]0x255D } else { [char]0x2518 }  # ╝ or ┘
    Write-Host ("  {0}{1}{2}" -f $bl, ([string]$h * $InnerWidth), $br) -ForegroundColor $Color
}

function Write-BoxLine {
    <#
    .SYNOPSIS
        Writes a content line inside a box: ║ text ║ or │ text │
    .DESCRIPTION
        Pads or truncates text to fit InnerWidth. Text longer than available
        space is truncated with an ellipsis character.
    #>
    param(
        [string]$Text = '',
        [int]$InnerWidth,
        [string]$Color = 'Cyan',
        [switch]$Double
    )
    $v = if ($Double) { [char]0x2551 } else { [char]0x2502 }  # ║ or │
    $contentWidth = $InnerWidth - 2  # 1 space padding each side
    if ($Text.Length -gt $contentWidth) {
        $Text = $Text.Substring(0, $contentWidth - 1) + [char]0x2026  # …
    }
    $padded = $Text.PadRight($contentWidth)
    Write-Host ("  {0} {1} {2}" -f $v, $padded, $v) -ForegroundColor $Color
}

function Write-BoxSeparator {
    <#
    .SYNOPSIS
        Writes a separator line inside a box: ╠═══════╣ or ├───────┤
    #>
    param([int]$InnerWidth, [string]$Color = 'Cyan', [switch]$Double)
    $h = if ($Double) { [char]0x2550 } else { [char]0x2500 }
    $ml = if ($Double) { [char]0x2560 } else { [char]0x251C }  # ╠ or ├
    $mr = if ($Double) { [char]0x2563 } else { [char]0x2524 }  # ╣ or ┤
    Write-Host ("  {0}{1}{2}" -f $ml, ([string]$h * $InnerWidth), $mr) -ForegroundColor $Color
}

function Get-BoxInnerWidth {
    <#
    .SYNOPSIS
        Computes the inner width for a box based on terminal width.
    .PARAMETER MaxWidth
        Maximum inner width. The result is clamped to (terminal width - 6)
        to leave room for the 2-char indent and the border characters.
    #>
    param([int]$MaxWidth = 62)
    $termWidth = (Get-TerminalSize).Width
    return [Math]::Min($MaxWidth, $termWidth - 6)
}

function Write-SectionHeader {
    <#
    .SYNOPSIS
        Writes an adaptive-width section header: ── Text ──────────
    .PARAMETER Text
        The header text.
    .PARAMETER Color
        Foreground color. Default: DarkCyan.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Text,
        [string]$Color = 'DarkCyan'
    )
    $termWidth = (Get-TerminalSize).Width
    $h = [char]0x2500  # ─
    $prefix = "  ${h}${h} "
    $suffix = " "
    $remaining = $termWidth - $prefix.Length - $Text.Length - $suffix.Length - 2
    if ($remaining -lt 2) { $remaining = 2 }
    $trail = [string]$h * $remaining
    Write-Host "${prefix}${Text}${suffix}${trail}" -ForegroundColor $Color
}

function Write-Banner {
    <#
    .SYNOPSIS
        Writes a bordered banner box with a title and optional content lines.
    .DESCRIPTION
        Adapts to terminal width. Uses double-line borders for emphasis.
        All content is padded/truncated to ensure the right border is always present.
    .PARAMETER Title
        The banner title (displayed on first content line).
    .PARAMETER Lines
        Optional additional content lines.
    .PARAMETER Color
        Border and text color. Default: Cyan.
    .PARAMETER Double
        Use double-line border characters. Default: True.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        [string[]]$Lines = @(),
        [string]$Color = 'Cyan',
        [switch]$Single
    )

    $double = -not $Single
    $innerWidth = Get-BoxInnerWidth -MaxWidth 62
    Write-BoxTop -InnerWidth $innerWidth -Color $Color -Double:$double
    Write-BoxLine -Text $Title -InnerWidth $innerWidth -Color $Color -Double:$double
    if ($Lines.Count -gt 0) {
        Write-BoxSeparator -InnerWidth $innerWidth -Color $Color -Double:$double
        foreach ($line in $Lines) {
            Write-BoxLine -Text $line -InnerWidth $innerWidth -Color $Color -Double:$double
        }
    }
    Write-BoxBottom -InnerWidth $innerWidth -Color $Color -Double:$double
}

function Write-DashboardFrame {
    <#
    .SYNOPSIS
        Renders a dashboard frame in-place using ANSI escape sequences.
    .DESCRIPTION
        Builds the entire frame in a StringBuilder and emits it with a single
        [Console]::Write() call for flicker-free rendering. Uses ANSI cursor-up
        (relative positioning) instead of absolute SetCursorPosition to avoid
        scroll drift. Hides the cursor during redraw.

        Each line in $Lines should be a hashtable with Text and Color keys.
        The Color value should be a PowerShell ConsoleColor name.
    .PARAMETER Lines
        Array of @{ Text = '...'; Color = 'Cyan' } hashtables.
    .PARAMETER PreviousLineCount
        The line count from the previous frame render. Used to calculate how
        many lines to cursor-up. Pass 0 on first render.
    .OUTPUTS
        [int] The number of lines rendered (pass back as PreviousLineCount next call).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.ArrayList]$Lines,

        [int]$PreviousLineCount = 0
    )

    $term = Get-TerminalSize
    $width = $term.Width
    $height = $term.Height

    # Clamp to terminal height (leave 2 lines safety margin)
    $maxLines = $height - 2
    $renderCount = [Math]::Min($Lines.Count, $maxLines)

    # Build frame in StringBuilder for single-write output
    $sb = [System.Text.StringBuilder]::new(4096)

    # Hide cursor
    [void]$sb.Append("`e[?25l")

    # Move cursor up to overwrite previous frame (relative, scroll-safe)
    if ($PreviousLineCount -gt 0) {
        [void]$sb.Append("`e[${PreviousLineCount}A`e[0G")
    }

    $prevAnsi = ''
    $linesWritten = 0

    for ($i = 0; $i -lt $renderCount; $i++) {
        $line = $Lines[$i]
        $text = if ($line.Text) { $line.Text } else { '' }
        $colorName = if ($line.Color) { $line.Color } else { 'White' }

        # Truncate to terminal width
        if ($text.Length -gt $width) { $text = $text.Substring(0, $width) }

        # Apply ANSI color (only emit if changed from previous line)
        $ansi = if ($script:AnsiColorMap.ContainsKey($colorName)) { $script:AnsiColorMap[$colorName] } else { "`e[37m" }
        if ($ansi -ne $prevAnsi) {
            [void]$sb.Append($ansi)
            $prevAnsi = $ansi
        }

        [void]$sb.Append($text)
        [void]$sb.Append("`e[K")  # Erase to end of line (clears stale chars)
        [void]$sb.Append("`n")
        $linesWritten++
    }

    # Clear leftover lines from previous taller frame
    if ($PreviousLineCount -gt $linesWritten) {
        $extra = $PreviousLineCount - $linesWritten
        for ($j = 0; $j -lt $extra; $j++) {
            [void]$sb.Append("`e[K`n")
            $linesWritten++
        }
    }

    # Reset colors and show cursor
    [void]$sb.Append("`e[0m`e[?25h")

    [Console]::Write($sb.ToString())

    return $linesWritten
}

#endregion

