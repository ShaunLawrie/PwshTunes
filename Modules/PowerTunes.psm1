# https://www.undocumented-features.com/2020/10/23/announcing-the-end-of-a-script-with-a-powershell-music/

$script:Notes = @{
    "D#1" = 38.89
    Eb1 = 38.89
    E1 = 41.20
    F1 = 43.65
    "F#1" = 46.25
    Gb1 = 46.25
    G1 = 49.00
    "G#1" = 51.91
    Ab1 = 51.91
    A1 = 55.00
    "A#1" = 58.27
    Bb1 = 58.27
    B1 = 61.74
    C2 = 65.41
    "C#2" = 69.30
    Cb2 = 69.30
    D2 = 73.42
    "D#2" = 77.78
    Db2 = 77.78
    E2 = 82.41
    F2 = 87.31
    "F#2" = 92.50
    Fb2 = 92.50
    G2 = 98.00
    "G#2" = 103.83
    Ab2 = 103.83
    A2 = 110.00
    "A#2" = 116.54
    Bb2 = 116.54
    B2 = 123.47
    C3 = 130.81
    "C#3" = 138.59
    Db3 = 138.59
    D3 = 146.83
    "D#3" = 155.56
    Eb3 = 155.56
    E3 = 164.81
    F3 = 174.61
    "F#3" = 185.00
    Gb3 = 185.00
    G3 = 196.00
    "G#3" = 207.65
    Ab3 = 207.65
    A3 = 220.00
    "A#3" = 233.08
    Bb3 = 233.08
    B3 = 246.94
    C4 = 261.63 # Middle C
    "C#4" = 277.18
    Db4 = 277.18
    D4 = 293.66
    "D#4" = 311.13
    Eb4 = 311.13
    E4 = 329.63
    F4 = 349.23
    "F#4" = 369.99
    Gb4 = 369.99
    G4 = 392.00
    "G#4" = 415.30
    Ab4 = 415.30
    A4 = 440.00
    "A#4" = 466.16
    Bb4 = 466.16
    B4 = 493.88
    C5 = 523.25
    "C#5" = 554.37
    Db5 = 554.37
    D5 = 587.33
    "D#5" = 622.25
    Eb5 = 622.25
    E5 = 659.26
    F5 = 698.46
    "F#5" = 739.99
    Gb5 = 739.99
    G5 = 783.99
    "G#5" = 830.61
    Ab5 = 830.61
    A5 = 880.00
    "A#5" = 932.33
    Bb5 = 932.33
    B5 = 987.77
    C6 = 1046.50
    "C#6" = 1108.73
    Db6 = 1108.73
    D6 = 1174.66
    "D#6" = 1244.51
    Eb6 = 1244.51
    E6 = 1318.51
    F6 = 1396.91
    "F#6" = 1479.98
    Gb6 = 1479.98
    G6 = 1567.98
    "G#6" = 1661.22
    Ab6 = 1661.22
    A6 = 1760.00
    "A#6" = 1864.66
    Bb6 = 1864.66
    B6 = 1975.53
    C7 = 2093.00
    "C#7" = 2217.46
    Db7 = 2217.46
    D7 = 2349.32
    "D#7" = 2489.02
    Eb7 = 2489.02
    E7 = 2637.02
    F7 = 2793.83
    "F#7" = 2959.96
    Gb7 = 2959.96
    G7 = 3135.96
    "G#7" = 3322.44
    Ab7 = 3322.44
    A7 = 3520.00
    "A#7" = 3729.31
    Bb7 = 3729.31
    B7 = 3951.07
    C8 = 4186.01
    "C#8" = 4434.92
    Db8 = 4434.92
    D8 = 4698.64
    "D#8" = 4978.03
}

$script:NoteBreak = 0
$script:QuarterNote = 500
$script:Durations

Function Update-Durations {
    [CmdletBinding()] param()
    $script:Durations = @{
        Longa = [int] ($script:QuarterNote * 16 - $script:NoteBreak)
        DottedDoubleWhole = [int] ($script:QuarterNote * 12 - $script:NoteBreak)
        DoubleWhole = [int] ($script:QuarterNote * 8 - $script:NoteBreak)
        Whole = [int] ($script:QuarterNote * 4 - $script:NoteBreak)
        DottedHalf = [int] ($script:QuarterNote * 3 - $script:NoteBreak)
        Half = [int] ($script:QuarterNote * 2 - $script:NoteBreak)
        DottedQuarter = [int] ($script:QuarterNote * 1.5 - $script:NoteBreak)
        Quarter = [int] ($script:QuarterNote - $script:NoteBreak)
        DottedEighth = [int] ($script:QuarterNote / 2.5 - $script:NoteBreak)
        Eighth = [int] ($script:QuarterNote / 2 - $script:NoteBreak)
        DottedSixteenth = [int] ($script:QuarterNote / 6 - $script:NoteBreak)
        Sixteenth = [int] ($script:QuarterNote / 4 - $script:NoteBreak)
        DottedThirtySecond = [int] ($script:QuarterNote / 10 - $script:NoteBreak)
        ThirtySecond = [int] ($script:QuarterNote / 8 - $script:NoteBreak)
    }
}

Function Set-Tempo {
    [CmdletBinding()] param (
        [int] $Bpm = 120
    )
    Write-Verbose "Old tempo was based on quarter note of $script:QuarterNote ms ($([int](60000 / $script:QuarterNote)) bpm)"
    $newQuarter = [int] (60000 / $Bpm)
    Write-Verbose "Old tempo is based on quarter note of $newQuarter ms ($Bpm bpm)"
    $script:QuarterNote = $newQuarter
    Update-Durations
}

Function Set-Staccato {
    [CmdletBinding()] param (
        [int] $MillisecondsPause = 0
    )
    Write-Verbose "Old staccato amount was pausing for $script:NoteBreak ms between notes"
    $script:NoteBreak = $MillisecondsPause
    Update-Durations
}

function Invoke-Tune {
    [CmdletBinding()] param (
        [string] $Path
    )
    $MyInvocation.PSScriptRoot
    $content = Get-Content $Path
    foreach($line in $content) {
        if([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $parts = $line.Split(" ")
        $action = $parts[0]
        $value = $parts[1]
        switch ($action.ToUpper()) {
            "SILENCE" {
                $duration = $script:Durations."$value"
                Write-Verbose "Silence for $value ($duration + $script:NoteBreak ms)"
                Start-Sleep -Milliseconds ($duration + $script:NoteBreak)
            }
            "TEMPO" {
                Write-Verbose "Tempo change to $value bpm"
                Set-Tempo -Bpm $value
            }
            "STACCATO" {
                Write-Verbose "Setting staccato (pause between notes) to $value ms"
                Set-Staccato -MillisecondsPause $value
            }
            default {
                $note = $script:Notes."$action"
                $duration = $script:Durations."$value"
                if($line.StartsWith("#")) {
                    Write-Verbose "Comment: $($parts -replace '^#\s+', '')"
                } else {
                    Write-Verbose "Note: $action ($note mhz), Duration: $value ($duration ms)"
                    [Console]::Beep($note, $duration)
                    if($script:NoteBreak -gt 0) {
                        Write-Verbose "Staccato pause for $($script:NoteBreak) ms"
                        Start-Sleep -Milliseconds $script:NoteBreak
                    }
                }
            }
        }
    }
}

Update-Durations

Export-ModuleMember -Function Invoke-Tune