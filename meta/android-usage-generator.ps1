
[CmdletBinding()]
param (
    [string]$Target = "$PSScriptRoot/../index.md",
    [switch]$LastMonth,
    [switch]$JustPrint
)

begin {
    $ValidVersions = @{
        '15.0' = 35; '14.0' = 34; '13.0' = 33; '12.0' = 32; '11.0' = 30; '10.0' = 29; '9.0' = 28; '8.1' = 27
        '8.0' = 26; '7.1' = 25; '7.0' = 24; '6.0' = 23; '5.1' = 22; '5.0' = 21; '4.4' = 20; '4.3' = 18; '4.2' = 17
        '4.1' = 16; '4.0' = 15
    }
    $DataUrlTemplate = 'https://gs.statcounter.com/android-version-market-share/mobile-tablet/chart.php?' +
    'device=Mobile%20%26%20Tablet&device_hidden=mobile%2Btablet&multi-device=true&statType_hidden=' +
    'android_version&region_hidden=ww&granularity=monthly&statType=Android%20Version&region=Worldwide&' +
    'fromInt=201706&toInt={0}&fromMonthYear=2017-06&toMonthYear={1}&csv=1'
}

process {
    if (-not (Test-Path $Target -PathType Leaf)) {
        throw "Target file '$Target' not found!"
    }

    if ($LastMonth) { $AddM = -1 } else { $AddM = 0 }
    $CurMonth = [datetime]::Now.AddMonths($AddM).ToString('yyyy-MM')
    $CurMonthInt = $CurMonth -ireplace '-'
    $DataUrl = $DataUrlTemplate -f $CurMonthInt, $CurMonth

    Write-Host "Retrieving monthly usage stats from URL '$DataUrl'..."
    $Data = Invoke-RestMethod -Uri $DataUrl | ConvertFrom-Csv
    $Latest = $Data | Sort-Object -Property { [datetime]::Parse($_.Date) } -Descending | Select-Object -First 1

    Write-Host "Getting usage stats for month: $($Latest.Date)..."
    $Versions = $Latest.PSObject.Properties |
        Select-Object @{Name = 'Version'; Expression = {
                $_.Name -ireplace '^([\d.]+).*', '$1'
            }
        }, @{Name = 'Share'; Expression = { $_.Value -as [Double] } } |
        Where-Object Version -In $ValidVersions.Keys |
        Sort-Object -Property { $_.Version -as [Single] } -Descending

    Write-Host 'Summing cumulative usages...'
    $Cumulative = 0
    $LevelPcs = @{}
    foreach ($Field in $Versions) {
        $Share = $Field.Share
        $Cumulative += $Share
        $LevelPcs[$ValidVersions[$Field.Version]] = [Double]::Min($Cumulative, 100)
    }

    if ($JustPrint) {
        Write-Host "Printing cumulative usage for each API Level..."
        return [pscustomobject]$LevelPcs
    }

    Write-Host "Replacing cumulative usage stats in file '$Target'..."
    $PageContent = Get-Content -Path $Target -Raw
    foreach ($Level in $LevelPcs.Keys) {
        $Pattern = ('(?s)((?><td>\s*Level {0}.+?</td>)\s*(?><td>.+?</td>)\s*(?><td.+?</td>)?\s*' +
            '\{{% include progress-cell.html rowspan=\d+ percentage=)([\d\.]+)( %\}})') -f [regex]::Escape($Level)
        if ($PageContent -inotmatch $Pattern) {
            Write-Host -ForegroundColor Red "Couldn't find replacement region for API level $Level!"
            continue
        }
        $PcPretty = [math]::Round($LevelPcs[$Level], 1)
        Write-Host "    API Level ${Level}: Replacing $($Matches[2])% with $PcPretty%."
        $PageContent = $PageContent -ireplace $Pattern, ('${{1}}{0}$3' -f $PcPretty)
    }

    Write-Host "Replacing last updated footnote in file '$Target'..."
    $TodayPretty = Get-Date -Format 'MMMM dd, yyyy'
    $PageContent = $PageContent -ireplace
    '(Cumulative usage distribution figures were last updated on <b>).+(</b>)', ('$1{0}$2' -f $TodayPretty)

    Write-Host "Writing changes to '$Target'..."
    Set-Content -Path $Target -Value $PageContent -NoNewline
}
