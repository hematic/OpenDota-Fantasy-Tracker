class FantasyMatch {
    [String]$MatchID
    [String]$hero
    [Object]$Info
    [Double]$Kills
    [Double]$KillsPoints
    [Double]$Deaths
    [Double]$DeathsPoints
    [Double]$LastHits
    [Double]$LasthitPoints
    [Double]$Denies
    [Double]$DenyPoints
    [Double]$GPM
    [Double]$GPMPoints
    [Double]$TFPercentage
    [Double]$TFPoints
    [Int]$ObsWards
    [Double]$ObsPoints
    [Double]$CampsStacked
    [Double]$StackPoints
    [Double]$RunesGathered
    [Double]$RunesPoints
    [Double]$FirstBloodPoints
    [Double]$TowerPoints
    [Double]$RoshanPoints
    [Double]$SecondsofStun
    [Double]$StunPoints

    # Constructor
    FantasyMatch([Object]$Match,$AccountID) {
        $this._parseInfo($Match, $AccountID)
        $this.MatchID = $this.Info.'match_id'
        $this.Hero = $Global:Heroes | Where-Object {$_.id -eq $this.info.'hero_id'} | Select-Object -ExpandProperty 'Localized_name'
        $this._CalcKillPoints()
        $this._CalcDeathPoints()
        $this._CalcLastHitPoints()
        $this._CalcDenyPoints()
        $this._CalcGPMPoints()
        $this._CalcTeamFightPoints()
        $this._CalcObserverWardPoints()
        $this._CalcStackPoints()
        $this._CalcRunesPoints()
        $this._FirstBloodPoints()
        $this._CalcStunPoints()
        $this.TowerPoints = $this.Info.'tower_kills'
        $this.RoshanPoints = $this.Info.'roshan_kills'

    }

    # Method: Parse Player Data from the match
    hidden [void] _parseInfo($Match, $AccountID) {
        Try {
            $this.Info = $Match.players | Where-object {$_.'account_id' -eq $AccountID}
        }
        Catch [System.Exception] {
            Write-Error $_.Exception.Message
        }
    }

    #Calculate Kill Points
    hidden [void] _CalcKillPoints() {
        $this.Kills = $this.info.Kills
        [Double]$this.KillsPoints = $this.Kills * .3
    }

    #Calculate Death Points
    hidden [void] _CalcDeathPoints() {
        $this.Deaths = $this.info.Deaths
        [Double]$this.DeathsPoints = 3 - ($this.Deaths * .3)
    }

    #Calculate Last Hit Points
    hidden [void] _CalcLastHitPoints() {
        $this.LastHits = $this.info.'Last_Hits'
        [Double]$this.LasthitPoints = $this.Lasthits * .003
    }

    #Calculate Deny Points (not that they matter)
    hidden [void] _CalcDenyPoints() {
        $this.Denies = $this.info.Denies
        [Double]$this.DenyPoints = $this.Denies * .003
    }

    #Calculate GPM Points
    hidden [void] _CalcGPMPoints() {
        $this.GPM = $this.info.'gold_per_min'
        [Double]$this.GPMPoints = $this.GPM * .002
    }

    #Calculate TeamFight Points
    hidden [void] _CalcTeamFightPoints() {
        $this.TFPercentage = $this.info.'teamfight_participation'
        [Double]$this.TFPoints = $this.TFPercentage * 3
    }

    #Calculate Observer Ward Points
    hidden [void] _CalcObserverWardPoints() {
        [Int]$this.ObsWards = $this.info.'observer_uses'
        [Double]$this.ObsPoints = $this.ObsWards * .5
    }

    #Calculate Stacked Camps Points
    hidden [void] _CalcStackPoints() {
        $this.CampsStacked = $this.info.'camps_stacked'
        [Double]$this.StackPoints = $this.CampsStacked * .5
    }

    #Calculate Rune Points
    hidden [void] _CalcRunesPoints() {
        $this.RunesGathered = $this.info.'rune_pickups'
        [Double]$this.RunesPoints = $this.RunesGathered * .25
    }

    #Calculate First Blood Points
    hidden [void] _FirstBloodPoints() {
        [Double]$this.FirstBloodPoints = $this.info.'firstblood_claimed'  * 4
    }

    #Calculate Seconds of Stun Points
    hidden [void] _CalcStunPoints() {
        $this.SecondsofStun = $this.info.stuns
        [Double]$this.StunPoints = $this.SecondsofStun * .05
    }

}

Function New-OutputFile {
    Try {
        $Date = (Get-Date)
        $Month = (Get-Culture).DateTimeFormat.GetMonthName($Date.month)
        $Day = $Date.Day
        $year = $Date.Year
        $ReportFile = "c:\temp\Dota 2 Fantasy - $Month-$Day-$Year.xlsx"
        Copy-Item -Path 'C:\temp\Dota 2 Fantasy.xlsx' -Destination $ReportFile
        Write-Output $ReportFile
    }
    Catch {
        Write-Host "Unable to create Dota 2 Fantasy.xlsx from the template." -ForegroundColor Red
        Write-Error "Unable to create Dota 2 Fantasy.xlsx from the template."
    }
}

Function Set-Worksheet {
    [CmdletBinding()]
    Param(
        [ValidateSet('Phillip', 'Matt', 'Brad','Summary')]
        [String]$WorkSheet,
        $FilteredData,
        [String]$FilePath
    )

    #region Open the Excel Doc
        Try {
            $Excel = New-Object -ComObject Excel.Application
            $ExcelWorkBook = $Excel.Workbooks.Open($FilePath)
            $ExcelWorkSheet = $Excel.WorkSheets.item($WorkSheet)
            $ExcelWorkSheet.activate()
        }
        Catch {
            Write-Host $_.exception.message
            Write-Host "Unable to set worksheet : $Worksheet"
            Write-Error "Unable to set worksheet : $Worksheet"
            $ExcelWorkBook.Save()
            $ExcelWorkBook.Close()
            $Excel.Quit() 
            Return;
        }
    #endregion

    #Summary Worksheet
    If($Worksheet -eq 'Summary'){
        #region Set the header row
            Try {
                $ExcelWorkSheet.Cells.Item(1, 1) = 'Player'
                $ExcelWorkSheet.Cells.Item(1, 2) = 'Total Points'
                $ExcelWorkSheet.Cells.Item(1, 3) = 'Kill Points'
                $ExcelWorkSheet.Cells.Item(1, 4) = 'Death Points'
                $ExcelWorkSheet.Cells.Item(1, 5) = 'GPM Points'
                $ExcelWorkSheet.Cells.Item(1, 6) = 'Last Hit Points'
                $ExcelWorkSheet.Cells.Item(1, 7) = 'Deny Points'
                $ExcelWorkSheet.Cells.Item(1, 8) = 'Stack Points'
                $ExcelWorkSheet.Cells.Item(1, 9) = 'Rune Points'
                $ExcelWorkSheet.Cells.Item(1, 10) = 'Tower Points'
                $ExcelWorkSheet.Cells.Item(1, 11) = 'First Blood Points'
                $ExcelWorkSheet.Cells.Item(1, 12) = 'Roshan Points'
                $ExcelWorkSheet.Cells.Item(1, 13) = 'Observers Points'
                $ExcelWorkSheet.Cells.Item(1, 14) = 'Stuns Points'
                $ExcelWorkSheet.Cells.Item(1, 15) = 'Teamfight Points'
                $headerRange = $ExcelWorksheet.Range("a1", "o1")
                $headerRange.AutoFilter() | Out-Null
            }
            Catch {
                Write-Host $_.exception.message
                Write-Host "Unable to set headers for worksheet : $Worksheet"
                $ExcelWorkBook.Save()
                $ExcelWorkBook.Close()
                $Excel.Quit() 
                Return;
            }
        #endregion

        #region export the data

            $ExcelWorkSheet.Cells.Item(2, 1) = 'Brad'
            $ExcelWorkSheet.Cells.Item(2, 2) = "=Brad!C$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 3) = "=Brad!D$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 4) = "=Brad!E$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 5) = "=Brad!F$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 6) = "=Brad!G$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 7) = "=Brad!H$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 8) = "=Brad!I$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 9) = "=Brad!J$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 10) = "=Brad!K$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 11) = "=Brad!L$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 12) = "=Brad!M$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 13) = "=Brad!N$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 14) = "=Brad!O$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 15) = "=Brad!P$($Global:BradCount + 3)"


            $ExcelWorkSheet.Cells.Item(3, 1) = 'Phillip'
            $ExcelWorkSheet.Cells.Item(3, 2) = "=Phillip!C$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 3) = "=Phillip!D$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 4) = "=Phillip!E$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 5) = "=Phillip!F$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 6) = "=Phillip!G$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 7) = "=Phillip!H$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 8) = "=Phillip!I$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 9) = "=Phillip!J$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 10) = "=Phillip!K$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 11) = "=Phillip!L$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 12) = "=Phillip!M$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 13) = "=Phillip!N$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 14) = "=Phillip!O$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 15) = "=Phillip!P$($Global:PhillipCount + 3)"

            $ExcelWorkSheet.Cells.Item(4, 1) = 'Matt'
            $ExcelWorkSheet.Cells.Item(4, 2) = "=Matt!C$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 3) = "=Matt!D$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 4) = "=Matt!E$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 5) = "=Matt!F$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 6) = "=Matt!G$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 7) = "=Matt!H$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 8) = "=Matt!I$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 9) = "=Matt!J$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 10) = "=Matt!K$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 11) = "=Matt!L$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 12) = "=Matt!M$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 13) = "=Matt!N$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 14) = "=Matt!O$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 15) = "=Matt!P$($Global:MattCount + 3)"

        #endregion 
    }
    #User Worksheets
    Else{
        #region Set the header row
            Try {
                $ExcelWorkSheet.Cells.Item(1, 1) = 'Match ID'
                $ExcelWorkSheet.Cells.Item(1, 2) = 'Hero'
                $ExcelWorkSheet.Cells.Item(1, 3) = 'Total Points'
                $ExcelWorkSheet.Cells.Item(1, 4) = 'Kill Points'
                $ExcelWorkSheet.Cells.Item(1, 5) = 'Death Points'
                $ExcelWorkSheet.Cells.Item(1, 6) = 'GPM Points'
                $ExcelWorkSheet.Cells.Item(1, 7) = 'Last Hit Points'
                $ExcelWorkSheet.Cells.Item(1, 8) = 'Deny Points'
                $ExcelWorkSheet.Cells.Item(1, 9) = 'Stack Points'
                $ExcelWorkSheet.Cells.Item(1, 10) = 'Rune Points'
                $ExcelWorkSheet.Cells.Item(1, 11) = 'Tower Points'
                $ExcelWorkSheet.Cells.Item(1, 12) = 'First Blood Points'
                $ExcelWorkSheet.Cells.Item(1, 13) = 'Roshan Points'
                $ExcelWorkSheet.Cells.Item(1, 14) = 'Observers Points'
                $ExcelWorkSheet.Cells.Item(1, 15) = 'Stuns Points'
                $ExcelWorkSheet.Cells.Item(1, 16) = 'Teamfight Points'
                $headerRange = $ExcelWorksheet.Range("a1", "p1")
                $headerRange.AutoFilter() | Out-Null
            }
            Catch {
                Write-Host $_.exception.message
                Write-Host "Unable to set headers for worksheet : $Worksheet"
                $ExcelWorkBook.Save()
                $ExcelWorkBook.Close()
                $Excel.Quit() 
                Return;
            }
        #endregion

        #region export the data
            for ($i = 0; $i -lt $FilteredData.Count; $i++) {
                $Row = $I + 2
                $FilteredData | OUT-FILE C:\TEMP\WTF.TXT
                Try {
                    $ExcelWorkSheet.Cells.Item($Row, 1) = $FilteredData[$I].MatchID
                    $ExcelWorkSheet.Cells.Item($Row, 2) = $FilteredData[$I].Hero
                    $ExcelWorkSheet.Cells.Item($Row, 3) = "=SUM(D$($Row):P$($Row)"
                    $ExcelWorkSheet.Cells.Item($Row, 4) = $FilteredData[$I].KillsPoints
                    $ExcelWorkSheet.Cells.Item($Row, 5) = $FilteredData[$I].DeathsPoints
                    $ExcelWorkSheet.Cells.Item($Row, 6) = $FilteredData[$I].GPMPoints
                    $ExcelWorkSheet.Cells.Item($Row, 7) = $FilteredData[$I].LastHitPoints
                    $ExcelWorkSheet.Cells.Item($Row, 8) = $FilteredData[$I].DenyPoints
                    $ExcelWorkSheet.Cells.Item($Row, 9) = $FilteredData[$I].StackPoints
                    $ExcelWorkSheet.Cells.Item($Row, 10) = $FilteredData[$I].RunesPoints
                    $ExcelWorkSheet.Cells.Item($Row, 11) = $FilteredData[$I].TowerPoints
                    $ExcelWorkSheet.Cells.Item($Row, 12) = $FilteredData[$I].FirstBloodPoints
                    $ExcelWorkSheet.Cells.Item($Row, 13) = $FilteredData[$I].RoshanPoints
                    $ExcelWorkSheet.Cells.Item($Row, 14) = $FilteredData[$I].ObsPoints
                    $ExcelWorkSheet.Cells.Item($Row, 15) = $FilteredData[$I].StunPoints
                    $ExcelWorkSheet.Cells.Item($Row, 16) = $FilteredData[$I].TFPoints
                }
                Catch {
                    Write-Host $_.exception.message
                }
            }
        #endregion

        #region Make totals row
            $Row = $FilteredData.count + 3
            $ExcelWorkSheet.Cells.Item($Row, 1) = 'Grand Totals'
            $ExcelWorkSheet.Cells.Item($Row, 3) = "=SUM(C2:C$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 4) = "=SUM(D2:D$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 5) = "=SUM(E2:E$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 6) = "=SUM(F2:F$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 7) = "=SUM(G2:G$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 8) = "=SUM(H2:H$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 9) = "=SUM(I2:I$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 10) = "=SUM(J2:J$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 11) = "=SUM(k2:k$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 12) = "=SUM(l2:l$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 13) = "=SUM(m2:m$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 14) = "=SUM(n2:n$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 15) = "=SUM(o2:o$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 16) = "=SUM(p2:p$($FilteredData.count + 1)"
        #endregion
    }
      
    #region Close out excel.
        $ExcelWorksheet.Columns.AutoFit()
        $ExcelWorkBook.Save()
        $ExcelWorkBook.Close()
        $Excel.Quit()
    #endregion
}

Function Get-Heroes {

    $URI = 'https://api.opendota.com/api/heroes'
    $Result = invoke-restmethod -Uri $Uri
    Write-Output $Result

}

Function Get-Matches {
    Param(
        [String]$AccountID,
        [Int]$DaysBack = 30
    )

    $URI = "https://api.opendota.com/api/players/$AccountID/matches?date=$DaysBack"
    $Result = invoke-restmethod -Uri $Uri
    Write-Output $Result
}

Function Get-MatchData {
    Param(
        [String]$MatchID
    )

    $URI = "https://api.opendota.com/api/matches/$MatchID"
    $Result = invoke-restmethod -Uri $Uri
    Write-Output $Result
}

Function Get-FantasyTotals {
    Param(
        $D
    )
    [Double]$Total = $D.StunPoints + $D.RunesPoints + $D.RoshanPoints + $D.LastHitPoints + $D.GPMPoints + $D.StackPoints + $D.TFPoints + $D.ObsPoints + $D.FirstBloodPoints + $D.TowerPoints + $D.KillsPoints + $D.DeathsPoints + $D.DenyPoints 
    Write-Output $Total
}

#region verify that template.xlsx is present.
	Try{
		$ReportFile = New-OutputFile
	}
	Catch{
		Write-Error $_
		Exit;
	}
#endregion

#Get Heroes
$Global:Heroes = Get-Heroes

#AccountIDs
$PhillipAccountID = '7057906'
$BradAccountID = '25287058'
$MattAccountID = '71462475'

#Get Matches to cross-reference
[array]$PhillipRecentMatches = Get-Matches -AccountID $PhillipAccountID
[array]$BradRecentMatches    = Get-Matches -AccountID $BradAccountID
[array]$MattRecentMatches    = Get-Matches -AccountID $MattAccountID

#Define Array Lists
[Array]$PhillipMatches = @()
[Array]$BradMatches = @()
[Array]$MattMatches = @()

Foreach ($Match in $PhillipRecentMatches) {
    If ($BradRecentMatches.'match_id' -contains $match.'match_id' -and $MattRecentMatches.'match_id' -contains $match.'match_id') {
        Write-Host "Gathering Match Data for Match : $($Match.'match_id')"
        $Data = Get-MatchData -MatchID $Match.'match_id'
        $PhillipFantasy = New-Object -TypeName FantasyMatch -ArgumentList $Data,$PhillipAccountID
        $PhillipMatches += $PhillipFantasy
        $BradFantasy = New-Object -TypeName FantasyMatch -ArgumentList $Data,$BradAccountID
        $BradMatches += $BradFantasy
        $MattFantasy = New-Object -TypeName FantasyMatch -ArgumentList $Data,$MattAccountID
        $MattMatches += $MattFantasy
    }
}

$Global:PhillipCount = $PhillipMatches.count
$Global:MattCount = $MattMatches.count
$Global:BradCount = $BradMatches.count

Set-Worksheet -WorkSheet 'Phillip' -FilteredData $PhillipMatches -Filepath $ReportFile
Set-Worksheet -WorkSheet 'Matt' -FilteredData $MattMatches -Filepath $ReportFile
Set-Worksheet -WorkSheet 'Brad' -FilteredData $BradMatches -Filepath $ReportFile
Set-Worksheet -WorkSheet 'Summary' -Filepath $ReportFile
