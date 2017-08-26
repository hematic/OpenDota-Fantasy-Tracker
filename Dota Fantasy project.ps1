class FantasyMatch {
    #region Property Declarations
    [String]$MatchID
    [String]$hero
    [Object]$Info
    [Int]$DurationSeconds
    [Int]$Kills
    [Int]$Assists
    [Int]$Deaths
    [Int]$LastHits
    [Int]$Denies
    [Int]$ObsWards
    [Int]$CampsStacked
    [Int]$RunesGathered
    [Int]$KillStreakPoints
    [Int]$NetWorth
    [Int]$CourierKills
    [Int]$BarracksDamage
    [Int]$Tier1Damage
    [Int]$Tier2Damage
    [Int]$Tier3Damage
    [Double]$CourierKillsPoints
    [Double]$DurationMinutesRounded
    [Double]$KillsPoints
    [Double]$AssistsPoints
    [Double]$DeathsPoints
    [Double]$LasthitPoints
    [Double]$DenyPoints
    [Double]$GPM
    [Double]$GPMPoints
    [Double]$TFPercentage
    [Double]$TFPoints
    [Double]$ObsPoints
    [Double]$StackPoints
    [Double]$RunesPoints
    [Double]$FirstBloodPoints
    [Double]$TowerPoints
    [Double]$BldgDamagePoints
    [Double]$RoshanPoints
    [Double]$SecondsofStun
    [Double]$StunPoints
    [Double]$DewardPoints
    [Double]$TotalPoints
    [Double]$PointsPerMinute
    #endregion

    #region Constructor
    FantasyMatch([Object]$Match,$AccountID) {
        $this._parseInfo($Match, $AccountID)
        $this._ParseMatchID()
        $this._ParseHero()
        $this._CalcDuration()
        $this._CalcKillPoints()
        $this._CalcAssistPoints()
        $this._CalcDeathPoints()
        $this._CalcLastHitPoints()
        $this._CalcDenyPoints()
        $this._CalcGPMPoints()
        $this._CalcTeamFightPoints()
        $this._CalcObserverWardPoints()
        $this._CalcStackPoints()
        $this._CalcRunesPoints()
        $this._CalcFirstBloodPoints()
        $this._CalcCourierPoints()
        $this._CalcStunPoints()
        $this._CalcDewardPoints()
        $this._CalcHighestKillStreakPoints()
        $this._CalcBuildingPoints()
        $this._CalcRoshanPoints()
        $this._CalcTotalPoints()
        $this._CalcPointsPerMinute()

    }
    #endregion

    # Method: Parse Player Data from the match
    hidden [void] _parseInfo($Match, $AccountID) {
        Try {
            $this.Info = $Match.players | Where-object {$_.'account_id' -eq $AccountID}
        }
        Catch [System.Exception] {
            Write-Error $_.Exception.Message
        }
    }

    #Determine Hero
    hidden [void] _ParseHero() {
        $this.Hero = $Global:Heroes | Where-Object {$_.id -eq $this.info.'hero_id'} | Select-Object -ExpandProperty 'Localized_name'
    }

    #Determine Match ID
    hidden [void] _ParseMatchID() {
        $this.MatchID = $this.Info.'match_id'
    }

    #Calculate Duration
    hidden [void] _CalcDuration() {
        $this.DurationSeconds = $this.info.Duration
        $this.DurationMinutesRounded = [Math]::Round($this.DurationSeconds / 60, 0)
    }

    #Calculate Kill Points
    hidden [void] _CalcKillPoints() {
        $this.Kills = $this.info.Kills
        [Double]$this.KillsPoints = [Math]::Round($this.Kills * .5, 2)
    }

    #Calculate Assist Points
    hidden [void] _CalcAssistPoints() {
        $this.Assists = $this.info.Assists
        [Double]$this.AssistsPoints = [Math]::Round($this.Assists * .125, 2)
    }

    #Calculate Death Points
    hidden [void] _CalcDeathPoints() {
        $this.Deaths = $this.info.Deaths
        [Double]$this.DeathsPoints = 3 - [Math]::Round($this.Deaths * .3, 2)
    }

    #Calculate Last Hit Points
    hidden [void] _CalcLastHitPoints() {
        $this.LastHits = $this.info.'Last_Hits'
        [Double]$this.LasthitPoints = [Math]::Round($this.Lasthits * .004, 2)
    }

    #Calculate Deny Points
    hidden [void] _CalcDenyPoints() {
        $this.Denies = $this.info.Denies
        [Double]$this.DenyPoints = [Math]::Round($this.Denies * .025, 2)
    }

    #Calculate GPM Points
    hidden [void] _CalcGPMPoints() {
        $this.GPM = $this.info.'gold_per_min'
        [Double]$this.GPMPoints =[Math]::Round($this.GPM * .004, 2)
    }

    #Calculate TeamFight Points
    hidden [void] _CalcTeamFightPoints() {
        If(-not $this.info.'teamfight_participation'){
            $this.TFPoints = 0
            $this.TFPercentage
        }
        Else{
            $this.TFPercentage = $this.info.'teamfight_participation'
            [Double]$this.TFPoints  = [Math]::Round(([math]::log10($this.TFPercentage * 100) - [math]::log10(30)) * 3 / (2 - [math]::log10(30)), 2)
        }
    }

    #Calculate Observer Ward Points
    hidden [void] _CalcObserverWardPoints() {
        [Int]$this.ObsWards = $this.info.'observer_uses'
        [Double]$this.ObsPoints = [Math]::Round($this.ObsWards * .5, 2)
    }

    #Calculate Stacked Camps Points
    hidden [void] _CalcStackPoints() {
        $this.CampsStacked = $this.info.'camps_stacked'
        [Double]$this.StackPoints = [Math]::Round($this.CampsStacked * .5, 2)
    }

    #Calculate Rune Points
    hidden [void] _CalcRunesPoints() {
        $this.RunesGathered = $this.info.'rune_pickups'
        [Double]$this.RunesPoints = [Math]::Round($this.RunesGathered * .125, 2)
    }

    #Calculate Tower Points
    hidden [void] _CalcBuildingPoints() {
        $this.TowerPoints = $this.Info.'tower_kills' / 2
        $this.Tier1Damage = $this.Info.damage.'npc_dota_badguys_tower1_bot' + $this.Info.'damage.npc_dota_badguys_tower1_mid' + $this.Info.damage.'npc_dota_badguys_tower1_top'
        $this.Tier2Damage = $this.Info.damage.'npc_dota_badguys_tower2_bot' + $this.Info.'damage.npc_dota_badguys_tower2_mid' + $this.Info.damage.'npc_dota_badguys_tower2_top'
        $this.Tier3Damage = $this.Info.damage.'npc_dota_badguys_tower3_bot' + $this.Info.'damage.npc_dota_badguys_tower3_mid' + $this.Info.damage.'npc_dota_badguys_tower3_top'
        $this.BarracksDamage = $this.Info.damage.'npc_dota_badguys_melee_rax_bot' + $this.Info.damage.'npc_dota_badguys_range_rax_bot' + $this.Info.damage.'npc_dota_badguys_melee_rax_mid' + $this.Info.damage.'npc_dota_badguys_range_rax_mid' + $this.Info.damage.'npc_dota_badguys_melee_rax_top' + $this.Info.damage.'npc_dota_badguys_range_rax_top'
        [Double]$this.BldgDamagePoints = [Math]::Round(($this.Tier1Damage + $this.Tier2Damage + $this.Tier3Damage + $this.BarracksDamage) / 750, 2)
    }

    #Calculate First Blood Points
    hidden [void] _CalcFirstBloodPoints() {
        [Double]$this.FirstBloodPoints = [Math]::Round($this.info.'firstblood_claimed' * 4, 2)
    }

    #Calculate Courier Kills Points
    hidden [void] _CalcCourierPoints() {
        [Double]$this.CourierKillsPoints = [Math]::Round($this.info.'courier_kills' * 1, 2)
    }

    #Calculate Roshan Kills Points
    hidden [void] _CalcRoshanPoints() {
        $this.RoshanPoints = $this.Info.'roshan_kills'
    }

    #Calculate Seconds of Stun Points
    hidden [void] _CalcStunPoints() {
        $this.SecondsofStun = $this.info.stuns
        [Double]$this.StunPoints = [Math]::Round($this.SecondsofStun * .045, 2)
    }

    #Calculate Deward Points
    hidden [void] _CalcDewardPoints() {
        $this.DewardPoints = [Math]::Round(($this.info.'observer_kills' + $this.info.'sentry_kills') * .75, 2)
    }

    #Calculate Highest Kill STreak
    hidden [void] _CalcHighestKillStreakPoints() {
        $Streaks = @()
        If(-not $this.info.'kill_streaks'){
            $HighestKillStreak = 0
        }
        Else{
        ($this.info.'kill_streaks' | Get-Member -MemberType NoteProperty).Name | %{$Streaks += [Int]$_}
        $HighestKillStreak = $Streaks | Sort-Object -Descending | select -First 1

            Switch ($HighestKillStreak) {

                {$_ -lt 3}  {$this.KillStreakPoints = 0}
                {$_ -eq 3}  {$this.KillStreakPoints = 1}    
                {$_ -eq 4}  {$this.KillStreakPoints = 2} 
                {$_ -eq 5}  {$this.KillStreakPoints = 2.5} 
                {$_ -eq 6}  {$this.KillStreakPoints = 3}
                {$_ -eq 7}  {$this.KillStreakPoints = 3.5}
                {$_ -eq 8}  {$this.KillStreakPoints = 4} 
                {$_ -eq 9}  {$this.KillStreakPoints = 4.5} 
                {$_ -eq 10} {$this.KillStreakPoints = 5}
                {$_ -eq 11} {$this.KillStreakPoints = 5.5} 
                {$_ -eq 12} {$this.KillStreakPoints = 6}
                {$_ -eq 13} {$this.KillStreakPoints = 6.5} 
                {$_ -eq 14} {$this.KillStreakPoints = 7} 
                {$_ -ge 15} {$this.KillStreakPoints = 10} 
            }
        }
    }

    #Calculate Total Fantasy Points
    hidden [void] _CalcTotalPoints() {
        [Double]$this.TotalPoints = $this.KillsPoints + $this.DeathsPoints + $this.LasthitPoints + $this.DenyPoints + $this.GPMPoints + $this.TFPoints + $this.ObsPoints + $this.StackPoints + $this.RunesPoints + $this.FirstBloodPoints + $this.TowerPoints + $this.RoshanPoints + $this.StunPoints + $this.DewardPoints + $this.KillStreakPoints + $this.AssistsPoints + $this.CourierKillsPoints + $this.BldgDamagePoints

    }

    #Calculate Points Per Minute
    hidden [void] _CalcPointsPerMinute() {
        [Double]$this.PointsPerMinute = [Math]::Round(($this.TotalPoints / $this.DurationSeconds) * 60, 2)
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
        [String]$WorkSheet,
        [Array]$FilteredData,
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
                #General Info
                $ExcelWorkSheet.Cells.Item(1, 1)  = 'Player'
                $ExcelWorkSheet.Cells.Item(1, 2)  = 'Total'
                #Fighting
                $ExcelWorkSheet.Cells.Item(1, 3)  = 'Kills'
                $ExcelWorkSheet.Cells.Item(1, 4)  = 'Kill Streaks'
                $ExcelWorkSheet.Cells.Item(1, 5)  = 'Assists'
                $ExcelWorkSheet.Cells.Item(1, 6)  = 'Deaths'
                $ExcelWorkSheet.Cells.Item(1, 7)  = 'Stuns'
                $ExcelWorkSheet.Cells.Item(1, 8)  = 'Teamfights'
                #Farming
                $ExcelWorkSheet.Cells.Item(1, 9)  = 'GPM'
                $ExcelWorkSheet.Cells.Item(1, 10) = 'Last Hits'
                $ExcelWorkSheet.Cells.Item(1, 11) = 'Denies'
                $ExcelWorkSheet.Cells.Item(1, 12) = 'Runes'
                #Objectives
                $ExcelWorkSheet.Cells.Item(1, 13) = 'Towers'
                $ExcelWorkSheet.cells.Item(1, 14) = 'Bldg Damage'
                $ExcelWorkSheet.Cells.Item(1, 15) = 'First Blood'
                $ExcelWorkSheet.Cells.Item(1, 16) = 'Roshan'
                $ExcelWorkSheet.Cells.Item(1, 17) = 'Courier'
                #Supporting
                $ExcelWorkSheet.Cells.Item(1, 18) = 'Observers'
                $ExcelWorkSheet.Cells.Item(1, 19) = 'Stacks'
                $ExcelWorkSheet.Cells.Item(1, 20) = 'Dewards'
  
                $headerRange = $ExcelWorksheet.Range("A1", "T1")
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
            #Brad Data
            $ExcelWorkSheet.Cells.Item(2, 1)  = 'Brad'
            $ExcelWorkSheet.Cells.Item(2, 2)  = "=Brad!C$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 3)  = "=Brad!E$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 4)  = "=Brad!F$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 5)  = "=Brad!G$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 6)  = "=Brad!H$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 7)  = "=Brad!I$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 8)  = "=Brad!J$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 9)  = "=Brad!K$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 10) = "=Brad!L$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 11) = "=Brad!M$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 12) = "=Brad!N$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 13) = "=Brad!O$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 14) = "=Brad!P$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 15) = "=Brad!Q$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 16) = "=Brad!R$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 17) = "=Brad!S$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 18) = "=Brad!T$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 19) = "=Brad!U$($Global:BradCount + 3)"
            $ExcelWorkSheet.Cells.Item(2, 20) = "=Brad!V$($Global:BradCount + 3)"

            #Phillip Data
            $ExcelWorkSheet.Cells.Item(3, 1)  = 'Phillip'
            $ExcelWorkSheet.Cells.Item(3, 2)  = "=Phillip!C$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 3)  = "=Phillip!E$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 4)  = "=Phillip!F$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 5)  = "=Phillip!G$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 6)  = "=Phillip!H$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 7)  = "=Phillip!I$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 8)  = "=Phillip!J$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 9)  = "=Phillip!K$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 10) = "=Phillip!L$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 11) = "=Phillip!M$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 12) = "=Phillip!N$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 13) = "=Phillip!O$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 14) = "=Phillip!P$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 15) = "=Phillip!Q$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 16) = "=Phillip!R$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 17) = "=Phillip!S$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 18) = "=Phillip!T$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 19) = "=Phillip!U$($Global:PhillipCount + 3)"
            $ExcelWorkSheet.Cells.Item(3, 20) = "=Phillip!V$($Global:PhillipCount + 3)"
            
            #Matt Data
            $ExcelWorkSheet.Cells.Item(4, 1)  = 'Matt'
            $ExcelWorkSheet.Cells.Item(4, 2)  = "=Matt!C$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 3)  = "=Matt!E$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 4)  = "=Matt!F$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 5)  = "=Matt!G$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 6)  = "=Matt!H$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 7)  = "=Matt!I$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 8)  = "=Matt!J$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 9)  = "=Matt!K$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 10) = "=Matt!L$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 11) = "=Matt!M$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 12) = "=Matt!N$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 13) = "=Matt!O$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 14) = "=Matt!P$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 15) = "=Matt!Q$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 16) = "=Matt!R$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 17) = "=Matt!S$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 18) = "=Matt!T$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 19) = "=Matt!U$($Global:MattCount + 3)"
            $ExcelWorkSheet.Cells.Item(4, 20) = "=Matt!V$($Global:MattCount + 3)"
            
        #endregion 
    }

    #User Worksheets
    Else{
        #region Set the header row
            Try {
                #General Info
                $ExcelWorkSheet.Cells.Item(1, 1)  = 'Hero'
                $ExcelWorkSheet.Cells.Item(1, 2)  = 'Total'
                $ExcelWorkSheet.Cells.Item(1, 3)  = 'Per Minute'
                $ExcelWorkSheet.Cells.Item(1, 4)  = 'Duration'
                #Fighting
                $ExcelWorkSheet.Cells.Item(1, 5)  = 'Kills'
                $ExcelWorkSheet.Cells.Item(1, 6)  = 'Kill Streaks'
                $ExcelWorkSheet.Cells.Item(1, 7)  = 'Assists'
                $ExcelWorkSheet.Cells.Item(1, 8)  = 'Deaths'
                $ExcelWorkSheet.Cells.Item(1, 9)  = 'Stuns'
                $ExcelWorkSheet.Cells.Item(1, 10) = 'Teamfights'
                #Farming
                $ExcelWorkSheet.Cells.Item(1, 11) = 'GPM'
                $ExcelWorkSheet.Cells.Item(1, 12) = 'Last Hits'
                $ExcelWorkSheet.Cells.Item(1, 13) = 'Denies'
                $ExcelWorkSheet.Cells.Item(1, 14) = 'Runes'
                #Objectives
                $ExcelWorkSheet.Cells.Item(1, 15) = 'Towers'
                $ExcelWorkSheet.cells.Item(1, 16) = 'Bldg Damage'
                $ExcelWorkSheet.Cells.Item(1, 17) = 'First Blood'
                $ExcelWorkSheet.Cells.Item(1, 18) = 'Roshan'
                $ExcelWorkSheet.Cells.Item(1, 19) = 'Couriers'
                #Supporting
                $ExcelWorkSheet.Cells.Item(1, 20) = 'Observers'
                $ExcelWorkSheet.Cells.Item(1, 21) = 'Stacks'
                $ExcelWorkSheet.Cells.Item(1, 22) = 'Dewards'
                #ID
                $ExcelWorkSheet.Cells.Item(1, 23) = 'Match ID'
                $headerRange = $ExcelWorksheet.Range("A1", "W1")
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
                Try {
                    $ExcelWorkSheet.Cells.Item($Row, 1)  = $FilteredData[$I].Hero
                    $ExcelWorkSheet.Cells.Item($Row, 2)  = $FilteredData[$I].TotalPoints
                    $ExcelWorkSheet.Cells.Item($Row, 3)  = $FilteredData[$I].PointsPerMinute
                    $ExcelWorkSheet.Cells.Item($Row, 4)  = $FilteredData[$I].DurationMinutesRounded
                    $ExcelWorkSheet.Cells.Item($Row, 5)  = $FilteredData[$I].KillsPoints
                    $ExcelWorkSheet.Cells.Item($Row, 6)  = $FilteredData[$I].KillStreakPoints
                    $ExcelWorkSheet.Cells.Item($Row, 7)  = $FilteredData[$I].AssistsPoints
                    $ExcelWorkSheet.Cells.Item($Row, 8)  = $FilteredData[$I].DeathsPoints
                    $ExcelWorkSheet.Cells.Item($Row, 9)  = $FilteredData[$I].StunPoints
                    $ExcelWorkSheet.Cells.Item($Row, 10) = $FilteredData[$I].TFPoints
                    $ExcelWorkSheet.Cells.Item($Row, 11) = $FilteredData[$I].GPMPoints
                    $ExcelWorkSheet.Cells.Item($Row, 12) = $FilteredData[$I].LastHitPoints
                    $ExcelWorkSheet.Cells.Item($Row, 13) = $FilteredData[$I].DenyPoints
                    $ExcelWorkSheet.Cells.Item($Row, 14) = $FilteredData[$I].RunesPoints
                    $ExcelWorkSheet.Cells.Item($Row, 15) = $FilteredData[$I].TowerPoints
                    $ExcelWorkSheet.Cells.Item($Row, 16) = $FilteredData[$I].BldgDamagePoints
                    $ExcelWorkSheet.Cells.Item($Row, 17) = $FilteredData[$I].FirstBloodPoints
                    $ExcelWorkSheet.Cells.Item($Row, 18) = $FilteredData[$I].RoshanPoints
                    $ExcelWorkSheet.Cells.Item($Row, 19) = $FilteredData[$I].CourierKillsPoints
                    $ExcelWorkSheet.Cells.Item($Row, 20) = $FilteredData[$I].ObsPoints
                    $ExcelWorkSheet.Cells.Item($Row, 21) = $FilteredData[$I].StackPoints
                    $ExcelWorkSheet.Cells.Item($Row, 22) = $FilteredData[$I].DewardPoints
                    $ExcelWorkSheet.Cells.Item($Row, 23) = $FilteredData[$I].MatchID
                }
                Catch {
                    Write-Host $_.exception.message
                }
            }
        #endregion

        #region Make totals row
            $Row = $FilteredData.count + 3
            $ExcelWorkSheet.Cells.Item($Row, 1)  = 'Grand Totals'
            $ExcelWorkSheet.Cells.Item($Row, 3)  = "=SUM(B2:B$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 5)  = "=SUM(E2:E$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 6)  = "=SUM(F2:F$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 7)  = "=SUM(G2:G$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 8)  = "=SUM(H2:H$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 9)  = "=SUM(I2:I$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 10) = "=SUM(J2:J$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 11) = "=SUM(K2:K$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 12) = "=SUM(L2:L$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 13) = "=SUM(M2:M$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 14) = "=SUM(N2:N$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 15) = "=SUM(O2:O$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 16) = "=SUM(P2:P$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 17) = "=SUM(Q2:Q$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 18) = "=SUM(R2:R$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 19) = "=SUM(S2:S$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 20) = "=SUM(T2:T$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 21) = "=SUM(U2:U$($FilteredData.count + 1)"
            $ExcelWorkSheet.Cells.Item($Row, 22) = "=SUM(V2:V$($FilteredData.count + 1)"
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

#region verify that template.xlsx is present.
	Try{
		$ReportFile = New-OutputFile
	}
	Catch{
		Write-Error $_
		Exit;
	}
#endregion

#region Get Heroes
$Global:Heroes = Get-Heroes
#endregion

#region AccountIDs
$PhillipAccountID = '7057906'
$BradAccountID = '25287058'
$MattAccountID = '71462475'
#endregion

#region Get Matches to cross-reference
[array]$PhillipRecentMatches = Get-Matches -AccountID $PhillipAccountID -DaysBack 3
[array]$BradRecentMatches    = Get-Matches -AccountID $BradAccountID -DaysBack 3
[array]$MattRecentMatches    = Get-Matches -AccountID $MattAccountID -DaysBack 3
#endregion

#region Define Array Lists
[Array]$PhillipMatches = @()
[Array]$BradMatches = @()
[Array]$MattMatches = @()
#endregion

#region Create fantasy objects
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
#endregion

#region Define Global Count objects
$Global:PhillipCount = $PhillipMatches.count
$Global:MattCount = $MattMatches.count
$Global:BradCount = $BradMatches.count
#endregion

#region Export Data to the XLSX
Set-Worksheet -WorkSheet 'Phillip' -FilteredData $PhillipMatches -Filepath $ReportFile
Set-Worksheet -WorkSheet 'Matt' -FilteredData $MattMatches -Filepath $ReportFile
Set-Worksheet -WorkSheet 'Brad' -FilteredData $BradMatches -Filepath $ReportFile
Set-Worksheet -WorkSheet 'Summary' -Filepath $ReportFile
#endregion
