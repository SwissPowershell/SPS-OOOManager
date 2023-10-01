class OOORuleSet {
    [DayOfWeek[]] ${WorkingDaysOfWeek} = @([DayOfWeek]::Monday,[DayOfWeek]::Tuesday,[DayOfWeek]::Wednesday,[DayOfWeek]::Thursday,[DayOfWeek]::Friday)
    [System.DateOnly[]] ${PublicHolidays} = @()
    [System.TimeOnly] ${DayStartTime} = [System.TimeOnly]::New(8,00)
    [System.TimeOnly] ${MidDayStartTime} = [System.TimeOnly]::New(12,00)
    [System.TimeOnly] ${MidDayStopTime} = [System.TimeOnly]::New(14,00)
    [System.TimeOnly] ${DayStopTime} = [System.TimeOnly]::New(18,00)
    OOORuleSet() {}
    OOORuleSet($PublicHolidays) {
        $This.PublicHolidays = $PublicHolidays
    }
    [Boolean] IsOff([DateTime] $DayToCheck) {
        # Reset the Hour Minute Second Millisecond to 0
        $RetVal = $False
        Switch ($DayToCheck) {
            {$_.DayOfWeek -notin $This.WorkingDaysOfWeek} {
                Write-Verbose "$([System.DateOnly]::FromDateTime($_).ToLongDateString()) is not a working day"
                $Retval = $True
                BREAK
            }
            {[System.DateOnly]::FromDateTime($_) -in $this.PublicHolidays} {
                Write-Verbose "$([System.DateOnly]::FromDateTime($_).ToLongDateString()) is part of official days off"
                $Retval = $True
                BREAK
            }
            Default {
                Write-Verbose "$([System.DateOnly]::FromDateTime($_).ToLongDateString()) is a working day and is not part of official days off"
                BREAK
            }
        }
        Return $RetVal
    }
}
Class PersonalHolyday {
    Hidden [DateTime] ${RequestedStart}
    Hidden [DateTime] ${RequestedStop}
    [DateTime] ${OutOfOfficeStart}
    [DateTime] ${OutOfOfficeStop}
    Hidden [OOORuleSet] ${OOORuleSet} = [OOORuleSet]::New()
    PersonalHolyday (){}
    PersonalHolyday ([DateTime] $Start,[DateTime] $Stop){
        $This.RequestedStart = $Start
        $This.RequestedStop = $Stop
        $This.GetOutOfOfficeSpan()
    }
    PersonalHolyday ([DateTime] $Start,[DateTime] $Stop,[OOORuleSet] $OOORuleSet){
        $This.RequestedStart = $Start
        $This.OutOfOfficeStart = $Start
        $This.RequestedStop = $Stop
        $This.OutOfOfficeStop = $Stop
        $This.OOORuleSet = $OOORuleSet
        $This.GetOutOfOfficeSpan()
    }
    [void] GetOutOfOfficeSpan() {
        Write-Verbose "Creating the out of office Span"
        # if the RequestedStart is not starting in middle of the day
        if ([System.TimeOnly]::FromDateTime($This.RequestedStart) -lt $This.OOORuleSet.MidDayStopTime) {
            Write-Verbose "The request start $($This.RequestedStart.Hour)h is before $($This.OOORuleSet.MidDayStopTime.Hour)h => checking days before"
            # Check the days before until I find a working day
            Do {
                $DateToTest = $This.OutOfOfficeStart.AddDays(-1)
                $IsOff = $this.OOORuleSet.IsOff($DateToTest)
                if ($IsOff -eq $True) {
                    # Write-Verbose "The $([System.DateOnly]::FromDateTime($DateToTest).ToLongDateString()) is a day off"
                    $This.OutOfOfficeStart = $DateToTest
                }
            } Until (-Not $IsOff)
        }Else{
            Write-Verbose "The request start $($This.RequestedStart.ToLongDateString()) is after $($This.OOORuleSet.MidDayStopTime.Hour)h out of office will start at that time"
        }
        
        # if the RequestedStop is not ending in middle of the day
        if ([System.TimeOnly]::FromDateTime($This.RequestedStop) -gt $This.OOORuleSet.MidDayStartTime) {
            Write-Verbose "The request Stop $($This.RequestedStop.Hour)h is after $($This.OOORuleSet.MidDayStartTime.Hour)h => checking days after"
            Do {
                $DateToTest = $This.OutOfOfficeStop.AddDays(1)
                $IsOff = $this.OOORuleSet.IsOff($DateToTest)
                if ($IsOff -eq $True) {
                    # Write-Verbose "The $($DateToTest.Day).$($DateToTest.Month).$($DateToTest.Year) is a day off"
                    $This.OutOfOfficeStop = $DateToTest
                }
            } Until (-Not $IsOff)
        }Else{
            Write-Verbose "The request start $($This.RequestedStart.ToLongDateString()) is before $($This.OOORuleSet.MidDayStartTime.Hour)h out of office will start at that time"
        }
    }
    [Boolean] IsOff($GivenObject) {
        $RetVal = $Null
        if ($GivenObject -is [DateTime]) {
            # Is just a date
            # Check if the given date is part of the holyday
            $RetVal = $False
            if ($GivenObject -le $this.OutOfOfficeStop) {
                if ($GivenObject -ge $this.OutOfOfficeStart) {
                    $RetVal = $True
                }
            }
        }ElseIf ($GivenObject -is [PersonalHolyday]) {
            # check if the given personal holyday is matching the existing (can either be within or during (after and before))
            if (($GivenObject.OutOfOfficeStart -ge $this.OutOfOfficeStart) -and ($GivenObject.OutOfOfficeStop -le $this.OutOfOfficeStop)) {
                # the given object is within this holyday
                $RetVal = $True
            }ElseIf (($GivenObject.OutOfOfficeStart -ge $this.OutOfOfficeStart) -and ($GivenObject.OutOfOfficeStart -le $this.OutOfOfficeStop)) {
                # The given object start during this holyday
                $RetVal = $True
            }ElseIf (($GivenObject.OutOfOfficeStop -le $this.OutOfOfficeStop) -and ($GivenObject.OutOfOfficeStop -ge $this.OutOfOfficeStop)) {
                # the given object stop during this holyday
                $RetVal = $True
            }
        }
        Return $RetVal
    }
    [void] Update([PersonalHolyday] $Update) {
        if (($Update.OutOfOfficeStart -ge $this.OutOfOfficeStart) -and ($Update.OutOfOfficeStop -le $this.OutOfOfficeStop)) {
            # is within nothing to do
        }ElseIf (($Update.OutOfOfficeStart -ge $this.OutOfOfficeStart) -and ($Update.OutOfOfficeStart -le $this.OutOfOfficeStop)) {
            # The given object start during this holyday but end after => Should update
            $This.RequestedStop = $Update.RequestedStop
            $This.OutOfOfficeStop = $Update.OutOfOfficeStop
            # $This.GetOutOfOfficeSpan()
        }ElseIf (($Update.OutOfOfficeStop -le $this.OutOfOfficeStop) -and ($Update.OutOfOfficeStop -ge $this.OutOfOfficeStop)) {
            # the given object stop during this holyday but start before => Should update
            $This.RequestedStart = $Update.RequestedStart
            $This.OutOfOfficeStart = $Update.OutOfOfficeStart
            # $This.GetOutOfOfficeSpan()
        }
    }
    
}
Class PersonalHolydays {
    [PersonalHolyday[]] ${PersonalHolyDays}
    [OOORuleSet] ${OOORuleSet} = [OOORuleSet]::New()
    PersonalHolydays(){}
    PersonalHolydays([OOORuleSet] $OOORuleSet){
        $This.OOORuleSet = $OOORuleSet
    }
    [Void] AddRequest([PersonalHolyday] $PersonalHolyDay) {
        $Updated = $False
        # Check if the add request is extending an existing request
        ForEach ($ExistingHolyday in $This.PersonalHolydays) {
            # if the Personal HolyDay Request is existing
            if ($ExistingHolyday.IsOff($PersonalHolyDay)) {
                Write-Verbose "Update existing holyday [$([System.DateOnly]::FromDateTime($ExistingHolyday.RequestedStart).ToLongDateString()) - $([System.DateOnly]::FromDateTime($ExistingHolyday.RequestedStop).ToLongDateString())]"
                Write-Verbose "`t With $([System.DateOnly]::FromDateTime($PersonalHolyDay.RequestedStart).ToLongDateString()) - $([System.DateOnly]::FromDateTime($PersonalHolyDay.RequestedStop).ToLongDateString())"
                $ExistingHolyday.Update($PersonalHolyDay)
                $Updated = $True
            }
        }
        if ($Updated -eq $False) {
            # No existing request updated => Add this request as is
            $This.PersonalHolyDays += $PersonalHolyDay
        }
    }
    [Void] RemoveRequest([PersonalHolyday] $RemoveHolyDay) {
        # Check if the remove request is part of an existing request
        $NewPersonalHolydays = @()
        ForEach ($ExistingHolyday in $This.PersonalHolydays) {
            # if the Personal HolyDay Request is existing
            if ($ExistingHolyday.IsOff($RemoveHolyDay)) {
                # of here the request should be split in two
                Write-Verbose "Removing [$([System.DateOnly]::FromDateTime($RemoveHolyDay.RequestedStart).ToLongDateString()) - $([System.DateOnly]::FromDateTime($RemoveHolyDay.RequestedStop).ToLongDateString())] from [$([System.DateOnly]::FromDateTime($ExistingHolyday.RequestedStart).ToLongDateString()) - $([System.DateOnly]::FromDateTime($ExistingHolyday.RequestedStop).ToLongDateString())]"
                $NewPersonalHolydays += [PersonalHolyday]::New($ExistingHolyday.RequestedStart,$RemoveHolyDay.RequestedStart,$This.OOORuleSet)
                $NewPersonalHolydays += [PersonalHolyday]::New($RemoveHolyDay.RequestedStop,$ExistingHolyday.RequestedStop,$This.OOORuleSet)
            }Else {
                $NewPersonalHolydays += $ExistingHolyday
            }
        }
        $This.PersonalHolyDays = $NewPersonalHolydays
    }
}

Function New-OOORuleSet {
    [CMDLetBinding()]
    Param(
        [Parameter(Mandatory = $False, Position = 0)]
        [DayOfWeek[]] ${WorkingDaysOfWeek},
        [Parameter(Mandatory = $False,Position = 1)]
        [System.DateOnly[]] ${PublicHolidays},
        [Parameter(Mandatory = $False,Position = 2)]
        [System.TimeOnly] ${DayStartTime},
        [Parameter(Mandatory = $False,Position = 3)]
        [System.TimeOnly] ${MidDayStartTime},
        [Parameter(Mandatory = $False,Position = 4)]
        [System.TimeOnly] ${MidDayStopTime},
        [Parameter(Mandatory = $False,Position = 5)]
        [System.TimeOnly] ${DayStopTime}
    )
    BEGIN {
        #region Function initialisation DO NOT REMOVE
        [String] ${FunctionName} = $MyInvocation.MyCommand
        [DateTime] ${FunctionEnterTime} = [DateTime]::Now
        Write-Verbose "Entering : $($FunctionName)"
        #endregion Function initialisation DO NOT REMOVE
    }
    PROCESS {
        #region Function Processing DO NOT REMOVE
        Write-Verbose "Processing : $($FunctionName)"
        #region Function Processing DO NOT REMOVE
        $OOORuleSet = [OOORuleSet]::New()
        if ($WorkingDaysOfWeek) {$OOORuleSet.WorkingDaysOfWeek = $WorkingDaysOfWeek}
        if ($PublicHolidays) {$OOORuleSet.PublicHolidays = $PublicHolidays}
        if ($DayStartTime) {$OOORuleSet.DayStartTime = $DayStartTime}
        if ($MidDayStartTime) {$OOORuleSet.MidDayStartTime = $MidDayStartTime}
        if ($MidDayStopTime) {$OOORuleSet.MidDayStopTime = $MidDayStopTime}
        if ($DayStopTime) {$OOORuleSet.DayStopTime = $DayStopTime}
    }
    END {
        $TimeSpentinFunc = New-TimeSpan -Start $FunctionEnterTime -Verbose:$False -ErrorAction SilentlyContinue;$TimeUnits = [ORDERED] @{TotalDays = "$($TimeSpentinFunc.TotalDays) D.";TotalHours = "$($TimeSpentinFunc.TotalHours) h.";TotalMinutes = "$($TimeSpentinFunc.TotalMinutes) min.";TotalSeconds = "$($TimeSpentinFunc.TotalSeconds) s.";TotalMilliseconds = "$($TimeSpentinFunc.TotalMilliseconds) ms."}
        ForEach ($Unit in $TimeUnits.GetEnumerator()) {if ($TimeSpentinFunc.($Unit.Key) -gt 1) {$TimeSpentString = $Unit.Value;break}};if (-not $TimeSpentString) {$TimeSpentString = "$($TimeSpentinFunc.Ticks) Ticks"}
        Write-Verbose "Ending : $($MyInvocation.MyCommand) - TimeSpent : $($TimeSpentString)"
        #endregion Function closing DO NOT REMOVE
        Return $OOORuleSet
    }
}
Function New-PersonalHolydayRequest {
    [CMDLetBinding()]
    Param(
        [Parameter(Mandatory = $True,Position = 0)]
        [DateTime] ${Start},
        [Parameter(Mandatory = $True,Position = 1)]
        [DateTime] ${Stop},
        [Parameter(Mandatory = $False,Position = 2)]
        [OOORuleSet] ${OOORuleSet} = [OOORuleSet]::New()
    )
    BEGIN {
        #region Function initialisation DO NOT REMOVE
        [String] ${FunctionName} = $MyInvocation.MyCommand
        [DateTime] ${FunctionEnterTime} = [DateTime]::Now
        Write-Verbose "Entering : $($FunctionName)"
        #endregion Function initialisation DO NOT REMOVE
        
    }
    PROCESS {
        #region Function Processing DO NOT REMOVE
        Write-Verbose "Processing : $($FunctionName)"
        #region Function Processing DO NOT REMOVE
        $RetVal = [PersonalHolyday]::New($Start,$Stop,$OOORuleSet)
    }
    END {
        $TimeSpentinFunc = New-TimeSpan -Start $FunctionEnterTime -Verbose:$False -ErrorAction SilentlyContinue;$TimeUnits = [ORDERED] @{TotalDays = "$($TimeSpentinFunc.TotalDays) D.";TotalHours = "$($TimeSpentinFunc.TotalHours) h.";TotalMinutes = "$($TimeSpentinFunc.TotalMinutes) min.";TotalSeconds = "$($TimeSpentinFunc.TotalSeconds) s.";TotalMilliseconds = "$($TimeSpentinFunc.TotalMilliseconds) ms."}
        ForEach ($Unit in $TimeUnits.GetEnumerator()) {if ($TimeSpentinFunc.($Unit.Key) -gt 1) {$TimeSpentString = $Unit.Value;break}};if (-not $TimeSpentString) {$TimeSpentString = "$($TimeSpentinFunc.Ticks) Ticks"}
        Write-Verbose "Ending : $($MyInvocation.MyCommand) - TimeSpent : $($TimeSpentString)"
        #endregion Function closing DO NOT REMOVE
        Return $RetVal
    }
}
Function New-PersonalHolydaysList {
    [CMDLetBinding()]
    Param(
        [Parameter(Mandatory = $False,Position = 0)]
        [OOORuleSet] ${OOORuleSet} = [OOORuleSet]::New()
    )
    BEGIN {
        #region Function initialisation DO NOT REMOVE
        [String] ${FunctionName} = $MyInvocation.MyCommand
        [DateTime] ${FunctionEnterTime} = [DateTime]::Now
        Write-Verbose "Entering : $($FunctionName)"
        #endregion Function initialisation DO NOT REMOVE
    }
    PROCESS {
        #region Function Processing DO NOT REMOVE
        Write-Verbose "Processing : $($FunctionName)"
        #region Function Processing DO NOT REMOVE
        $RetVal = [PersonalHolydays]::New($OOORuleSet)
    }
    END {
        $TimeSpentinFunc = New-TimeSpan -Start $FunctionEnterTime -Verbose:$False -ErrorAction SilentlyContinue;$TimeUnits = [ORDERED] @{TotalDays = "$($TimeSpentinFunc.TotalDays) D.";TotalHours = "$($TimeSpentinFunc.TotalHours) h.";TotalMinutes = "$($TimeSpentinFunc.TotalMinutes) min.";TotalSeconds = "$($TimeSpentinFunc.TotalSeconds) s.";TotalMilliseconds = "$($TimeSpentinFunc.TotalMilliseconds) ms."}
        ForEach ($Unit in $TimeUnits.GetEnumerator()) {if ($TimeSpentinFunc.($Unit.Key) -gt 1) {$TimeSpentString = $Unit.Value;break}};if (-not $TimeSpentString) {$TimeSpentString = "$($TimeSpentinFunc.Ticks) Ticks"}
        Write-Verbose "Ending : $($MyInvocation.MyCommand) - TimeSpent : $($TimeSpentString)"
        #endregion Function closing DO NOT REMOVE
        Return $RetVal
    }
    
}
