
Class O3Helper { # Helper methods used by the Out of office module
    static [DateTime] ToDateOnly([DateTime] $GivenDate) { # This will normalize the given Date in order to be able to compare hours without date
        Return $(Get-Date -Date $GivenDate -Hour 0 -Minute 0 -Second 0 -Millisecond 0)
    }
    static [DateTime] ToTimeOnly([DateTime] $GivenDate) { # This will normalize the given Date in order to be able to compare Date without time
        Return $(Get-Date -Date $GivenDate -Year 1 -Day 1 -Month 1)
    }
    static [DateTime] EndOfDay([DateTime] $GivenDate) {
        $EndOfCurrentDay = $True # Put true if you want 23h59.59.999 or false if you want next day 00h00.00.000
        $ReturnDate = $GivenDate

        if ($EndOfCurrentDay -eq $True) {
            $EndHour = 23
            $EndMinute = 59
            $EndSecond = 59
            $EndMillisecond = 999
            $AddDay = 0 # There will be 0 day added to the current date
        }Else{
            $EndHour = 0
            $EndMinute = 0
            $EndSecond = 0
            $EndMillisecond = 0
            if (($GivenDate.Hour -eq $EndHour) -and ($GivenDate.Minute -eq $EndMinute) -and ($GivenDate.Second -eq $EndSecond) -and ($GivenDate.Millisecond -eq $EndMillisecond)) { # Prevent adding a day to an allready normalized Date
                $AddDay = 0
            }Else{
                $AddDay = 1 # There will be 1 day added to the current date
            }
        }
        $ReturnDate = $(Get-Date -Date $($GivenDate.AddDays($AddDay)) -Hour $EndHour -Minute $EndMinute -Second $EndSecond -Millisecond $EndMillisecond)
        Return $ReturnDate
    }
}
Class O3RuleSet { # Will hold global Out of office Informations
    [UInt32] ${EmployeeID}                                                                      # Employee ID
    [DayOfWeek[]] ${WeekEndDays} =              @([DayOfWeek]::Saturday,[DayOfWeek]::Sunday)    # The day for which the user did not work # Ignored if WeekEndInHolydays
    [Collections.ArrayList] ${PublicHolyDays} = [Collections.ArrayList]::New()                  # A list of Holyday Date # Can also contains WeekEndDays
    [DateTime] ${DayStartTime} =                [DateTime]::Parse('0001-1-1T07:00:00')             # The default beginning of the day # Note See if it's needed
    [DateTime] ${MidDayStartTime} =             [DateTime]::Parse('0001-1-1T12:00:00')             # The default beginning of the Mid Day break
    [DateTime] ${MidDayStopTime} =              [DateTime]::Parse('0001-1-1T14:00:00')             # The default end of the Mid Day break
    [DateTime] ${DayStopTime} =                 [DateTime]::Parse('0001-1-1T18:00:00')             # The default end of Day # Note See if it's needed
    [Boolean] ${WeekEndInHolydays} =            $True                                           # When true the WeekEndDays properties will be ignored
    O3RuleSet() {}
    O3RuleSet([Collections.ArrayList] $PublicHolyDays) {
        # Create the Default RuleSet with given Public Holydays
        $This.PublicHolyDays = $PublicHolyDays
        $This.NormalizeDate()
    }
    [Boolean] IsOff([DateTime] ${GivenDate}) {
        $IsOff = $False
        # Normalize Given Date
        $DateToCheck = [O3Helper]::ToDateOnly($GivenDate)
        Switch ($DateToCheck) {
            {($This.WeekEndInHolydays -eq $False) -and ($_.DayOfWeek -in $WeekEndDays)} {
                # WeekEndInHolyDays is false and Given Date WeekDay is part of WeekEndDays
                $IsOff = $True
                BREAK
            }
            {$_ -in $PublicHolyDays} {
                # The Given Date is part of PublicHolydays
                $IsOff = $True
                BREAK
            }
            Default {
                BREAK
            }
        }
        Return $IsOff
    }
    Hidden [Void] NormalizeDate() {
        $This.PublicHolyDays = $This.PublicHolyDays | ForEach-Object {[O3Helper]::ToDateOnly($_)}
    }
}
Class O3HolyDayRequest { # Single HolyDayRequest
    [UInt32] ${EmployeeID}              # Employee ID
    [DateTime] ${StartDate}             # Request Start as requested by user
    [DateTime] ${StopDate}              # Request Stop as requested by user
    [DateTime] ${O3StartDate}           # Out of office Start after calculation
    [DateTime] ${O3StopDate}            # Out of office Stop after calculation
    [O3RuleSet] ${O3RuleSet}            # O3 Rule set applied for calculation
    [Boolean] ${IsRemoveRequest}        # True if the request is a remove request (this will set O3Start and O3Stop to StartDate and StopDate)
    O3HolyDayRequest() {
        Throw 'Not authorized constructor use [O3HolyDayRequest]::new($Start,$Stop) , [O3HolyDayRequest]::new($Start,$Stop,$RuleSet), [O3HolyDayRequest]::New($EmployeeID,$Start,$Stop,$RuleSet)'
    }
    O3HolyDayRequest ([DateTime] $Start,[DateTime] $Stop) { # Create a new O3HolyDayRequest with default O3RuleSet and employee ID 0
        $This.EmployeeID = 0
        $This.StartDate = $Start
        $This.StopDate = $Stop
        $This.O3StartDate = [O3Helper]::ToDateOnly($Start)
        $This.O3StopDate = [O3Helper]::EndOfDay($Stop)
        $This.O3RuleSet = [O3RuleSet]::New()
        $This.Calculate()
    }
    O3HolyDayRequest ([DateTime] $Start,[DateTime] $Stop,[O3RuleSet] $O3RuleSet) { # Create a new O3HolyDayRequest with employee ID 0
        $This.EmployeeID = 0
        $This.StartDate = $Start
        $This.StopDate = $Stop
        $This.O3StartDate = [O3Helper]::ToDateOnly($Start)
        $This.O3StopDate = [O3Helper]::EndOfDay($Stop)
        $This.O3RuleSet = $O3RuleSet
        $This.Calculate()
    }
    O3HolyDayRequest ([UINT32] $EmployeeID,[DateTime] $Start,[DateTime] $Stop,[O3RuleSet] $O3RuleSet) { # Create a new O3HolyDayRequest with given datas
        $This.EmployeeID = $EmployeeID
        $This.StartDate = $Start
        $This.StopDate = $Stop
        $This.O3StartDate = [O3Helper]::ToDateOnly($Start)
        $This.O3StopDate = [O3Helper]::EndOfDay($Stop)
        $This.O3RuleSet = $O3RuleSet
        $This.Calculate()
    }
    Hidden Calculate() { # Calculate O3StartDate O3Stop and O3Dates
        if ($This.StartDate.Hour -lt $This.O3RuleSet.MidDayStopTime.Hour) { # if the request did not start in the middle of the day (usually a midday start request is 14h)
            $DayToTest = $This.StartDate.AddDays(-1)
            $DayIsOff = $False
            Do {
                # if the previous day is off
                $DayIsOff = $This.O3RuleSet.IsOff($DayToTest)
                if ($DayIsOff -eq $True) {
                    # Store this date as day to start
                    $This.O3StartDate = [O3Helper]::ToDateOnly($DayToTest)
                    # Test the previous day
                    $DayToTest = $DayToTest.AddDays(-1)
                }
            } Until ($DayIsOff -eq $False)
        }Else{ # it start in the middle of the day
            $This.O3StartDate = $This.StartDate
        }
        if ($This.StopDate.Hour -gt $This.O3RuleSet.MidDayStartTime.Hour) { # if the request did not stop in the middle of the day (usually a midday stop request is 12h)
            $DayToTest = $This.StopDate.AddDays(1)
            $DayIsOff = $False
            Do {
                # if the previous day is off
                $DayIsOff = $This.O3RuleSet.IsOff($DayToTest)
                if ($DayIsOff -eq $True) {
                    # Store this date as day to start
                    $This.O3StopDate = [O3Helper]::ToDateOnly($DayToTest)
                    # Test the previous day
                    $DayToTest = $DayToTest.AddDays(1)
                }
            } Until ($DayIsOff -eq $False)
            # Normalize the end of day (let you decide between 23h59 or 0h the next day by changing the helper class EndOfDay static method)
            $This.O3StopDate = [O3Helper]::EndOfDay($This.O3StopDate)
        }Else{ # it stop in the middle of the day
            $This.O3StopDate = $This.StopDate
        }
    }
    [Boolean] IsIn($GivenObject) { # return true if the given object is between O3Start and O3Stop (for datetime) or is in (for O3HolyDayRequest)
        $IsInResult = $False
        if ($GivenObject -is [DateTime]) { # Test if the given date object is in the O3HolyDayRequest
            if (($GivenObject -ge $This.O3StartDate) -and ($GivenObject -le $This.O3StopDate)) { # it's beetween O3start and O3stop
                $IsInResult = $True
            }
        }Elseif ($GivenObject -is [O3HolyDayRequest]){ # Test if the given O3HolyDayRequest is in the current O3HolyDayRequest
            if (($GivenObject.O3StartDate -ge $This.O3StartDate) -and ($GivenObject.O3StopDate -le $This.O3StopDate)) {
                $IsInResult = $True
            }
        }Else{
            Throw "$($GivenObject.GetType().FullName) not allowed, You can only get IsIn for [DateTime] or [O3HolyDayRequest]"
        }
        Return $IsInResult
    }
    [Boolean] Overlapse([O3HolyDayRequest] $O3HolyDayRequest) { # return true if the O3HolyDayRequest overlapse the current O3HolyDayRequest
        $OverlapseResult = $False
        # if it start before but end during return true
        if (($O3HolyDayRequest.O3StartDate -le $This.O3StartDate) -and ($O3HolyDayRequest.O3StopDate -le $This.O3StopDate) -and ($O3HolyDayRequest.O3StopDate -ge $This.O3StartDate)) {
            # It start before start, it stop before before stop and it stop after start
            $OverlapseResult = $True
        }ElseIf (($O3HolyDayRequest.O3StartDate -ge $This.O3StartDate) -and ($O3HolyDayRequest.O3StopDate -ge $This.O3StopDate) -and ($O3HolyDayRequest.O3StartDate -le $This.O3StopDate)) {# if it start during but end after return true
            # it Start after start, it stop after stop and it start before stop
            $OverlapseResult = $True
        }ElseIf (($O3HolyDayRequest.O3StartDate -le $This.O3StartDate) -and ($O3HolyDayRequest.O3StopDate -ge $This.O3StopDate)) {
            # it start before and it stop after
            $OverLapseResult = $True
        }
        # else return false
        Return $OverlapseResult
    }
    [Void] Update([O3HolyDayRequest] $O3HolyDayRequest) {
        if ($This.IsIn($O3HolyDayRequest) -eq $True) { # The given O3Request is allready handled by this request => Do Nothing
            # Do Nothing
        }Elseif ($This.Overlapse($O3HolyDayRequest) -eq $True) { # The given O3Request overlapse the existing => Update
            if ($O3HolyDayRequest.O3StartDate -le $This.O3StartDate) { # it start before update startdate
                $This.StartDate = $O3HolyDayRequest.StartDate
                $This.O3StartDate = $O3HolyDayRequest.O3StartDate
            }
            if ($O3HolyDayRequest.O3StopDate -ge $This.O3StopDate) { # it end after update stopdate
                $This.StopDate = $O3HolyDayRequest.StopDate
                $This.O3StopDate = $O3HolyDayRequest.O3StopDate
            }
        }Else{ # The given O3Request as nothing to do with the existing => do nothing
            # Do Nothing
        }
    }
    [Void] SetRemove() {
        $This.IsRemoveRequest = $true
        # Change the way the data are displayed inverting start and stop it will help for removal
        if ($This.StartDate.Hour -ge $This.O3RuleSet.MidDayStartTime.hour) { # if the start is in the middle of the day (14h) => it will become a stop at 12h and an O3Stop at the same time
            $NewStop = Get-Date $This.StartDate -Hour $This.O3RuleSet.MidDayStartTime.Hour -Minute $This.O3RuleSet.MidDayStartTime.Minute -Second $This.O3RuleSet.MidDayStartTime.Second -Millisecond $This.O3RuleSet.MidDayStartTime.Millisecond
            $NewO3Stop = $NewStop
        }Else{ # if the start is in the morning => it will become a stop at 18h the day before and an O3Stop at end of day
            $NewStop = Get-Date $This.StartDate.AddDays(-1) -Hour $This.O3RuleSet.DayStopTime.Hour -Minute $This.O3RuleSet.DayStopTime.Minute -Second $This.O3RuleSet.DayStopTime.Second -Millisecond $This.O3RuleSet.DayStopTime.Millisecond
            $NewO3Stop = [O3Helper]::EndOfDay($NewStop)
        }

        if ($This.StopDate.Hour -le $This.O3RuleSet.MidDayStopTime.hour) { # if the stop is in the middle of the day (12h) => it will become a start at 14h and an O3Start at the same time
            $NewStart = Get-Date $This.StopDate -Hour $This.O3RuleSet.MidDayStopTime.Hour -Minute $This.O3RuleSet.MidDayStopTime.Minute -Second $This.O3RuleSet.MidDayStopTime.Second -Millisecond $This.O3RuleSet.MidDayStopTime.Millisecond
            $NewO3Start = $NewStart
        }Else{ # if the stop is in the end of the day => it will become a start at 8h the day after and an O3Start in the beginning of the day
            $NewStart = Get-Date $This.StopDate.AddDays(1) -Hour $This.O3RuleSet.DayStartTime.Hour -Minute $This.O3RuleSet.DayStartTime.Minute -Second $This.O3RuleSet.DayStartTime.Second -Millisecond $This.O3RuleSet.DayStartTime.Millisecond
            $NewO3Start = [O3Helper]::ToDateOnly($NewStart)
        }
        $This.StartDate = $NewStart
        $This.O3StartDate = $NewO3Start
        $This.StopDate = $NewStop
        $This.O3StopDate = $NewO3Stop
    }
    [String] ToString() {
        Return "$($This.StartDate.ToString()) => $($This.StopDate.ToString()) ($($This.O3StartDate.ToString()) => $($This.O3StopDate.ToString()))"
    }
}
Class O3 {
    [UInt32] ${EmployeeID}                          # Employee ID
    [Collections.ArrayList] ${O3HolyDayRequests}    # All the Holyday Request made by the user
    [O3RuleSet] ${O3RuleSet}                        # O3 Rule set applied
    O3() {
        $This.EmployeeID = 0
        $This.O3RuleSet = [O3RuleSet]::new()
        $This.O3HolyDayRequests = @()
    }
    O3([UInt32] $EmployeeId) {
        $This.EmployeeID = $EmployeeId
        $This.O3RuleSet = [O3RuleSet]::new()
        $This.O3HolyDayRequests = @()
    }
    O3([UInt32] $EmployeeId, [O3RuleSet] $O3RuleSet) {
        $This.EmployeeID = $EmployeeID
        $This.O3RuleSet = $O3RuleSet
        $This.O3HolyDayRequests = @()
    }
    O3([UInt32] $EmployeeId, [O3RuleSet] $O3RuleSet,[Collections.ArrayList] $O3HolyDayRequests) {
        $This.EmployeeID = $EmployeeID
        $This.O3RuleSet = $O3RuleSet
        $This.O3HolyDayRequests = $O3HolyDayRequests
    }
    [Boolean] IsOff([DateTime] $GivenDate) {
        $IsOffResult = $False
        ForEach ($O3Request in $This.O3HolyDayRequests) {
            if ($O3Request.IsIn($GivenDate)) {
                $IsOffResult = $True
                Break
            }
        }
        Return $IsOffResult
    }
    [Void] AddRequest([DateTime] $Start, [DateTime] $Stop) { # Add a HolyDay Request by giving start stop
        $This.AddRequest([O3HolyDayRequest]::new($This.EmployeeID,$Start,$Stop,$This.O3RuleSet))
    }
    [Void] AddRequest([O3HolyDayRequest] $O3HolyDayRequest) { # Add a HolyDay Request by giving an O3HolyDayRequest
        # Check if employeeID match
        if ($O3HolyDayRequest.EmployeeID -ne $This.EmployeeID) { # Employee ID did not match
            Throw "HolyDayRequest EmployeeID did not match O3 EmployeeID"
        }Else{ # Employee ID match continue adding
            # Search if a request match the new request
            # to do if the update match more tha one request
            $Updated = $False
            ForEach ($O3Request in $This.O3HolyDayRequests) {
                if (($O3Request.IsIn($O3HolyDayRequest)) -or ($O3Request.Overlapse($O3HolyDayRequest))) {
                    $O3Request.Update($O3HolyDayRequest)
                    $Updated = $True
                    Break
                }
            }
            if ($Updated -eq $False) {
                $This.O3HolyDayRequests.Add($O3HolyDayRequest) | out-null
            }
        }
    }
    [Void] RemoveRequest([DateTime] $Start, [DateTime] ${Stop}) {
        $RemoveRequest = [O3HolyDayRequest]::new($This.EmployeeID,$Start,$Stop,$This.O3RuleSet)
        $RemoveRequest.SetRemove()
        $This.RemoveRequest($RemoveRequest)
    }
    [Void] RemoveRequest([O3HolyDayRequest] $O3HolyDayRemoveRequest) {
        if ($O3HolyDayRemoveRequest.IsRemoveRequest -eq $False) {
            Throw "Input is not a remove request please apply SetRemove() to the object"
        }Else{
            $Updated = $False
            [Collections.ArrayList] $NewO3HolyDayRequests = [Collections.ArrayList]::New()
            # Find the in the O3HolyDayRequests table the entry that match this remove request
            ForEach ($O3Request in $This.O3HolyDayRequests) {
                if (($O3Request.IsIn($O3HolyDayRemoveRequest)) -or ($O3Request.Overlapse($O3HolyDayRemoveRequest))) {
                    if (($O3Request.StartDate -ne $O3HolyDayRemoveRequest.StopDate) -and ($O3Request.StartDate -lt $O3HolyDayRemoveRequest.StopDate)) {
                        $UpdatedRequest1 = [O3HolyDayRequest]::New($O3Request.EmployeeID,$O3Request.StartDate,$O3HolyDayRemoveRequest.StopDate,$O3Request.O3RuleSet)
                        $NewO3HolyDayRequests.Add($UpdatedRequest1) | out-null
                    }
                    if (($O3HolyDayRemoveRequest.StartDate -ne $O3Request.StopDate) -and ($O3HolyDayRemoveRequest.StartDate -lt $O3Request.StopDate)) {
                        $UpdatedRequest2 = [O3HolyDayRequest]::New($O3Request.EmployeeID,$O3HolyDayRemoveRequest.StartDate,$O3Request.StopDate,$O3Request.O3RuleSet)
                        $NewO3HolyDayRequests.Add($UpdatedRequest2) | out-null
                    }
                    $Updated = $True
                }Else{
                    $NewO3HolyDayRequests.Add($O3Request) | out-null
                }
            }
            if ($Updated -eq $True) {
                $This.O3HolyDayRequests = $NewO3HolyDayRequests
                Write-Verbose 'Successfully updated (remove)'
            }Else{
                Write-Verbose 'Nothing to remove'
            }
        }
    }

}

Function New-O3RuleSet {
    [CMDLetBinding()]
    Param(
        [Parameter(Mandatory = $False, Position = 0)]
        [DayOfWeek[]] ${WeekEndDays},
        [Parameter(Mandatory = $False,Position = 1)]
        [DateTime[]] ${PublicHolydays},
        [Parameter(Mandatory = $False,Position = 2)]
        [DateTime] ${DayStartTime},
        [Parameter(Mandatory = $False,Position = 3)]
        [DateTime] ${MidDayStartTime},
        [Parameter(Mandatory = $False,Position = 4)]
        [DateTime] ${MidDayStopTime},
        [Parameter(Mandatory = $False,Position = 5)]
        [DateTime] ${DayStopTime},
        [Parameter(Mandatory = $False,Position = 6)]
        [Boolean] ${WeekEndInHolydays}
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
        $O3RuleSet = [O3RuleSet]::New()
        if ($WorkingDaysOfWeek) {$O3RuleSet.WorkingDaysOfWeek = $WorkingDaysOfWeek}
        if ($PublicHolydays) {$O3RuleSet.PublicHolydays = $PublicHolydays}
        if ($DayStartTime) {$O3RuleSet.DayStartTime = $DayStartTime}
        if ($MidDayStartTime) {$O3RuleSet.MidDayStartTime = $MidDayStartTime}
        if ($MidDayStopTime) {$O3RuleSet.MidDayStopTime = $MidDayStopTime}
        if ($DayStopTime) {$O3RuleSet.DayStopTime = $DayStopTime}
        if ($WeekEndInHolydays) {$O3RuleSet.WeekEndInHolydays = $WeekEndInHolydays}
        $O3RuleSet.NormalizeDate()
    }
    END {
        $TimeSpentinFunc = New-TimeSpan -Start $FunctionEnterTime -Verbose:$False -ErrorAction SilentlyContinue;$TimeUnits = [ORDERED] @{TotalDays = "$($TimeSpentinFunc.TotalDays) D.";TotalHours = "$($TimeSpentinFunc.TotalHours) h.";TotalMinutes = "$($TimeSpentinFunc.TotalMinutes) min.";TotalSeconds = "$($TimeSpentinFunc.TotalSeconds) s.";TotalMilliseconds = "$($TimeSpentinFunc.TotalMilliseconds) ms."}
        ForEach ($Unit in $TimeUnits.GetEnumerator()) {if ($TimeSpentinFunc.($Unit.Key) -gt 1) {$TimeSpentString = $Unit.Value;break}};if (-not $TimeSpentString) {$TimeSpentString = "$($TimeSpentinFunc.Ticks) Ticks"}
        Write-Verbose "Ending : $($MyInvocation.MyCommand) - TimeSpent : $($TimeSpentString)"
        #endregion Function closing DO NOT REMOVE
        Return $O3RuleSet
    }
}
Function New-O3Request {
    [CMDLetBinding()]
    Param(
        [Parameter(Mandatory = $True,Position = 0)]
        [DateTime] ${Start},
        [Parameter(Mandatory = $True,Position = 1)]
        [DateTime] ${Stop},
        [Parameter(Mandatory = $False,Position = 2)]
        [O3RuleSet] ${O3RuleSet} = [O3RuleSet]::New(),
        [Switch] ${Remove}
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
        $RetVal = [O3HolyDayRequest]::New($Start,$Stop,$O3RuleSet)
        if ($Remove -eq $True) {
            $RetVal.SetRemove()
        }
    }
    END {
        $TimeSpentinFunc = New-TimeSpan -Start $FunctionEnterTime -Verbose:$False -ErrorAction SilentlyContinue;$TimeUnits = [ORDERED] @{TotalDays = "$($TimeSpentinFunc.TotalDays) D.";TotalHours = "$($TimeSpentinFunc.TotalHours) h.";TotalMinutes = "$($TimeSpentinFunc.TotalMinutes) min.";TotalSeconds = "$($TimeSpentinFunc.TotalSeconds) s.";TotalMilliseconds = "$($TimeSpentinFunc.TotalMilliseconds) ms."}
        ForEach ($Unit in $TimeUnits.GetEnumerator()) {if ($TimeSpentinFunc.($Unit.Key) -gt 1) {$TimeSpentString = $Unit.Value;break}};if (-not $TimeSpentString) {$TimeSpentString = "$($TimeSpentinFunc.Ticks) Ticks"}
        Write-Verbose "Ending : $($MyInvocation.MyCommand) - TimeSpent : $($TimeSpentString)"
        #endregion Function closing DO NOT REMOVE
        Return $RetVal
    }
}
Function New-O3List {
    [CMDLetBinding()]
    Param(
        [Parameter(Mandatory = $False, Position = 0)]
        [Uint32] ${EmployeeId} = 0,
        [Parameter(Mandatory = $False,Position = 1)]
        [O3RuleSet] ${O3RuleSet} = [O3RuleSet]::New()
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
        $RetVal = [O3]::New($EmployeeId,$O3RuleSet)
    }
    END {
        $TimeSpentinFunc = New-TimeSpan -Start $FunctionEnterTime -Verbose:$False -ErrorAction SilentlyContinue;$TimeUnits = [ORDERED] @{TotalDays = "$($TimeSpentinFunc.TotalDays) D.";TotalHours = "$($TimeSpentinFunc.TotalHours) h.";TotalMinutes = "$($TimeSpentinFunc.TotalMinutes) min.";TotalSeconds = "$($TimeSpentinFunc.TotalSeconds) s.";TotalMilliseconds = "$($TimeSpentinFunc.TotalMilliseconds) ms."}
        ForEach ($Unit in $TimeUnits.GetEnumerator()) {if ($TimeSpentinFunc.($Unit.Key) -gt 1) {$TimeSpentString = $Unit.Value;break}};if (-not $TimeSpentString) {$TimeSpentString = "$($TimeSpentinFunc.Ticks) Ticks"}
        Write-Verbose "Ending : $($MyInvocation.MyCommand) - TimeSpent : $($TimeSpentString)"
        #endregion Function closing DO NOT REMOVE
        Return $RetVal
    }

}
Function Write-O3Calendar {
    Param(
        [Int] $StartYear = 2023,
        [Int] $StopYear = $StartYear,
        [ConsoleColor] $WeekEndColor = 'Yellow',
        [ConsoleColor] $WeekDayColor = 'Green',
        [ConsoleColor] $BusyDaycolor = 'Red',
        [O3] $PersonalHolyDays,
        [O3RuleSet] $OOORuleSet
    )
    Function Write-Month {
        Param(
            $Year = 2023,
            $Month = 2,
            [ConsoleColor] $WeekEndColor = [ConsoleColor]::Yellow,
            [ConsoleColor] $WeekDayColor = [ConsoleColor]::Green,
            [ConsoleColor] $BusyDayColor = [ConsoleColor]::Red,
            [O3] $PersonalHolyDays,
            [O3RuleSet] $OOORuleSet
        )
        Function Write-WeekDays {
            Param(
                [ConsoleColor] $WeekEndColor = [ConsoleColor]::Yellow,
                [ConsoleColor] $WeekDayColor = [ConsoleColor]::Green
            )
            $WeekDayNames = [Enum]::GetNames("DayOfWeek")
            $WeedEndDay = @([DayOfWeek]::Saturday,[DayOfWeek]::Sunday)
            $MaxLength = $WeekDayNames | ForEach-Object {$_.Length} | sort-Object -Descending | Select-Object -First 1
            $DefaultSize = $MaxLength + 3
            ForEach ($DayName in $WeekDayNames) {
                if ($DayName -in $WeedEndDay) {
                    $ForegroundColor = $WeekEndColor
                }Else{
                    $ForegroundColor = $WeekDayColor
                }
                $String = "{0,-$($DefaultSize)}" -f $DayName
                Write-Host $String -NoNewLine -ForegroundColor $ForegroundColor
            }
            Write-Host ''
            ForEach ($DayName in $WeekDayNames) {
                $String = "{0,-$($DefaultSize)}" -f $("-" * $DefaultSize)
                Write-Host $String -NoNewLine -ForegroundColor DarkGray
            }
            Write-Host ''
        }
        # Get the first and last weekday
        $FirstWeekDay = [Enum]::GetNames("DayOfWeek") | Select-Object -first 1
        $LastWeekDay = [Enum]::GetNames("DayOfWeek") | Select-Object -Last 1
        $WeekDayNames = [Enum]::GetNames("DayOfWeek")
        $WeedEndDay = @([DayOfWeek]::Saturday,[DayOfWeek]::Sunday)
        $MaxLength = $WeekDayNames | ForEach-Object {$_.Length} | sort-Object -Descending | Select-Object -First 1
        $DefaultSize = $MaxLength + 3
        # Get First day of week from month
        $FirstMonthDay = Get-Date -Year $Year -Month $Month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        $ExitLoop = $False
        Do {
            if ($FirstMonthDay.DayOfWeek -eq $FirstWeekDay) {
                $ExitLoop = $True
            }Else{
                $FirstMonthDay = $FirstMonthDay.AddDays(-1)
            }
        }Until ($ExitLoop)
        $DaysInMonth = [DateTime]::DaysInMonth($Year, $Month)
        $LastMonthDay = Get-Date -Year $Year -Month $Month -Day $DaysInMonth -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        $ExitLoop = $False
        Do {
            if ($LastMonthDay.DayOfWeek -eq $LastWeekDay) {
                $ExitLoop = $True
            }Else{
                $LastMonthDay = $LastMonthDay.AddDays(1)
            }
        }Until ($ExitLoop)

        # Display Month Name
        $MonthName = (Get-UICulture).DateTimeFormat.GetMonthName($Month)
        Write-Host $MonthName.ToUpper() -ForegroundColor $WeekDayColor -NoNewline
        Write-Host " $($Year)" -ForegroundColor $WeekDayColor
        Write-Host $("-" * ($MonthName.Length + 1 + $Year.toString().Length)) -ForegroundColor $WeekDayColor
        # Display Week Days
        Write-WeekDays -WeekEndColor $WeekEndColor -WeekDayColor $WeekDayColor | out-null

        # Display Numbers
        $CurrDay = $FirstMonthDay
        $String = $CurrDay.Day
        Do {
            if ($OOORuleSet.IsOff($CurrDay)) {
                $ForegroundColor = $WeekEndColor
            }Else{
                $ForegroundColor = $WeekDayColor
            }
            # to be Enhanced by showing half of the day (test 10h and 16h)
            if ($PersonalHolyDays.IsOff($CurrDay)) {
                $ForegroundColor = $BusyDayColor
            }
            $Morning = Get-Date -Date $CurrDay -Hour 10 -Minute 0 -Second 0 -Millisecond 0
            $AfterNoon = Get-Date -Date $CurrDay -Hour 16 -Minute 0 -Second 0 -Millisecond 0
            $MorningIsOff = $PersonalHolyDays.IsOff($Morning)
            $AfterNoonIsOff = $PersonalHolyDays.IsOff($AfterNoon)
            if ($MorningIsOff -ne $AfterNoonIsOff) {
                if ($MorningIsOff -eq $True) {
                    $String = "$($String)-"
                    $ForegroundColor = $BusyDayColor
                }Elseif ($AfterNoonIsOff -eq $True) {
                    $String = "$($String)+"
                    $ForegroundColor = $BusyDayColor
                }
            }

            if ($CurrDay.Month -ne $Month) {
                $Color = "Dark$($ForegroundColor)"
                # Remove the overlapse for better reading
                $String = ''
            }Else{
                $Color = $ForegroundColor

            }
            $String = "{0,-$($DefaultSize)}" -f $String
            Write-Host $String -ForegroundColor $Color -NoNewLine
            if ($CurrDay.DayOfWeek -eq $LastWeekDay) {
                Write-Host ''
            }
            $CurrDay = $CurrDay.AddDays(1)
            $String = $CurrDay.Day
        }until($CurrDay -gt $LastMonthDay)
        Write-Host ''
    }
    Write-Host "Tips ! " -ForegroundColor 'Blue'
    Write-Host "`t Worked days" -ForegroundColor $WeekDayColor
    Write-Host "`t Official days off" -ForegroundColor $WeekEndColor
    Write-Host "`t Days off" -ForegroundColor $BusyDaycolor
    Write-Host "`t`t Morning only (-) " -ForegroundColor $BusyDaycolor
    Write-Host "`t`t Afternoon only (+) " -ForegroundColor $BusyDaycolor
    Write-Host ''
    For($y = $StartYear;$y -le $StopYear;$y++) {
        For($m = 1;$m -le 12;$m++) {
            Write-Month -Year $y -Month $m -WeekEndColor $WeekEndColor -WeekDayColor $WeekDayColor -BusyDayColor $BusyDayColor -PersonalHolyDays $PersonalHolyDays -OOORuleSet $OOORuleSet
        }
    }
}
