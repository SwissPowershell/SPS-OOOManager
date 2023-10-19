# SPS-OOOManager
Tentative to create an out of office manager (readme to be updated to reflect new names)
## New-OOORuleSet
Create a new out of office rule set object
- by default the rule set has no holyday you'll have to add it by yourself
- by default the rule set has a work week starting monday and ending friday
- by default the rule set define 8-12 14-18 as half days and 8-18 as full day
### Function Properties
- DayOfWeek[] : WorkingDaysOfWeek - Default : Monday,Tuesday,Wednesday,Thursday,Friday
- System.DateOnly[] : PublicHolidays - Default :
- System.TimeOnly : DayStartTime - Default 8h00
- System.TimeOnly : MidDayStartTime - Default 12h00
- System.TimeOnly : MidDayStopTime - Default 14h00
- System.TimeOnly : DayStopTime - Default 18h00

### OOORuleSet Properties
- Same than function properties

### OOORuleSet Method
- IsOff(DateTime) : Will return true if the given day is part of public holyday or not a working weekday

## New-PersonalHolydayRequest
Create a new Holyday request (this can be also used to create a remove request)
- While creating it will auto calculate the out of office span (not needed for remove request)
### Function Properties
- DateTime : Start
- DateTime : Stop
- OOORuleSet : OOORuleSet

### PersonalHolydayRequest Properties
- DateTime : RequestedStart
- DateTime : RequestedStop
- DateTime : OutOfOfficeStart
- DateTime : OutOfOfficeStop
- OOORuleSet : OOORuleSet

### PersonalHolydayRequest Methods
- GetOutOfOfficeSpan() : Calculate OutOfOfficeStart and OutOfOfficeStop based on Request
- IsOff(DateTime) : Return True if the given datetime is part of the holyday
- IsOff(PersonalHolydayRequest) : Return True if the given PersonalHolydayRequest coincide with holyday (Before after or during)
- Update(PersonalHolydayRequest) : Update the request with a new PersonalHolydayRequest

## New-PersonalHolydaysList
Create a list of holyday requests
### Function properties
- OOORuleSet : OOORuleSet

### PersonalHolydaysList Properties
- PersonalHolyday[] : PersonalHolyDays
- OOORuleSet : OOORuleSet

### PersonalHolydaysList Methods
- AddRequest(PersonalHolyday) : Add a Personal holyday
- RemoveRequest(PersonalHolyday) : Remove a personal holyday
