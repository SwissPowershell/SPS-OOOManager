<#
    .SYNOPSIS
        This script file help reloading a module and test it's internal functions.
    .DESCRIPTION
        This script file help reloading a module and test it's internal functions.
    .INPUTS
        None.
    .OUTPUTS
        None.
    .LINK
        https://github.com/SwissPowershell/PowershellHelpers/tree/main/DebugModule.ps1
#>
# Import the module based on the current directory
# Get the module name, version and definition file
$ModuleVersion = Split-Path -Path $PSScriptRoot -leaf;$ModuleName = Split-Path -Path $(Split-Path -Path $PSScriptRoot) -leaf;$ModuleDefinitionFile = Get-ChildItem -Path $PSScriptRoot -Filter '*.psd1'
# Remove the module (if loaded)
if (Get-Module -Name $ModuleName -ErrorAction Ignore) {Try {Remove-Module -Name $ModuleName} Catch {Write-Warning "Unable to remove module: $($_.Exception.Message)"};Write-Warning "The script cannot continue...";BREAK}
# Add the module using the definition file
Try {Import-Module $ModuleDefinitionFile.FullName -ErrorAction Stop}Catch {Write-Warning "Unable to load the module: $($_.Exception.Message)";Write-Warning "The script cannot continue...";BREAK}
# Control that the module added is in the same version as the detected version
$Module = Get-Module -Name $ModuleName -ErrorAction Ignore
if (($Module | Select-Object -ExpandProperty Version) -ne $ModuleVersion) {Write-Warning "The module version loaded does not match the folder version: please review !";Write-Warning "The script cannot continue...";BREAK}
# List all the exposed function from the module
Write-Host "Module [" -ForegroundColor Yellow -NoNewline; Write-Host $ModuleName -NoNewline -ForegroundColor Magenta; Write-Host "] Version [" -ForegroundColor Yellow -NoNewline;Write-Host $ModuleVersion -NoNewline -ForegroundColor Magenta;Write-Host "] : " -NoNewline; Write-Host "Loaded !" -ForegroundColor Green
if ($Module.ExportedCommands.count -gt 0) {Write-Host "Available Commands:" -ForegroundColor Yellow;$Module.ExportedCommands | ForEach-Object {Write-Host "`t - $($_.Keys)" -ForegroundColor Magenta};Write-Host ''}Else{Write-Host "`t !! There is no exported command in this module !!" -ForegroundColor Red}
Write-Host "------------------ Starting script ------------------" -ForegroundColor Yellow
$DebugStart = Get-Date
############################
# Test your functions here #
############################

$PublicHolydays = @(
    $([DateTime]::Parse('2023-01-01T00:00:00'))     # Nouvel Ans
    $([DateTime]::Parse('2023-04-07T00:00:00'))     # Vendredi-Saint
    $([DateTime]::Parse('2023-04-10T00:00:00'))     # Lundi de Paques
    $([DateTime]::Parse('2023-05-18T00:00:00'))     # Jeudi de l'Ascension
    $([DateTime]::Parse('2023-05-29T00:00:00'))     # Lundi de Pentecote
    $([DateTime]::Parse('2023-08-01T00:00:00'))     # Fete Nationale
    $([DateTime]::Parse('2023-09-07T00:00:00'))     # Jeune Genevois
    $([DateTime]::Parse('2023-12-25T00:00:00'))     # Noel
    $([DateTime]::Parse('2023-12-31T00:00:00'))     # Restauration de la Republique
)
# add the week end to my PublicHolidays as it is in our database otherwyse you can set $OOORuleSet.WeekEndInHolidays = $False and it will take into account the given working days
$Start = [DateTime]::Parse('2023-01-01T00:00:00')
$CurrDate = $Start
Do {
    if (($CurrDate.DayOfWeek -eq [DayOfWeek]::Saturday) -or ($CurrDate.DayOfWeek -eq [DayOfWeek]::Sunday)) {
        if ($CurrDate -NotIn $PublicHolydays) {
            $PublicHolydays += $CurrDate
        }
    }
    $CurrDate = $CurrDate.AddDays(1)
}Until($CurrDate.Year -eq 2024)
$VerbosePreference = 'Continue'
$O3RuleSet = New-O3RuleSet
$O3RuleSet.PublicHolydays = $PublicHolydays

$PersonalHolyDays = New-O3List -O3RuleSet $O3RuleSet

$RequestStart1 = Get-Date -Year 2023 -Month 5 -Day 30 -Hour 7 -Minute 0 -Second 0 -Millisecond 0
$RequestStop1 = Get-Date -Year 2023 -Month 6 -Day 2 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest1 = New-O3Request -Start $RequestStart1 -Stop $RequestStop1 -O3RuleSet $O3RuleSet # take the week of "Lundi de pentecote"
$PersonalHolyDays.AddRequest($PersonalHolydayRequest1)

$PersonalHolydayRequest2 = New-O3Request -Start $RequestStart1 -Stop $RequestStop1 -O3RuleSet $O3RuleSet # take the week of "Lundi de pentecote" a second time (should do nothing)
$PersonalHolyDays.AddRequest($PersonalHolydayRequest2)


$RequestStart3 = Get-Date -Year 2023 -Month 7 -Day 31 -Hour 7 -Minute 0 -Second 0 -Millisecond 0
$RequestStop3 = Get-Date -Year 2023 -Month 8 -Day 2 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest3 = New-O3Request -Start $RequestStart3 -Stop $RequestStop3 -O3RuleSet $O3RuleSet # take the monday before "Fete Nationale" and the day after (Wednesday)

$PersonalHolyDays.AddRequest($PersonalHolydayRequest3)

$RequestStop4 = Get-Date -Year 2023 -Month 8 -Day 3 -Hour 12 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest4 = New-O3Request -Start $RequestStart3 -Stop $RequestStop4 -O3RuleSet $O3RuleSet # take the monday before "Fete Nationale" and 1.5 days after (Thursday)

$PersonalHolyDays.AddRequest($PersonalHolydayRequest4)

$RequestStop5 = Get-Date -Year 2023 -Month 8 -Day 4 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest5 = New-O3Request -Start $RequestStop3 -Stop $RequestStop5 -O3RuleSet $O3RuleSet # take the day after "Fete Nationale" and 3 days after (Friday)
$PersonalHolyDays.AddRequest($PersonalHolydayRequest5)


$RequestStart6 = Get-Date -Year 2023 -Month 8 -Day 2 -Hour 7 -Minute 0 -Second 0 -Millisecond 0
$RequestStop6 = Get-Date -Year 2023 -Month 8 -Day 2 -Hour 12 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest6 = New-O3Request -Start $RequestStart6 -Stop $RequestStop6 -O3RuleSet $O3RuleSet -Remove
$PersonalHolyDays.RemoveRequest($PersonalHolydayRequest6) # Remove 2 in the morning shoud do a 2+

$RequestStart7 = Get-Date -Year 2023 -Month 8 -Day 3 -Hour 7 -Minute 0 -Second 0 -Millisecond 0
$RequestStop7 = Get-Date -Year 2023 -Month 8 -Day 4 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest7 = New-O3Request -Start $RequestStart7 -Stop $RequestStop7 -O3RuleSet $O3RuleSet -Remove
$PersonalHolyDays.RemoveRequest($PersonalHolydayRequest7) # Remove 3 and 4 of august

$RequestStart8 = Get-Date -Year 2023 -Month 8 -Day 3 -Hour 14 -Minute 0 -Second 0 -Millisecond 0
$RequestStop8 = Get-Date -Year 2023 -Month 8 -Day 4 -Hour 12 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest8 = New-O3Request -Start $RequestStart8 -Stop $RequestStop8 -O3RuleSet $O3RuleSet
$PersonalHolyDays.AddRequest($PersonalHolydayRequest8) # Take from 3 14h to 4 12h => 3+ 4-

# TO DO :
#   Check what happen if I take 1 week off 1 week in then 1 week off and then I take the week in => off
#   4.9 -> 8.9
#   18.9 -> 22.9
#   11.9 -> 15.9
#   It normally should become 1 unique off from 2 to 24 (but it will be two) => This has to be resolved

$RequestStart9 = Get-Date -Year 2023 -Month 9 -Day 4 -Hour 7 -Minute 0 -Second 0 -Millisecond 0
$RequestStop9 = Get-Date -Year 2023 -Month 9 -Day 8 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest9 = New-O3Request -Start $RequestStart9 -Stop $RequestStop9 -O3RuleSet $O3RuleSet
$PersonalHolyDays.AddRequest($PersonalHolydayRequest9) #   4.9 -> 8.9

$RequestStart10 = Get-Date -Year 2023 -Month 9 -Day 18 -Hour 7 -Minute 0 -Second 0 -Millisecond 0
$RequestStop10 = Get-Date -Year 2023 -Month 9 -Day 22 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest10 = New-O3Request -Start $RequestStart10 -Stop $RequestStop10 -O3RuleSet $O3RuleSet
$PersonalHolyDays.AddRequest($PersonalHolydayRequest10) #   18.9 -> 22.9

$RequestStart11 = Get-Date -Year 2023 -Month 9 -Day 11 -Hour 7 -Minute 0 -Second 0 -Millisecond 0
$RequestStop11 = Get-Date -Year 2023 -Month 9 -Day 15 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest11 = New-O3Request -Start $RequestStart11 -Stop $RequestStop11 -O3RuleSet $O3RuleSet
$PersonalHolyDays.AddRequest($PersonalHolydayRequest11) #   11.9 -> 15.9

$VerbosePreference = 'SilentlyContinue'
Write-O3Calendar -StartYear 2023 -PersonalHolyDays $PersonalHolyDays -OOORuleSet $O3RuleSet

##################################
# End of the tests show mettrics #
##################################
Write-Host "------------------- Ending script -------------------" -ForegroundColor Yellow
$TimeSpentInDebugScript = New-TimeSpan -Start $DebugStart -Verbose:$False -ErrorAction SilentlyContinue;$TimeUnits = [ORDERED] @{TotalDays = "$($TimeSpentInDebugScript.TotalDays) D.";TotalHours = "$($TimeSpentInDebugScript.TotalHours) h.";TotalMinutes = "$($TimeSpentInDebugScript.TotalMinutes) min.";TotalSeconds = "$($TimeSpentInDebugScript.TotalSeconds) s.";TotalMilliseconds = "$($TimeSpentInDebugScript.TotalMilliseconds) ms."}
ForEach ($Unit in $TimeUnits.GetEnumerator()) {if ($TimeSpentInDebugScript.($Unit.Key) -gt 1) {$TimeSpentString = $Unit.Value;break}};if (-not $TimeSpentString) {$TimeSpentString = "$($TimeSpentInDebugScript.Ticks) Ticks"}
Write-Host "Ending : " -ForegroundColor Yellow -NoNewLine; Write-Host $($MyInvocation.MyCommand) -ForegroundColor Magenta -NoNewLine;Write-Host " - TimeSpent : " -ForegroundColor Yellow -NoNewLine; Write-Host $TimeSpentString -ForegroundColor Magenta
