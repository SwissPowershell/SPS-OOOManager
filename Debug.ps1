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

$PublicHolidays = @(
    [System.DateOnly]::New(2023,1,1)    # Nouvel Ans
    [System.DateOnly]::New(2023,4,7)    # Vendredi-Saint
    [System.DateOnly]::New(2023,4,10)   # Lundi de Pâques
    [System.DateOnly]::New(2023,5,18)   # Jeudi de l'Ascension
    [System.DateOnly]::New(2023,5,29)   # Lundi de Pentecôte
    [System.DateOnly]::New(2023,8,1)    # Fête Nationale
    [System.DateOnly]::New(2023,9,7)    # Jeûne Genevois
    [System.DateOnly]::New(2023,12,25)  # Noel
    [System.DateOnly]::New(2023,12,31)  # Restauration de la République
)
$VerbosePreference = 'Continue'
$OOORuleSet = New-OOORuleSet -PublicHolidays $PublicHolidays
#$OOORuleSet.PublicHolidays = $PublicHolidays
$NewYear = Get-Date -Year 2023 -Month 1 -Day 1
$RequestStart1 = Get-Date -Year 2023 -Month 5 -Day 30 -Hour 8 -Minute 0 -Second 0 -Millisecond 0
$RequestStop1 = Get-Date -Year 2023 -Month 6 -Day 2 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest1 = New-PersonalHolydayRequest -Start $RequestStart1 -Stop $RequestStop1 -OOORuleSet $OOORuleSet # take the week of "Lundi de pentecote"
$PersonalHolydayRequest2 = New-PersonalHolydayRequest -Start $RequestStart1 -Stop $RequestStop1 -OOORuleSet $OOORuleSet # take the week of "Lundi de pentecote" a second time (should do nothing)

$RequestStart3 = Get-Date -Year 2023 -Month 7 -Day 31 -Hour 8 -Minute 0 -Second 0 -Millisecond 0
$RequestStop3 = Get-Date -Year 2023 -Month 8 -Day 2 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest3 = New-PersonalHolydayRequest -Start $RequestStart3 -Stop $RequestStop3 -OOORuleSet $OOORuleSet # take the monday before "Fête Nationale" and the day after (Wednesday)

$RequestStop4 = Get-Date -Year 2023 -Month 8 -Day 3 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest4 = New-PersonalHolydayRequest -Start $RequestStart3 -Stop $RequestStop4 -OOORuleSet $OOORuleSet # take the monday before "Fête Nationale" and 2 days after (Thursday)

$RequestStop5 = Get-Date -Year 2023 -Month 8 -Day 4 -Hour 18 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest5 = New-PersonalHolydayRequest -Start $RequestStop3 -Stop $RequestStop5 -OOORuleSet $OOORuleSet # take the day after "Fête Nationale" and 3 days after (Friday)

$PersonalHolyDays = New-PersonalHolydaysList -OOORuleSet $OOORuleSet
$PersonalHolyDays.AddRequest($PersonalHolydayRequest1)
$PersonalHolyDays.AddRequest($PersonalHolydayRequest2)
$PersonalHolyDays.AddRequest($PersonalHolydayRequest3)
$PersonalHolyDays.AddRequest($PersonalHolydayRequest4)
$PersonalHolyDays.AddRequest($PersonalHolydayRequest5)

$RequestStart6 = Get-Date -Year 2023 -Month 8 -Day 2 -Hour 8 -Minute 0 -Second 0 -Millisecond 0
$RequestStop6 = Get-Date -Year 2023 -Month 8 -Day 3 -Hour 8 -Minute 0 -Second 0 -Millisecond 0
$PersonalHolydayRequest6 = New-PersonalHolydayRequest -Start $RequestStart6 -Stop $RequestStop6 -OOORuleSet $OOORuleSet
$PersonalHolyDays.RemoveRequest($PersonalHolydayRequest6) # Remove 2 and 3 of august
##################################
# End of the tests show mettrics #
##################################
Write-Host "------------------- Ending script -------------------" -ForegroundColor Yellow
$TimeSpentInDebugScript = New-TimeSpan -Start $DebugStart -Verbose:$False -ErrorAction SilentlyContinue;$TimeUnits = [ORDERED] @{TotalDays = "$($TimeSpentInDebugScript.TotalDays) D.";TotalHours = "$($TimeSpentInDebugScript.TotalHours) h.";TotalMinutes = "$($TimeSpentInDebugScript.TotalMinutes) min.";TotalSeconds = "$($TimeSpentInDebugScript.TotalSeconds) s.";TotalMilliseconds = "$($TimeSpentInDebugScript.TotalMilliseconds) ms."}
ForEach ($Unit in $TimeUnits.GetEnumerator()) {if ($TimeSpentInDebugScript.($Unit.Key) -gt 1) {$TimeSpentString = $Unit.Value;break}};if (-not $TimeSpentString) {$TimeSpentString = "$($TimeSpentInDebugScript.Ticks) Ticks"}
Write-Host "Ending : " -ForegroundColor Yellow -NoNewLine; Write-Host $($MyInvocation.MyCommand) -ForegroundColor Magenta -NoNewLine;Write-Host " - TimeSpent : " -ForegroundColor Yellow -NoNewLine; Write-Host $TimeSpentString -ForegroundColor Magenta