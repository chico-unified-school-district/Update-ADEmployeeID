<#
 .Synopsis
  Sync Employee ID from Staff Database with AD User EmployeeId via the email address listed in AD.
 .DESCRIPTION
 .EXAMPLE
 .EXAMPLE
 .INPUTS
 .OUTPUTS
 .NOTES
#>
[cmdletbinding()]
param(
 [Alias('DCs')]
 [string[]]$DomainControllers,
 [System.Management.Automation.PSCredential]$ADCredential,
 [string]$SqlServer,
 [string]$Database,
 [string]$AccountsTable,
 [System.Management.Automation.PSCredential]$SqlCredential,
 [Alias('wi')]
 [switch]$WhatIf
)
function Complete-Processing {
 process {
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.db.emailWork) -F DarkGreen
  Write-Verbose ($MyInvocation.MyCommand.Name, $_ | Out-String )
 }
}

function Compare-EmpId {
 process {
  Write-Verbose ($MyInvocation.MyCommand.Name, $_.ad.EmployeeId, $_.db.empId | out-string)
  $_.status = if ($_.ad.EmployeeId -eq $_.db.empId) { 'success' }
  $_
 }
}

function Get-IntDBData ($table, $dbParams) {
 New-SqlOperation @dbParams -Query "SELECT * FROM $table WHERE status IS NULL OR status = '';"
}

function New-Obj {
 process {
  [PSCustomObject]@{
   db     = $_
   ad     = $null
   status = $null
  }
 }
}

function Set-ADData {
 process {
  $_.ad = $null
  $_.ad = Get-ADUser -Filter "mail -eq '$($_.db.emailWork)'" -Properties *
  if (!$_.ad ) {
   $_.status = 'AD Object Not Found'
   return $_
  }
  $_
 }
}

function Update-ADObj {
 process {
  if ($_.status) { return $_ }
  Write-Host ('{0},{1},Old EmpId:{2},New EmpId:{3}' -f $MyInvocation.MyCommand.Name, $_.db.emailWork, $_.ad.EmployeeID, $_.db.empId ) -Fore Blue
  $setParams = @{
   Identity              = $_.ad.ObjectGUID
   EmployeeID            = $_.db.empId
   AccountExpirationDate = $null
   Confirm               = $false
   WhatIf                = $WhatIf
   ErrorAction           = 'Stop'
  }
  Set-ADUser @setParams
  if (!$WhatIf) { Start-Sleep 10 }
  $_
 }
}

function Update-IntDB ($table, $dbParams) {
 process {
  $sql = "UPDATE $table SET gsuite = @gsuite, samid = @sam ,status = @status ,dts = CURRENT_TIMESTAMP WHERE id = @id ;"
  $sqlVars = "gsuite=$($_.ad.HomePage)", "sam=$($_.ad.SamAccountName)", "status=$($_.status)", "id=$($_.db.id)"
  Write-Host ('{0},{1},status:{2}' -f $MyInvocation.MyCommand.Name, $_.db.emailWork, $_.status) -F DarkMagenta
  Write-Verbose ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $sql, ($sqlVars -join ','))
  if (!$WhatIf -and $_.status) { New-SqlOperation @dbparams -Query $sql -Parameters $sqlVars }
  $_
 }
}

# ==================================================================

Import-Module CommonScriptFunctions
Import-Module -Name dbatools -Cmdlet Invoke-DbaQuery, Set-DbatoolsConfig
if ($WhatIf) { Show-TestRun }
Show-BlockInfo main

$intDBparams = @{
 Server     = $SqlServer
 Database   = $Database
 Credential = $SqlCredential
}

Write-Host 'Process looping every 60 seconds until 6PM' -F Green
do {
 Clear-SessionData

 $results = Get-IntDBData $AccountsTable $intDBparams | New-Obj
 if ($results) {
  Connect-ADSession -DomainControllers $DomainControllers -Credential $ADCredential -Cmdlets 'Get-ADUser', 'Set-ADUser'
  $results |
   Set-ADData |
    Update-ADObj |
     Set-ADData |
      Update-IntDB $AccountsTable $intDBparams |
       Complete-Processing
 }

 Clear-SessionData
 if (!$WhatIf) { Start-Sleep 60 }
} until ($WhatIf -or ((Get-Date) -ge (Get-Date "6:00pm")))
Show-BlockInfo End
if ($WhatIf) { Show-TestRun }