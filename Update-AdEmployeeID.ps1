<#
 .Synopsis
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
  [Alias('ADCred')]
  [System.Management.Automation.PSCredential]$ActiveDirectoryCredential,
  # [Alias('EmpServer')]
  # [string]$EmpDBServer,
  # [Alias('EmpDB')]
  # [string]$EmpDatabase,
  # [string]$EmpTable,
  # [Alias('EmpCred')]
  # [System.Management.Automation.PSCredential]$EmpDBCredential,
  [Alias('IntServer')]
  [string]$IntermediateSqlServer,
  [Alias('IntDB')]
  [string]$IntermediateDatabase,
  [string]$AccountsTable,
  [Alias('IntCred')]
  [System.Management.Automation.PSCredential]$IntermediateCredential,
  [Alias('wi')]
  [switch]$WhatIf
)

function Clear-AccountExpiration {
  process {
    # Set empid on AD obj if no errors reported up to this point
    if ($null -eq $_.status) {
      $msg = $MyInvocation.MyCommand.Name, $_.empId, $_.mail
      Write-Host ('{0},[{1}],[{2}]' -f $msg) -Fore Blue
      $setParams = @{
        Identity              = $_.guid
        AccountExpirationDate = $null
        Confirm               = $false
        WhatIf                = $WhatIf
        ErrorAction           = 'Stop'
      }
      Set-ADUser @setParams
    }
    $_
  }
}

function Compare-EmpId {
  process {
    Write-Host ('{0},[{1}],[{2}]' -f $MyInvocation.MyCommand.Name, $_.empId, $_.mail)
    $obj = $_ | Get-ADObj
    Write-Verbose ($obj.EmployeeId | Out-String)
    Write-Verbose ($_.empId | out-string )
    if ($obj.EmployeeId -ne $_.empId) {
      Write-Error ('{0},[{1}],[{2}],EmployeeID not set correctly on AD object' -f $MyInvocation.MyCommand.Name, $_.empId, $_.mail)
      $status = 'Error - EmployeeId Not Set on AD Object.'
      $_.status = $status
    }
    $_
  }
}

function Get-IntDBData ($table, $dbParams) {
  process {
    $sql = 'SELECT * FROM {0} WHERE status IS NULL;' -f $table
    $msg = @(
      $MyInvocation.MyCommand.Name
      $dbParams.Server
      $dbParams.Database
      $dbParams.Credential.Username
      $sql
    )
    Write-Verbose ('{0},[{1}-{2}] as [{3}],[{4}]' -f $msg)
    Invoke-Sqlcmd @dbParams -Query $sql
  }
}

# function Get-EmpData ($dbParams, $table) {
#   process {
#     $sql = 'SELECT empId FROM {0} WHERE empId = {1};' -f $table, $_.empId
#     $emp = Invoke-SqlCmd @empDBParams -Query $sql
#     if (-not$emp) {
#       $msg = $MyInvocation.MyCommand.Name, $_.empId, $_.emailWork, $_.emailHome, $sql
#       Write-Error ('{0},EmpId [{1}] not found. EmailWork: [{2}],EmailHome: [{3}],[{4}]' -f $msg)
#       return
#     }
#     $_
#   }
# }

function Get-ADObj {
  process {
    $adParams = @{
      # this filter allows for our 2 types of email address
      Filter     = "Mail -eq '{0}' -or HomePage -eq '{0}'" -f $_.emailWork
      Properties = 'EmployeeId', 'Mail', 'HomePage'
    }
    Write-Verbose ($adParams.Filter | Out-String)
    Write-Verbose ($adParams.Properties | Out-String)
    $obj = Get-ADUser @adParams
    if ($obj.count -gt 1) {
      Write-Error ('Multiple AD objects with email address [{0}]' -f $_.emailWork)
      return
    }
    Write-Verbose ($obj | Out-String)
    $obj
  }
}

function New-PSObj {
  process {
    $status = $null
    Write-Verbose ('{0}' -f $MyInvocation.MyCommand.Name)
    if ($_.emailWork -is [DBNull]) {
      $msg = $MyInvocation.MyCommand.Name, $_.bpName, $_.instanceId, $_.empId
      Write-Error ('{0},BP Name:[{1}],InstanceId:[{2}],Empid [{3}], emailWork Missing from DB entry' -f $msg)
      $status = 'EmailWork Missing From DB'
      # Something went wrong in the LFForms process. Go fix it!
      return
    }
    $obj = $_ | Get-ADObj
    if ($null -eq $obj) {
      $msg = $MyInvocation.MyCommand.Name, $_.bpName, $_.instanceId, $_.empId
      Write-Warning ('{0},BP Name:[{1}],InstanceId:[{2}],Empid [{3}], AD Object not found' -f $msg)
      $status = 'AD Object Not Found'
    }
    # create object with AD ObjectGUID and Intermediate DB data
    [PSCustomObject]@{
      id         = $_.id
      empId      = $_.empId
      fn         = $_.fn
      ln         = $_.ln
      mail       = $_.emailWork
      emailWork  = $_.emailWork
      guid       = $obj.ObjectGUID
      gsuite     = $obj.HomePage
      samid      = $obj.SamAccountName
      status     = $status
      bpName     = $_.bpName
      instanceId = $_.instanceId
    }
  }
}

function Update-ADEmpId {
  process {
    # Set empid on AD obj if no errors reported up to this point
    if ($null -eq $_.status) {
      $msg = $MyInvocation.MyCommand.Name, $_.empId, $_.mail
      Write-Host ('{0},[{1}],[{2}]' -f $msg) -Fore Blue
      $setParams = @{
        Identity    = $_.guid
        EmployeeID  = $_.empId
        Confirm     = $false
        WhatIf      = $WhatIf
        ErrorAction = 'Stop'
      }
      Set-ADUser @setParams
      $_.status = 'success'
    }
    $_
  }
}

function Update-IntDB ($table, $dbParams) {
  process {
    # Write-Host ($_ | Out-String)
    $baseSql = "
      UPDATE {0}
      SET
      gsuite = '{1}'
      ,samid = '{2}'
      ,status = '{3}'
      ,dts = CURRENT_TIMESTAMP
      WHERE id = {4} ;"
    $sql = $baseSql -f $table, $_.gsuite, $_.samid, $_.status, $_.id
    $msg = $MyInvocation.MyCommand.Name, $_.empId, $_.mail, $_.status, $sql
    Write-Host ('{0},[{1}],[{2}],[{3}],[{4}]' -f $msg) -Fore Green
    if (-not$WhatIf) {
      Invoke-SqlCmd @dbparams -Query $sql
    }
  }
}

# ==================================================================

# Imported Functions
. .\lib\Clear-SessionData.ps1
. .\lib\Load-Module.ps1
. .\lib\New-ADSession.ps1
. .\lib\Select-DomainController.ps1
. .\lib\Show-TestRun.ps1

$intDBparams = @{
  Server     = $IntermediateSqlServer
  Database   = $IntermediateDatabase
  Credential = $IntermediateCredential
}

# $empDBParams = @{
#   Server     = $EmpDBServer
#   Database   = $EmpDatabase
#   Credential = $EmpDBCredential
# }

$stopTime = Get-Date "6:00pm"
$delay = 60
'Process looping every {0} seconds until {1}' -f $delay, $stopTime
do {
  Show-TestRun
  Clear-SessionData

  'SQLServer' | Load-Module

  $dc = Select-DomainController $DomainControllers
  New-ADSession -dc $dc -cmdlets 'Get-ADUser', 'Set-ADUser' -Cred $ActiveDirectoryCredential

  $intDBResults = Get-IntDBData $AccountsTable $intDBparams
  $opObjs = $intDBResults | New-PSObj
  $opObjs | Clear-AccountExpiration | Update-ADEmpId | Compare-EmpId | Update-IntDB $AccountsTable $intDBparams

  Clear-SessionData
  Show-TestRun
  if (-not$WhatIf) {
    # Loop delay
    Start-Sleep $delay
  }
} until ($WhatIf -or ((Get-Date) -ge $stopTime))

# ==================================================================