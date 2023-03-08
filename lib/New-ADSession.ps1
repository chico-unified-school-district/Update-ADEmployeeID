function New-ADSession ($dc, $cmdlets, $cred) {
 $msgVars = $MyInvocation.MyCommand.Name, $dc, ($cmdLets -join ',')
 Write-Verbose ('{0},{1}' -f $msgVars)
 $adSession = New-PSSession -ComputerName $dc -Credential $cred
 if ($cmdlets) {
  Write-Verbose ('{0},Limited cmdlets' -f $MyInvocation.MyCommand.Name)
  Import-PSSession -Session $adSession -Module ActiveDirectory -CommandName $cmdLets -AllowClobber | Out-Null
 }
 else {
  Write-Verbose ('{0},All cmdlets' -f $MyInvocation.MyCommand.Name)
  Import-PSSession -Session $adSession -Module ActiveDirectory -AllowClobber | Out-Null
 }
}