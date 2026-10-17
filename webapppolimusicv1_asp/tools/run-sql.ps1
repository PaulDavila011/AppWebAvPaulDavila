<#
  Ejecuta un script .sql (separado por GO) usando ADO + MSOLEDBSQL19, sin necesitar sqlcmd.
  Uso: powershell -ExecutionPolicy Bypass -File tools\run-sql.ps1 -File database\01_seed_roles_usuarios.sql
#>
param(
  [Parameter(Mandatory = $true)][string]$File,
  [string]$Server = 'localhost',
  [string]$User = 'usr_polimusic',
  [string]$Password = 'Politecnica1'
)
$ErrorActionPreference = 'Stop'
$sql = Get-Content -Raw -Encoding UTF8 $File
$batches = [regex]::Split($sql, '(?im)^\s*GO\s*$') | Where-Object { $_.Trim() -ne '' }
$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=MSOLEDBSQL19;Data Source=$Server;Initial Catalog=master;User ID=$User;Password=$Password;Use Encryption for Data=Mandatory;Trust Server Certificate=True")
foreach ($b in $batches) {
  $rs = $conn.Execute($b)
  while ($rs -ne $null) {
    if ($rs.State -eq 1) {
      while (-not $rs.EOF) {
        $vals = @(); for ($i = 0; $i -lt $rs.Fields.Count; $i++) { $vals += "$($rs.Fields.Item($i).Name)=$($rs.Fields.Item($i).Value)" }
        Write-Host ('  ' + ($vals -join '  '))
        $rs.MoveNext()
      }
    }
    try { $rs = $rs.NextRecordset() } catch { $rs = $null }
  }
}
$conn.Close()
Write-Host "OK: $File"
