# Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

$DateFrom   = '2020-04-25'
$DateTo     = '2020-09-12'
$DestFolder = "C:\Users\Andrei\Downloads\MI9"


$Shell = New-Object -ComObject Shell.Application
$TargetFolderShell = $Shell.NameSpace( $DestFolder ).self

# From - https://github.com/WillyMoselhy/Weekend-Projects/blob/master/Copy-MTPCameraByMonth.ps1
$ShellItem = $Shell.NameSpace(17).Self
$Phone = $ShellItem.GetFolder.Items() | Where-Object{$_.Name -eq 'MI 9'}
$f1 = $Phone.GetFolder.Items() | Where-Object{$_.Name -eq 'Internal shared storage'}
$f2 = $f1.GetFolder.Items() | Where-Object{$_.Name -eq 'DCIM'}
$f3 = $f2.GetFolder.Items() | Where-Object{$_.Name -eq 'Camera'}

$CameraFolder = $f3
$CameraItems = $CameraFolder.GetFolder.Items()

[System.Collections.ArrayList]$q = @( $CameraItems )
# $q.GetType() = 

$CultureDateTimeFormat = (Get-Culture).DateTimeFormat
$DateFormat = $CultureDateTimeFormat.ShortDatePattern
$TimeFormat = $CultureDateTimeFormat.LongTimePattern
$DateTimeFormat = "$DateFormat $TimeFormat"

while ($q.Count -ne 0) {
	$CameraItems = $q[0]
	$q.RemoveAt(0)
	foreach ($File in $CameraItems) {
		if ( $File.Type -eq "File folder") {
			Write-Warning "$($File.Name) $($File.Type)"
			$q = $q + $File.GetFolder.Items()
		} else {
			$datestr = $File.Parent.GetDetailsOf($File,3)
			$dt = [DateTime]::ParseExact($datestr,'dd.MM.yyyy H:mm',[System.Globalization.CultureInfo]::InvariantCulture)
			$dt_check = $dt.ToString("yyyy-MM-dd")
			# Write-Warning "$($File.Name) $($File.ModifyDate) | $($datestr) | $($dt_check) | $($File.Type)"
			# test against datestr and copy
			if ($dt_check -ge $DateFrom ) {
				if ($dt_check -le $DateTo) {
					Write-Warning "$($File.Name) $($File.ModifyDate) | $($datestr) | $($dt_check) | $($File.Type)"
					
					$TargetFolderShell.GetFolder.CopyHere($File)
				}
			}
		}
	}
}