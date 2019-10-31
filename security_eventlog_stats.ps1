<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2019 v5.6.157
	 Created on:   	2019-10-18 08:39
	 Created by:   	David Olsson
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

#Requires -Module ImportExcel -Version 5

Import-Module ImportExcel
$ws_list = Get-Content ".\workstation_list.txt"
$excel_name = "workstations_securitylog_stats.xlsx"
#$ws_list = $ws_list | select -First 10
$results_array = @()
$count = 1

foreach ($ws in $ws_list)
{
	if (Test-Connection -ComputerName $ws -Count 1 -ErrorAction SilentlyContinue)
	{
		Write-Host "$count. Getting winevents for $ws..."
		try
		{
			$payload = invoke-command -ScriptBlock { Get-WinEvent -LogName "Security" -MaxEvents 1000 | select id, TaskDisplayName, KeywordsDisplayNames } -ComputerName $ws -HideComputerName
			$results_array += $payload
			
		}
		catch
		{
			Write-Host "Error getting event logs!" -ForegroundColor Red
		}
		$count = $count + 1
	}
	
}

Write-Host "Measuring data.."

#$results_array | %{ $all += $_ }
$results_array | Group-Object -Property id | sort Count -Descending

$results_array | Group-Object -Property TaskDisplayName, id, KeywordsDisplayNames | sort Count -Descending | Export-Excel ".\$excel_name" -WorkSheetname "WinEvents" -HideSheet "WinEvents"`
																									-CellStyleSB {
	param (
		$workSheet,
		$totalRows,
		$lastColumn
	)
	
	Set-CellStyle $workSheet 1 $LastColumn Solid Cyan
	foreach ($row in (2 .. $totalRows | Where-Object { $_ % 2 -eq 0 }))
	{
		Set-CellStyle $workSheet $row $LastColumn Solid Gray
	}
	
	foreach ($row in (2 .. $totalRows | Where-Object { $_ % 2 -eq 1 }))
	{
		Set-CellStyle $workSheet $row $LastColumn Solid LightGray
	}
}`
								-IncludePivotTable -PivotRows Name -PivotData @{ "Count" = 'sum' }`
							  -IncludePivotChart -ChartType PieExploded3D -AutoSize -Show

