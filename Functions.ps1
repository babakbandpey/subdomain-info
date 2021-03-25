# Functions used in the Subdomain-Info.ps1

Function Filter-Chars {
	# Filtering the unwanted characters
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true)][string]$text,
		[Parameter(Mandatory=$true)][array]$filters
	)

	foreach ($item in $filters) {
		$text = $text.Replace($item, "")
	}

	return $text
}

function Get-HtmlTitle($data) {
	$title = $data.Content | Select-String '<title>.+</title>' -AllMatches
	Return $title.Matches.Value.Replace("<title>", "").Replace("</title>", "")
}

function Get-IPGeolocation {
  Param
  (
    [string]$IPAddress
  ) 
  $request = Invoke-RestMethod -Method Get -Uri "http://ip-api.com/json/$IPAddress"
  [PSCustomObject]@{
    City    = $request.city
    Country = $request.country
    Isp     = $request.isp
  }
}

function Export-Excel {

	param(
		[Parameter(Mandatory=$true)][string]$csv_file_name,
		[Parameter(Mandatory=$true)][string]$xlsx_file_name,
		[Parameter(Mandatory=$true)][string]$worksheet_name
	)

	try {

		if( -not $xlsx_file_name ) {
			throw "xlsx_file_name Variable is null"
		}

		Write-Host "**** Excel Part Running *****"
		#Define locations and delimiter
		$delimiter = "," #Specify the delimiter used in the file

		# Create a new Excel workbook with one empty sheet
		$excel = New-Object -ComObject excel.application
		$new = $false
		if(Test-Path -Path $xlsx_file_name) {
			# Open the file if file exists
			Write-Host "Openning the xlsx: '$xlsx_file_name'"
			$workbook = $excel.Workbooks.Open($xlsx_file_name)
		} else {
			# Create a new workbook
			Write-Host "Creating a new xlsx: '$xlsx_file_name'"
			$workbook = $excel.Workbooks.Add(1)
			$new = $true
		}

		if( $new ) {
			$worksheet = $workbook.Worksheets.Item(1)
			$worksheet.Name = "$worksheet_name"
		} else {
			$found = 0
			foreach( $ws in $workbook.Worksheets ) {
				if( $worksheet_name -eq $ws.Name) {
					$found = 1
					break
				}
			}

			if(-not $found) {
				$worksheet = $workbook.Worksheets.add()
				$worksheet.Activate()
				$worksheet.Name = "$worksheet_name"
				# Move the last sheet up one spot, making the new sheet the new effective last sheet
				$lastSheet = $workbook.WorkSheets.Item($workbook.WorkSheets.Count) 
				$worksheet.Move([System.Reflection.Missing]::Value, $lastSheet)
			} else {
				$worksheet = $workbook.Worksheets.Item($worksheet_name)
			}

		}

		# Build the QueryTables.Add command and reformat the data
		$TxtConnector = ("TEXT;" + $csv_file_name)
		$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
		$query = $worksheet.QueryTables.item($Connector.name)

		$query.TextFileOtherDelimiter = $delimiter
		$query.TextFileParseType  = 1
		$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
		$query.AdjustColumnWidth = 1
		# Execute & delete the import query
		$query.Refresh()
		$query.Delete()

		# Save & close the Workbook as XLSX.

		if(Test-Path -Path $xlsx_file_name) {
			Write-Host "Save"
			$workbook.Save()
		} else {
			Write-Host "SaveAs"
			$workbook.SaveAs($xlsx_file_name, 51)
		}
	} catch {
		Write-Host $_.Exception
	} finally {
		Write-Host "Closing workbook"
		$workbook.Close(0)
		Write-Host "Quitting Excel"
		$excel.Quit()

		#Check and you will see an excel process still exists after quitting
		#Remove the excel process by piping it to stop-process

		[System.GC]::Collect()
		[System.GC]::WaitForPendingFinalizers()
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
		Remove-Variable -Name excel

		Write-Host "Force Stop"
		Get-Process excel | Stop-Process -Force
	}
} # End Export-Excel

Function Get-TotalMilliseconds {
	return [Math]::Round((Get-Date).ToFileTime()/10000)
}

Function Port-Scan {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory=$true)]
		[String]
		$ip,
		[Parameter(Mandatory=$false)]
		[String]
		$ports
	)

	$filename = $ip.Replace(".", "_")
    $oxfile = "C:\Temp\" + $filename  + ".xml"
	if( $p ) {
		Write-Host "**********************"
		Write-Host "Ports: $p"
		Write-Host "**********************"
		Write-Host "Scanning $ip : $ports Starts"
		$command = "C:\Nmap\nmap.exe $ip -sS -R -p $ports -oX $oxfile --no-stylesheet" 
	} else {
		Write-Host "**********************"
		Write-Host "Scanning $ip"
		Write-Host "**********************"
		# $command = "C:\Nmap\nmap.exe $ip -p 19-26,30,32-33,37,42-43,49,53,70,79-85,88-90,99-100,106,109-111,113,119,125,135,139,143-144,146,161,163,179,199,211-212,222,254-256,259,264,280,301,306,311,340,366,389,406-407,416-417,425,427,443-445,8000,8080,8081 --open -oX $oxfile --no-stylesheet" 
		$command = "c:\Nmap\nmap.exe $ip -oX $oxfile --no-stylesheet -n -sn -PE -PP -PS21,22,23,25,80,113,443,31339 -PA80,113,443,10042 --source-port 53"
	}
	Write-Host $command
    cmd.exe /c $command
	
	# Select-Xml -Path $oxfile -XPath '/nmaprun/host/ports' | ForEach-Object { echo $_.Node.InnerXML }

	[xml]$xmlElm = Get-Content -Path $oxfile
	Write-Output $xmlElm
	Write-Host "**************************"
	Write-Host "Host Discovery Results:"
	Write-Host $xmlElm.nmaprun.host.status
	Write-Host "**************************"

	if( $p ) {
		
		Write-Host "Ports: $p $( $p )"
		echo $p

		$port_scan_results = "" | Select-Object Host, Ports, Protocol, State, Reason, ServiceName
		try {
			$port_scan_results.Host = "Up"
			$xmlElm.nmaprun.host.ports | ForEach-Object {
				$port_scan_results.Ports = $_.port.portid.Split("\r\n")
				$port_scan_results.Protocol = $_.port.protocol.Split("\r\n")
				$port_scan_results.State = $_.port.state.state.Split("\r\n")
				$port_scan_results.Reason = $_.port.state.reason.Split("\r\n")
				$port_scan_results.ServiceName = $_.port.service.name.Split("\r\n")
			}		
		}
		catch {
			Write-Host $_.Exception.Message
			$port_scan_results.Host = "Down"
		}
	
		$port_scan_results | ForEach-Object {
			Write-Host $_
		}

		
		Write-Host "Scanning $ip : $ports Ends"
		return $port_scan_results
	} else {
		try {
			echo "**************************"
			echo "Host Discovery Results:"
			echo $xmlElm.nmaprun.host.status
			echo "**************************"
		}
		catch {
			echo $_.Exception.Message
		}
		return ""
	}
	# [pscustomobject]@{hostname=$subdomain;port=$port;open=$open}
}