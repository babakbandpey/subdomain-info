<#
	.Description

	.SYNOPSIS
	Created By:
	Team Juno
	Jette Jerndal & Babak Bandpey
	2021

	╔═══╗    ╔╗    ╔╗                        ╔══╗     ╔═╗    
	║╔═╗║    ║║    ║║                        ╚╣╠╝     ║╔╝    
	║╚══╗╔╗╔╗║╚═╗╔═╝║╔══╗╔╗╔╗╔══╗ ╔╗╔═╗       ║║ ╔═╗ ╔╝╚╗╔══╗
	╚══╗║║║║║║╔╗║║╔╗║║╔╗║║╚╝║╚ ╗║ ╠╣║╔╗╗╔═══╗ ║║ ║╔╗╗╚╗╔╝║╔╗║
	║╚═╝║║╚╝║║╚╝║║╚╝║║╚╝║║║║║║╚╝╚╗║║║║║║╚═══╝╔╣╠╗║║║║ ║║ ║╚╝║
	╚═══╝╚══╝╚══╝╚══╝╚══╝╚╩╩╝╚═══╝╚╝╚╝╚╝     ╚══╝╚╝╚╝ ╚╝ ╚══╝

	Tool for collecting domains and subdomains information,                                                         

	This script will read the a .txt file containing a domain name or subdomain name on each lines and collect information about it.
	The informations will then be saved in a .csv file which has the same name and location as the .txt file
	The csv file will then be used to update or create an Excel worksheet with the gathered data.

	To see the full help run: Get-Help Subdomain-Info.ps -full

	Usage
	Subdomain-info.ps -subdomains full-path-to-file-containing-a-list-of-subdomain.txt -xlsx full-path-to-where-the-result-shall-be-saved.xlsx -worksheet the-name-of-the-worksheet -convert 0-or-1
	
	.PARAMETER subdomains
	Full path to a .txt file containing a list of subdomains

	.PARAMETER xlsx
	Full path to thea .xlsx Excel file which shall be updated or created

	.PARAMETER worksheet
	The name of the worksheet which shall be updated or created

	.PARAMETER convert
	If the flag convert is used, the scanning will not take place and an already existing csv file will be used to
	update the worksheet.
	The csv file must have the same name as the .txt file name

	If you wish to convert a csv file to a worksheet in the workbook without running the scans you can set -convert
#>

param (
	# subdomains The file to read from
	[Parameter(Mandatory=$true)][string]$subdomains = $(Write-Host " -subdomains parameter is required. A .txt file with subdomains full path required."), 
	# xlsx Workbook to save to
	[Parameter(Mandatory=$true)][string]$xlsx = $(Write-Host " -xlsx parameter is required. A .xlsx Excel file which the result will be saved in."),
	# worksheet Worksheet name
	[Parameter(Mandatory=$true)][string]$worksheet = $(Write-Host " -worksheet parameter is required. A Excel worksheet name which the result will be saved in."),
	# convert if set will only do conversion of an already existing csv file
	[switch]$convert
)

$subdomains_file = $subdomains.replace(";", "")
$csv_file_name = $subdomains_file.Replace("txt", "csv")
$xlsx_file_name = $xlsx.replace(";", "")

# Removing illegal characters from the worksheet name 
$escapables = ";:\\/?*"
foreach($sign in $escapables) {
	$worksheet_name = $worksheet.replace($sign, " ")
}

# Accumulating the results
$FinalResult = @()

function Get-HtmlTitle($data) {
	$title = $data.Content | Select-String '<title>.+</title>' -AllMatches
	Return $title.Matches.Value
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
		$csv_file_name,
		$xlsx_file_name,
		$worksheet_name
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
			Write-Host "Openning the xlsx"
			$workbook = $excel.Workbooks.Open($xlsx_file_name)
		} else {
			# Create a new workbook
			Write-Host "Creating a new xlsx: '$xlsx_file_name'"
			$workbook = $excel.Workbooks.Add(1)
			$new = $true
		}

		if( $new ) {
			$worksheet = $workbook.Worksheets.Item(1)
			$worksheet.Name = $worksheet_name
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
				$worksheet.Name = $worksheet_name
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


try {
	Clear-Host

	Write-Host $text_file_name

	if(-not $(Test-Path -Path $subdomains_file)) {
		throw "Subdomains File Not Found: '$subdomains_file'"
	}

	if ( $convert ) {
		Write-Host "Only Converting CSV to Excel"
		Export-Excel $csv_file_name $xlsx_file_name $worksheet_name
		Write-Host "Conversion to Excel finished"
		Write-Host "Quitting"
		return
	}

	Write-Host "The CSV result will be saved in the file: $csv_file_name"
	Write-Host "The Excel result will be saved in the file: $xlsx_file_name"
} catch {
	Write-Host "What could have gone wrong?"
	Write-Host $_.Exception.Message
	return
}

Write-Host "Reading the file: $subdomains_file"

try{
	
	$lines = Get-Content $subdomains_file
	
	Write-Host "Read $($lines.Length) lines from $subdomains_file"
	
	$lines | ForEach-Object{
		
		$Name = $_
		$ServerList = @('8.8.8.8','8.8.4.4')
		# $ServerList = @('10.0.50.11','10.0.50.12')
		$tempObj = "" | Select-Object Name, PingedAddress, PingResult, IPAddress, City, Country, ISP, DnsStatus, ErrorMessage, HttpUrl, HttpStatusCode, HttpHtmlTitle, HttpError, HttpHeadersLink, HttpsUrl, HttpsStatusCode, HttpsHtmlTitle, HttpsHeadersLink, HttpsError, FtpServer, FtpConnection, FtpError
		
		try 
		{   
			Write-Host $Name
			
			$dnsRecord = Resolve-DnsName $Name -Server $ServerList -ErrorAction Stop | Where-Object {$_.Type -eq 'A'}        
			$tempObj.Name = $Name
			$tempObj.IPAddress = ($dnsRecord.IPAddress -join ',')
			$tempObj.DnsStatus = 'OK'
			$tempObj.ErrorMessage = ''    
		}    
		catch 
		{
			$tempObj.Name = $Name
			$tempObj.IPAddress = ''
			$tempObj.DnsStatus = 'NOT_OK'        
			$tempObj.ErrorMessage = $_.Exception.Message    
		}
		
		try {
			$tempObj.PingedAddress = "" | ping $Name -4 -n 1 -w 1000			
			if( $tempObj.PingedAddress -eq "" ) {
				$tempObj.PingResult = $tempObj.PingedAddress[5].Split(",")[2]
				$tempObj.PingedAddress = $tempObj.PingedAddress[1].Split("[")[1].Split("]")[0]
				
				$IpGeolocation = Get-IPGeolocation($tempObj.PingedAddress)

				$tempObj.City = $IpGeolocation.City
				$tempObj.Country = $IpGeolocation.Country
				$tempObj.ISP = $IpGeolocation.Isp
				
				# Sleeping for 1.5 seconds to avoid blacklisting https://www.easy365manager.com/ip-geo-location-lookup-using-powershell/
				Start-Sleep -Seconds 1.5
				
			}
		} catch {
			$tempObj.PingedAddress = ''
			Write-Host $_.Exception.Message
		}
		
		if( $Name.IndexOf("ftp") -gt -1) {
			$tempObj.FtpServer = $Name
			try {
				Write-Host $tempObj.FtpServer
				$data = Test-NetConnection -ComputerName $tempObj.FtpServer -Port 21
				$tempObj.FtpConnection = $data.TcpTestSucceeded
			} catch {
				$tempObj.FtpError = $_.Exception.Message
			}
		} else {					
			$tempObj.HttpUrl = "http://$Name"
			$tempObj.HttpsUrl = "https://$Name"
			
			try {
				Write-Host $tempObj.HttpUrl
				$response = Invoke-WebRequest -Uri $tempObj.HttpUrl -TimeoutSec 2
				$tempObj.HttpStatusCode = $response.StatusCode
				
				if( $response.Headers.Link ) {
					$tempObj.HttpHeadersLink = $response.Headers.Link.Replace(",", " - ")
				}
				
				$tempObj.HttpHtmlTitle = Get-HtmlTitle $response
			
			} catch [System.Net.WebException] {
				If ($_.Exception.Response.StatusCode.value__) {
					$crap = ($_.Exception.Response.StatusCode.value__ ).ToString().Trim();
					Write-Output $crap;
				}
				If  ($_.Exception.Message) {
					$crapMessage = ($_.Exception.Message).ToString().Trim();
					Write-Output $crapMessage;
				}
			} catch {
				$tempObj.HttpError = $_.Exception.Message
			}
			
			try {
				Write-Host $tempObj.HttpsUrl
				if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type)
				{
					$certCallback = @"
						using System;
						using System.Net;
						using System.Net.Security;
						using System.Security.Cryptography.X509Certificates;
						public class ServerCertificateValidationCallback
						{
							public static void Ignore()
							{
								if(ServicePointManager.ServerCertificateValidationCallback ==null)
								{
									ServicePointManager.ServerCertificateValidationCallback += 
										delegate
										(
											Object obj, 
											X509Certificate certificate, 
											X509Chain chain, 
											SslPolicyErrors errors
										)
										{
											return true;
										};
								}
							}
						}
"@
					Add-Type $certCallback
				 }
				[ServerCertificateValidationCallback]::Ignore()

				$response = Invoke-WebRequest -Uri $tempObj.HttpsUrl -TimeoutSec 2
				$tempObj.HttpsStatusCode = $response.StatusCode
				if( $response.Headers.Link ) {
					$tempObj.HttpsHeadersLink = $response.Headers.Link.Replace(",", " - ")
				}
				
				$tempObj.HttpsHtmlTitle = Get-HtmlTitle($response)
			} catch {
				Write-Host $_.Exception
				$tempObj.HttpsError = $_.Exception.Message
			}
		}
		Write-Host $tempObj
		
		$FinalResult += $tempObj
	}

	$FinalResult | Export-Csv $csv_file_name -NoTypeInformation

	Export-Excel $csv_file_name $xlsx_file_name $worksheet_name
} catch {
	Write-Host $_.Exception.Message
}