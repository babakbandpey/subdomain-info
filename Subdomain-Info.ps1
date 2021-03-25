<#
	.Description

	.SYNOPSIS
	Created By:
	Team Juno
	Babak Bandpey
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
	Subdomain-info.ps -subdomains 'full-path-to-file-containing-a-list-of-subdomain.txt' -xlsx 'full-path-to-where-the-result-shall-be-saved.xlsx' -worksheet 'the-name-of-the-worksheet'
	
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

	.PARAMETER p
	Port number(s) to scan by nmap. Can be written as 1,2,3,4 or 2-20 or 2-20,80,443,8000,8080

	.PARAMETER d
	Host discovery by nmap. To see if the host is up or not.

	.PARAMETER f
	Fast mode: IP-address discovery of the subdomains only
#>

param (
	# subdomains The file to read from
	[Parameter(Mandatory=$true)][string]$subdomains = $(Write-Host " -subdomains parameter is required. A .txt file with subdomains full path required."), 
	# xlsx Workbook to save to
	[Parameter(Mandatory=$true)][string]$xlsx = $(Write-Host " -xlsx parameter is required. A .xlsx Excel file which the result will be saved in."),
	# worksheet Worksheet name
	[Parameter(Mandatory=$true)][string]$worksheet = $(Write-Host " -worksheet parameter is required. A Excel worksheet name which the result will be saved in."),
	[Parameter(Mandatory=$false)][string]$p = $(Write-Host " -p is for the port numbers usage: -p 20-80,1135,443"),
	[Parameter(Mandatory=$false)][switch]$d = $(Write-Host " -d is for the host discovery"),
	[Parameter(Mandatory=$false)][switch]$f = $(Write-Host " -f is for ip-address discovery"),
	# convert if set will only do conversion of an already existing csv file
	[switch]$convert
)

Clear-Host

. ".\Functions.ps1"

# Accumulating the results
$FinalResult = @()

# Removing illegal characters
$subdomains_file = Filter-Chars $subdomains @(";")
$csv_file_name = $subdomains_file.Replace("txt", "csv")
$xlsx_file_name = Filter-Chars $xlsx @(";")
$worksheet_name = Filter-Chars $worksheet ";",":","\\","/","?","*"

if( $p -ne "" ){
	$p = Filter-Chars $p ";"," -"
}

try {

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
	if( $f ) {
		Write-Host "******************************************************"
		Write-Host "* Running -f fast mode. Discovering ip-address only. *"
		Write-Host "******************************************************"
	}

	$lines = Get-Content $subdomains_file
	
	Write-Host "Read $($lines.Length) lines from $subdomains_file"
	
	$lines | ForEach-Object{
		
		$start_time = Get-TotalMilliseconds

		$subdomain = $_
		Write-Host $subdomain

		$ServerList = @('8.8.8.8','8.8.4.4')
		# $ServerList = @('10.0.50.11','10.0.50.12')
		$tempObj = "" | Select-Object Name, PingedAddress, PingResult, IPAddress, City, Country, ISP, DnsStatus, ErrorMessage, HttpUrl, HttpStatusCode, HttpHtmlTitle, HttpError, HttpHeadersLink, HttpsUrl, HttpsStatusCode, HttpsHtmlTitle, HttpsHeadersLink, HttpsError, FtpServer, FtpConnection, FtpError, PortScan
		
		try {
			$tempObj.Name = $subdomain

			if( $f ) {
				try {
					$testNetConnection = Test-NetConnection -ComputerName $subdomain
					$tempObj.PingedAddress = $testNetConnection.RemoteAddress
					$tempObj.PingResult = $testNetConnection.PingSucceeded
					# $Timeout = 100
					# $Ping = New-Object System.Net.NetworkInformation.Ping
					# $Response = $Ping.Send($subdomain, $Timeout)

					# $tempObj.PingedAddress = $Response.Address
					# $tempObj.PingResult = $Response.Status
				} catch {
					$tempObj.PingedAddress = ""
					$tempObj.PingResult = "Failed"
				}

				$FinalResult += $tempObj
				return
			}
			
			$testNetConnection = Test-NetConnection -ComputerName $subdomain -InformationLevel "Detailed"
			$tempObj.PingedAddress = $testNetConnection.RemoteAddress
			$tempObj.PingResult = $testNetConnection.PingSucceeded

			try 
			{   
				$dnsRecord = Resolve-DnsName $subdomain -Server $ServerList -ErrorAction Stop | Where-Object {$_.Type -eq 'A'}        
				$tempObj.IPAddress = ($dnsRecord.IPAddress -join ',')
				$tempObj.DnsStatus = 'OK'
				$tempObj.ErrorMessage = ''    
			}    
			catch 
			{
				$tempObj.Name = $subdomain
				$tempObj.IPAddress = ''
				$tempObj.DnsStatus = 'NOT_OK'        
				$tempObj.ErrorMessage = $_.Exception.Message    
			}

			Write-Host $tempObj.PingedAddress

			if( $tempObj.PingedAddress -ne "" ) {			
				echo "GEO LOCATION"

				$IpGeolocation = Get-IPGeolocation($tempObj.PingedAddress)

				$tempObj.City = $IpGeolocation.City
				$tempObj.Country = $IpGeolocation.Country
				$tempObj.ISP = $IpGeolocation.Isp
				
				$elapsed_time = Get-TotalMilliseconds - $start_time
				if($elapsed_time -lt 1500 ) {
					# Sleeping for 1.5 seconds to avoid blacklisting https://www.easy365manager.com/ip-geo-location-lookup-using-powershell/
					Start-Sleep -Seconds (1500 - $elapsed_time)
				}
				
			}
		} catch {
			$tempObj.PingedAddress = ''
			Write-Host $_.Exception.Message
		}
		
		# Port scanning part
		
		try {
			if( $d ) {
				$tempObj.PortScan = Port-Scan $tempObj.PingedAddress
			}elseif ( $p ) {
				$tempObj.PortScan = Port-Scan $tempObj.PingedAddress $p

				# $port_numbers = $p.Replace(" ", "").Split(",")
				
				# foreach($port_number in $port_numbers) {
				# 	if( $port_number.IndexOf("-") -ne -1) {
				# 		$ports_range = $port_number.Split("-")
				# 		$min_port = [Convert]::ToInt32($ports_range[0])
				# 		$max_port = [Convert]::ToInt32($ports_range[1])
				# 		for ($port_number = $min_port; $port_number -lt $max_port; $port_number++) {
				# 			Write-Host $port_number
				# 			Port-Scan $tempObj.PingedAddress $port_number
				# 		}
						
				# 	} else {
				# 		Write-Host $port_number
				# 		Port-Scan $tempObj.PingedAddress $port_number
				# 	}
				# }
			}
		}
		catch {
			Write-Host $_.Exception.Message
		}
		
		# *******************
		# Port Scan part ends
		# *******************

		if( $subdomain.IndexOf("ftp") -gt -1) {
			$tempObj.FtpServer = $subdomain
			try {
				Write-Host $tempObj.FtpServer
				$data = Test-NetConnection -ComputerName $tempObj.FtpServer -Port 21
				$tempObj.FtpConnection = $data.TcpTestSucceeded
			} catch {
				$tempObj.FtpError = $_.Exception.Message
			}
		} else {					
			$tempObj.HttpUrl = "http://$subdomain"
			$tempObj.HttpsUrl = "https://$subdomain"
			
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