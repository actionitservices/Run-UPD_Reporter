#.SYNOPSIS
#The following script is will scan and report on User Profile Disks UVHDs via email.
#Usernames and matching UVHDs are identified using the AD securityidentifier (SID).
#User sessions must be logged in or UVHD mounted for volume health details.  
#UPDPath will be scanned for *.vhdx files only.
#
#.EXAMPLE
# .\Run-UPD_Reporter.ps1 -To admin@example.com
#
#.NOTES 
#        File Name : .\Run-UPD_Reporter.ps1
#        Authors   : Neil Hennessy (Action IT Services)
#		 Email:	   : info@actionitservices.com.au
#        Requires  : Windows Server and User Profile Disks (UVHD)
#
#        Version   : 1.13 (13/12/17) - Initial Script 
#        Version   : 1.14 (16/12/17) - General Release

param(
[Parameter(mandatory=$true)][string]$To
)

$scriptVer = "1.14"

Write-Host "Run-UPD_Reporter $scriptVer Started"

#User Configuration

#Set SMTP and sender address
$E_Host = "smtp.example.com"
$E_FROM = "server@example.com"

#Set path to User Profile Disks (UVHD) and export path to archive html reports.
$UPDPath = "U:\"
$REPORT_PATH = "C:\Temp\"

#System Configuration

$E_TO = $To
$E_BODY = ""

$head = @"
<!DOCTYPE HTML PUBLIC “-//W3C//DTD HTML 4.01 Frameset//EN” “http://www.w3.org/TR/html4/frameset.dtd”&gt;
<html><head><title>Unmigrated Systems Report</title><meta http-equiv=”refresh” content=”120″ />
<style type=”text/css”>
<!–
body {
font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
}
 
#report { width: 835px; }
 
table{
border-collapse: collapse;
border: none;
font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
color: black;
margin-bottom: 10px;
}
 
table td{
font-size: 12px;
padding-left: 0px;
padding-right: 20px;
text-align: left;
}
 
table th {
font-size: 12px;
font-weight: bold;
padding-left: 0px;
padding-right: 20px;
text-align: left;
}
 
h2{ clear: both; font-size: 130%;color:#354B5E; }
 
h3{
clear: both;
font-size: 75%;
margin-left: 20px;
margin-top: 30px;
color:#475F77;
}
 
p{ margin-left: 20px; font-size: 12px; }
 
table.list{ float: left; }
 
table.list td:nth-child(1){
font-weight: bold;
border-right: 1px grey solid;
text-align: right;
}
 
table.list td:nth-child(2){ padding-left: 7px; }
table tr:nth-child(even) td:nth-child(even){ background: #BBBBBB; }
table tr:nth-child(odd) td:nth-child(odd){ background: #F2F2F2; }
table tr:nth-child(even) td:nth-child(odd){ background: #DDDDDD; }
table tr:nth-child(odd) td:nth-child(even){ background: #E5E5E5; }
div.column { width: 320px; float: left; }
div.first{ padding-right: 20px; border-right: 1px grey solid; }
div.second{ margin-left: 30px; }
table{ margin-left: 20px; }
–>
</style>
</head>
"@

function getStatusColour ($status)
{
    switch ($status)
    {

        "OK"    {"#82C61B"}
        "WARN"  {"#F8991D"}
        "ERROR" {"#FF0000"}
        default {"#FFFFFF"}
    }

}


$date = Get-Date
$NOW = $date.ToShortDateString() + " " + $date.ToShortTimeString()

if (-Not (Test-Path $UPDPath)) {
    Write-Warning -Message "Abort script - Invalid UPD Path: $UPDPath"
    Exit
}

$PathCCheck

$UPDfolder = Get-ChildItem -Path $UPDPath -Filter "*.vhdx" 

$E_BODY = $head
$E_BODY += "<h2>User Profile Disk (UVHD) - Usage Reporter</h2>"
$E_BODY += @"
    <p>Version:  $scriptVer<br>
    Datetime:  $NOW<br
    UPD Path:  $UPDPath</p>
	<table style="width:100%;" cellpadding="2">
		<tr style="background-color: #000000;color: #FFFFFF">
			<td><strong>Username:</strong></td>
			<td><strong>UPD</strong></td>			
			<td align="center"><strong>HealthStatus</strong></td>
            <td align="center"><strong>OperationalStatus</strong></td>
            <td align="center"><strong>Size (GB)</strong></td>
            <td align="center"><strong>Size Remaining (GB)</strong></td>
            <td align="center"><strong>Free Space (%)</strong></td>
		</tr>		
"@
 
foreach ($UPDV in $UPDfolder)
{
  $sid = $UPDV.Name 
  $sid = $sid.Substring(5,$sid.Length-10)
  if ($sid -ne "template")
  {
    # Match user to SID
    $securityidentifier = new-object security.principal.securityidentifier $sid
    $user = ( $securityidentifier.translate( [security.principal.ntaccount] ) )  

    # Get VHD Volume Details
    $theUPD = $UPDPath + $UPDV.Name
    
    $checkMounted = Get-DiskImage –ImagePath $theUPD
    
    if ($checkMounted.Attached -eq "True") {

        $VHDInfo = Get-DiskImage –ImagePath $theUPD | Get-Disk | Get-Partition | Get-Volume
            
        $VHDTotalGB = [math]::round(($VHDInfo.Size / 1073741824),2)
        $VHDFreeGB = [math]::round(($VHDInfo.SizeRemaining / 1073741824),2)
    
       
        if ($VHDInfo.SizeRemaining -ne 0) {
           $VHDFreePrec = [math]::round(((($VHDInfo.SizeRemaining / 1073741824)/($VHDInfo.Size / 1073741824)) * 100),0)        
        }
        else {$VHDFreePrec = 0}
    
        $FreeColour = getStatusColour("OK")
        $HealthColour = getStatusColour("OK")
        $StatusColour = getStatusColour("OK")

        if ($VHDFreePrec -le 15) {
            $FreeColour = getStatusColour("WARN")
        }

        if ($VHDFreePrec -le 5) {
            $FreeColour = getStatusColour("ERROR")
        }

        $theHealth = $VHDInfo.HealthStatus
        $theOp = $VHDInfo.OperationalStatus
       
        switch ($theHealth) 
	    {
		    "Healthy"  {$HealthColour = getStatusColour("OK")}
		    "Warning"  {$HealthColour = getStatusColour("WARN")}
            default {$HealthColour = getStatusColour("ERROR")}
        }
	
	    switch ($theOp) 
	    {
		    "OK"  {$StatusColour = getStatusColour("OK")}
		    "Spot Fix Needed"  {$StatusColour = getStatusColour("WARN")}
            default {$StatusColour = getStatusColour("ERROR")}
        }
    
    }
    else {

        Write-Warning -Message "User not logged in -Skipping UPD Health Status!"

        $FreeColour = getStatusColour("WARN")
        $HealthColour = getStatusColour("WARN")
        $StatusColour = getStatusColour("WARN")


        $theHealth = "Not Available"
        $theOp = "Not Available"
        $VHDFreePrec = "Not Available"

    }



    $E_BODY += "<tr style=""background-color: #E1E1E1"">" + "<td>" + $user + "</td><td>" + $theUPD + "</td><td style=""color: $HealthColour"" align=""center"">" + $theHealth + "</td><td style=""color: $StatusColour"" align=""center"">" + $theOp + "</td><td align=""center"">" + $VHDTotalGB + "</td><td align=""center"">" + $VHDFreeGB + "</td><td style=""color: $FreeColour"" align=""center"">" + $VHDFreePrec + "</td></tr>"
    }
  
}
$E_BODY += "</tr></table>"

$timestamp = Get-Date -Format o | foreach {$_ -replace ":", "-"}
$HTMLOutput = $REPORT_PATH + "upd_usage_report_" + $timestamp + ".html"
Write-Host "Saving Report to: $HTMLOutput"

$E_BODY | Out-File $HTMLOutput

Write-Host "Sending Report to: $E_TO"

Send-MailMessage -Body $E_BODY -From $E_FROM -To $E_TO -SmtpServer $E_Host -Subject "UPD Usage Reporter ($NOW)" -BodyAsHtml

Write-Host "Run-UPD_Reporter Finished"
