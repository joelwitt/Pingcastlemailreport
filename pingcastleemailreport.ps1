####################### EDIT THESE PARAMETERS ###########
$smtpServer = "smtp.server.local"
$smtpPort = 25
$smtpFrom = "sender@domain.com"
$smtpTo = "recipient1@domain.com; recipient2@domain.com"
########################################################

# Get working folder and domain name
$WorkingFolder = Split-Path ($myinvocation.mycommand.path) -Parent
$domain = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
$PingCastlePath = "$workingFolder\PingCastle\PingCastle.exe"
$date = (Get-Date).ToString("yyyyMMdd")

# Create report folder if it doesn't exist
If (!(Test-Path "$workingFolder\Reports")) {
    New-Item -ItemType Directory -Path "$workingFolder\Reports"
}
$XmlOutputPath = "$workingFolder\ad_hc_$domain.xml"
$NewXmlOutputPath = "$workingFolder\Reports\ad_hc_${domain}_$date.xml"
$HtmlOutputPath = "$workingFolder\ad_hc_$domain.html"
$NewHtmlOutputPath = "$workingFolder\Reports\ad_hc_${domain}_$date.html"

$LogFile = "$workingFolder\Reports\Audit-PingCastle_Run-$Date.log"

# Run PingCastle
Start-Process -FilePath $PingCastlePath -ArgumentList "--healthcheck --level Full --server $domain --log --xmls $WorkingFolder\Reports" -Wait -WindowStyle Hidden

# Wait a moment to ensure report is generated
Start-Sleep -Seconds 10

# Check if report files exist
if ((Test-Path $XmlOutputPath) -and (Test-Path $HtmlOutputPath)) {
    Move-Item -Path $XmlOutputPath  -Destination $NewXmlOutputPath -Force
    Move-Item -Path $HtmlOutputPath -Destination $NewHtmlOutputPath -Force
} else {
    "Report generation failed" | Out-File -FilePath $LogFile -Encoding unicode
    break
}

### Analyze the generated XML reports ###

# Define path to find previous report
$LastXmlOutputPath = ((Get-ChildItem -path "$WorkingFolder\Reports" -Filter *.xml | Sort-Object Name -Descending)[1]).FullName

$print_current_result = 1
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

# Function to extract XML data
Function ExtractXML($xml,$category) {
    $value = $xml.HealthcheckRiskRule | Select-Object Category, Points, Rationale, RiskId | Where-Object Category -eq $category 
    if ($value -eq $null) {
        $value = New-Object psobject -Property @{
            Category = $category
            Points = 0
        }
    }
    return $value
}

# Function to calculate group score sum
Function CaclSumGroup($a,$b,$c,$d) {
    $a1 = $a | Measure-Object -Sum Points
    $b1 = $b | Measure-Object -Sum Points
    $c1 = $c | Measure-Object -Sum Points
    $d1 = $d | Measure-Object -Sum Points
    return $a1.Sum + $b1.Sum + $c1.Sum + $d1.Sum 
}

# Function to compare equality of two sources
Function IsEqual($a,$b) {
    [int]$a1 = $a | Measure-Object -Sum Points | Select-Object -Expand Sum
    [int]$b1 = $b | Measure-Object -Sum Points | Select-Object -Expand Sum
    if($a1 -eq $b1) {
        return 1
    }
    return 0
}

# Function to get differences between reports
Function DiffReport($xml1,$xml2,$action) {
    $result = ""
    Foreach ($rule in $xml1) {
        $found = 0
        Foreach ($rule2 in $xml2) {
            if ($rule.RiskId -and $rule2.RiskId) {
                if ($action -ne "➡️" -and ($rule2.RiskId -eq $rule.RiskId)) {
                    $found = 1
                    break
                } elseIf ($action -eq "➡️" -and ($rule2.RiskId -eq $rule.RiskId) -and ($rule2.Rationale -ne $rule.Rationale)) {
                    Write-Host $action  + " *+" + $rule.Points + "* - " + $rule.Rationale $rule2.Rationale
                    $found = 2
                    break   
                }
            }
        }
        if ($found -eq 0 -and $rule.Rationale -and $action -ne "➡️") {
            Write-Host $action  + " *+" + $rule.Points + "* - " + $rule.Rationale  $rule2.RiskId $rule.RiskId
            If ($action -eq "❗") {
                $result = $result + $action  + " *+" + $rule.Points + "* - " + $rule.Rationale + "`n"
            } else {
                $result = $result + $action  + " *-" + $rule.Points + "* - " + $rule.Rationale + "`n"
            }
        } elseif ($found -eq 2 -and $rule.Rationale) {
            $result = $result + $action  + " *" + $rule.Points + "* - " + $rule.Rationale + "`n"
        }
    } 
    return $result   
}

# Load current report content
try {
    $contentPingCastleReportXML = (Select-Xml -Path $NewXmlOutputPath -XPath "/HealthcheckData/RiskRules").node
    $dateScan = [datetime](Select-Xml -Path $NewXmlOutputPath -XPath "/HealthcheckData/GenerationDate").node.InnerXML
    $Anomalies = ExtractXML $contentPingCastleReportXML "Anomalies"
    $PrivilegedAccounts = ExtractXML $contentPingCastleReportXML "PrivilegedAccounts"
    $StaleObjects = ExtractXML $contentPingCastleReportXML "StaleObjects"
    $Trusts = ExtractXML $contentPingCastleReportXML "Trusts"
    $total_point = CaclSumGroup $Trusts $StaleObjects $PrivilegedAccounts $Anomalies 
}
catch {
    Write-Error -Message ("Unable to read the content of the xml file {0}" -f $NewXmlOutputPath)
    break
}

$old_report.FullName
$current_scan = ""
$final_thread = ""

$newCategoryContent = $Anomalies + $PrivilegedAccounts + $StaleObjects + $Trusts 
Foreach ($rule in $newCategoryContent) {
    $action = "❗ *+"
    if ($rule.RiskId) {
        $current_scan = $current_scan + $action + $rule.Points + "* - " + $rule.Rationale + "`n"
    }
}
$current_scan = "`n`---`n" + $current_scan

# Load previous report
try {
    $contentOldPingCastleReportXML = (Select-Xml -Path $LastXmlOutputPath -XPath "/HealthcheckData/RiskRules").node
    $Anomalies_old = ExtractXML $contentOldPingCastleReportXML "Anomalies"  
    $PrivilegedAccounts_old = ExtractXML $contentOldPingCastleReportXML "PrivilegedAccounts" 
    $StaleObjects_old = ExtractXML $contentOldPingCastleReportXML "StaleObjects" 
    $Trusts_old = ExtractXML $contentOldPingCastleReportXML "Trusts" 
    $previous_score = CaclSumGroup $Trusts_old $StaleObjects_old $PrivilegedAccounts_old $Anomalies_old
    Write-Host "Previous Score " $previous_score
    Write-Host "Current Score " $total_point
} catch {
    Write-Error -Message ("Unable to read the content of the xml file {0}" -f $old_report)
    break
}

$newCategoryContent = $Anomalies + $PrivilegedAccounts + $StaleObjects + $Trusts
$oldCategoryContent = $Anomalies_old + $PrivilegedAccounts_old + $StaleObjects_old + $Trusts_old 
$addedVuln = DiffReport $newCategoryContent $oldCategoryContent "❗"
$removedVuln = DiffReport $oldCategoryContent $newCategoryContent "✅"
$warningVuln = DiffReport $newCategoryContent $oldCategoryContent "➡️"

if ([int]$previous_score -eq [int]$total_point -and (IsEqual $StaleObjects_old $StaleObjects) -and (IsEqual $PrivilegedAccounts_old $PrivilegedAccounts) -and (IsEqual $Anomalies_old $Anomalies) -and (IsEqual $Trusts_old $Trusts)) {
    if ($addedVuln -or $removedVuln -or $warningVuln) {
        $sentNotification = $True
    } else {
        $sentNotification = $False
    }
} elseif ([int]$previous_score -lt [int]$total_point) {
    Write-Host "Risk increased"
    $sentNotification = $true
} elseif ([int]$previous_score -gt [int]$total_point) {
    Write-Host "Risk decreased"
    $sentNotification = $true
} else {
    Write-Host "Same global score but different category scores"
    $sentNotification = $true
}

$final_thread = $addedVuln + $removedVuln + $warningVuln

try {
    "Last scan " + $dateScan | Out-File -Append $LogFile 
    $log = $BodyTeams 
    $log += $final_thread
    $log = $log.Replace("*","").Replace("🟢","").Replace("🟠","").Replace("🟡","").Replace("🔴","").Replace("❗","!").Replace("✅","-").Replace("➡️",">").Replace("🎉","")
    $log = $log.Replace("{","").Replace("   text:'","").Replace("&#129395;","")
    $log | Out-File -Append $LogFile
} catch {
    Write-Error -Message ("Error writing to log: $_")
    break
}

# Function to send email
function Send-Email {
    param(
        [string]$smtpServer,
        [int]$port,
        [string]$from,
        [string]$to,
        [string]$subject,
        [string]$body,
        [string]$attachment
    )

    $message = New-Object System.Net.Mail.MailMessage
    $message.From = $from
    $to -split ";" | ForEach-Object { $message.To.Add($_.Trim()) }
    $message.Subject = $subject
    $message.Body = $body
    $message.BodyEncoding = [System.Text.Encoding]::UTF8

    if (Test-Path $attachment) {
        $attachmentObj = New-Object System.Net.Mail.Attachment($attachment)
        $message.Attachments.Add($attachmentObj)
    } else {
        Write-Host "Attachment not found: $attachment"
        break
    }

    if (Test-Path $newHtmlOutputPath) {
        $htmlAttachmentObj = New-Object System.Net.Mail.Attachment($newHtmlOutputPath)
        $message.Attachments.Add($htmlAttachmentObj)
    } else {
        Write-Host "HTML attachment not found: $newHtmlOutputPath"
        break
    }

    $client = New-Object Net.Mail.SmtpClient($smtpServer, $port)
    $client.EnableSsl = $false

    try {
        $client.Send($message)
        Write-Host "Email successfully sent to $to"
    } catch {
        Write-Host "Error sending email: $_"
        Add-Content -Path $LogFile -Value "[$(Get-Date)] Error sending email: $_"
    }
}

# Prepare email body
$emailBody = "PingCastle report for domain $domain generated on $(Get-Date -Format 'dd/MM/yyyy').`n`n"
$emailBody += "Summary of Changes:`nPrevious Score $previous_score`nCurrent Score $total_point`n`n"
$emailBody += "New Risks:`n$addedVuln`n"
$emailBody += "Removed Risks:`n$removedVuln`n"
$emailBody += "Warnings:`n$warningVuln`n"

# Check if log file exists before sending
if (-not (Test-Path $LogFile)) {
    Write-Host "Log file not found: $logfile"
    break
}

# Send the email
Send-Email -smtpServer $smtpServer -port $smtpPort -from $smtpFrom -to $smtpTo -subject "PingCastle Report $domain" -body $emailBody -attachment $LogFile -htmlAttachment $newHtmlOutputPath