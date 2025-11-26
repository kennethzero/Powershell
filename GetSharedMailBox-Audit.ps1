# =============================================
# Shared Mailbox Audit – Permissions + Forwarding
# Full Access + Send As on SAME ROW
# Location: Hong Kong | Time: HKT | Export: HTML + Email
# =============================================

# ---------- CONFIG ----------
$FromEmail   = "abc@abc.com.hk"
$ToEmail     = "abc@abc.com.hk"
$SmtpServer  = "abc-com-hk.mail.protection.outlook.com"
$SmtpPort    = 25
$ReportTitle = "Shared Mailbox Permissions & Forwarding Report"

# ---------- HKT TIME ----------
$HKTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("China Standard Time")
$NowHKT     = [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-Date).ToUniversalTime(), $HKTimeZone)

Write-Host "Starting audit at: $($NowHKT.ToString('yyyy-MM-dd HH:mm:ss')) HKT" -ForegroundColor Cyan

# ---------- CONNECT ----------
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -UserPrincipalName "abc@abc.com.hk" -ShowProgress $true -ErrorAction Stop
    Write-Host "Connected!" -ForegroundColor Green
}
catch { Write-Error "Connection failed: $($_.Exception.Message)"; exit }

# ---------- GET SHARED MAILBOXES ----------
Write-Host "`nRetrieving shared mailboxes..." -ForegroundColor Cyan
$SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Sort-Object DisplayName
$total = $SharedMailboxes.Count
Write-Host "Found $total shared mailbox(es)." -ForegroundColor Yellow

# ---------- DATA COLLECTION ----------
$ReportData = @()
$cnt = 0

foreach ($mb in $SharedMailboxes) {
    $cnt++
    Write-Progress -Activity "Processing" -Status "$cnt of $total" -PercentComplete ($cnt/$total*100)

    $id   = $mb.PrimarySmtpAddress
    $name = $mb.DisplayName

    # ---- Forwarding ----
    $fwd = Get-Mailbox -Identity $id | Select-Object ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward
    $FwdTo = $null; $FwdType = $null; $MembersHTML = $null

    if ($fwd.ForwardingAddress) {
        try {
            $rec = Get-Recipient -Identity $fwd.ForwardingAddress -ErrorAction Stop
            $FwdTo = $rec.DisplayName; $FwdType = $rec.RecipientType
            if ($rec.RecipientType -like "*Group*") {
                $members = Get-DistributionGroupMember -Identity $rec.Identity -ResultSize Unlimited |
                           Select-Object DisplayName, PrimarySmtpAddress, RecipientType
                if ($members) {
                    $rows = foreach ($m in $members) {
                        "<tr><td>$([System.Web.HttpUtility]::HtmlEncode($m.DisplayName))</td><td>$($m.PrimarySmtpAddress)</td><td>$($m.RecipientType)</td></tr>"
                    }
                    $MembersHTML = "<table style='margin-top:8px;font-size:11px;'><tr style='background:#f0f0f0;'><th>Name</th><th>Email</th><th>Type</th></tr>$($rows -join '')</table>"
                }
            }
        } catch { $FwdTo = $fwd.ForwardingAddress; $FwdType = "Unknown" }
    }
    elseif ($fwd.ForwardingSmtpAddress) {
        $FwdTo = $fwd.ForwardingSmtpAddress; $FwdType = "External SMTP"
    }

    # ---- Permissions ----
    $Full = Get-MailboxPermission -Identity $id -ErrorAction SilentlyContinue |
            Where-Object {
                $_.AccessRights -contains 'FullAccess' -and
                $_.User -notlike 'NT AUTHORITY\*' -and
                $_.IsInherited -eq $false -and
                $_.Deny -eq $false
            } | ForEach-Object { ($_.User -split '\\')[-1] -replace '>', '' } | Sort-Object -Unique

    $SendAs = Get-RecipientPermission -Identity $id -ErrorAction SilentlyContinue |
              Where-Object {
                  $_.AccessRights -contains 'SendAs' -and
                  $_.Trustee -notlike 'NT AUTHORITY\*' -and
                  $_.IsInherited -eq $false
              } | ForEach-Object { ($_.Trustee -split '\\')[-1] -replace '>', '' } | Sort-Object -Unique

    $OnBehalf = $mb.GrantSendOnBehalfTo | ForEach-Object {
        $_.ToString().Split(',')[0].Trim()
    } | Where-Object { $_ } | Sort-Object -Unique

    # ---- Combine Full Access + Send As on SAME ROW ----
    $userPerms = @{}

    foreach ($u in $Full)    { $userPerms[$u] = $userPerms[$u] + @("Full Access") }
    foreach ($u in $SendAs)  { $userPerms[$u] = $userPerms[$u] + @("Send As") }
    foreach ($u in $OnBehalf){ $userPerms[$u] = $userPerms[$u] + @("Send on Behalf") }

    # Build HTML rows
    $permRows = foreach ($u in ($userPerms.Keys | Sort-Object)) {
        $perms = $userPerms[$u] | Sort-Object -Unique
        $permText = $perms -join ', '
        $cls = ''
        if ($perms -contains 'Full Access')   { $cls += 'full ' }
        if ($perms -contains 'Send As')       { $cls += 'sendas ' }
        if ($perms -contains 'Send on Behalf'){ $cls += 'behalf ' }

        "<tr><td class='user'>$([System.Web.HttpUtility]::HtmlEncode($u))</td><td class='$cls'>$permText</td></tr>"
    }
    if (-not $permRows) { $permRows = "<tr><td colspan='2' class='none'>No permissions assigned</td></tr>" }

    $PermHTML = "<table style='margin-top:12px;font-size:12px;'><tr style='background:#1a5fb4;color:white;'><th>User</th><th>Permissions</th></tr>$($permRows -join '')</table>"

    # ---- Store ----
    $ReportData += [pscustomobject]@{
        SharedMailbox     = $id
        DisplayName       = $name
        ForwardingEnabled = if ($fwd.ForwardingAddress -or $fwd.ForwardingSmtpAddress) { "Yes" } else { "No" }
        ForwardingTo      = $FwdTo
        ForwardingType    = $FwdType
        DeliverAndForward = $fwd.DeliverToMailboxAndForward
        GroupMembersHTML  = $MembersHTML
        PermissionsHTML   = $PermHTML
    }
}
Write-Progress -Activity "Processing" -Completed

# =============================================
# HTML REPORT
# =============================================
$htmlHead = @"
<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>$ReportTitle</title>
<style>
  body{ font-size: 12px;font-family:Segoe UI,Arial,sans-serif;margin:20px;background:#f9f9fb;color:#333}
  h1{color:#66bb6a;text-align:center}
  .summary{text-align:center;margin:15px 0;font-size:14px}
  .mailbox{background:white;margin:20px 0;padding:15px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.1)}
  .mailbox h2{color:#66bb6a;margin:0 0 10px;font-size:16px;border-bottom:2px solid #1a5fb4;padding-bottom:5px}
  table{width:100%;border-collapse:collapse;margin:10px 0;font-size:13px}
  th,td{padding:4px 5px;border-bottom:1px solid #ddd;text-align:left;vertical-align:top}
  th{background:#66bb6a;color:white}
  tr:hover{background:#f5f7fa}
  .yes{color:#d73a49;font-weight:bold}
  .no{color:#6a737d}
  .full{color:#d73a49;font-weight:bold}
  .sendas{color:#d73a49;font-weight:bold}
  .behalf{color:#0366d6}
  .none{color:#6a737d;font-style:italic}
  .footer{text-align:center;margin-top:40px;color:#777;font-size:12px}
</style></head><body>
<h1>$ReportTitle</h1>
<div class="summary">
  <strong>Generated (HKT):</strong> $($NowHKT.ToString('yyyy-MM-dd HH:mm')) |
  <strong>Total Shared Mailboxes:</strong> $total 
</div>
"@

$htmlBody = ""

foreach ($item in $ReportData) {
    $fwdClass = if ($item.ForwardingEnabled -eq 'Yes') { 'yes' } else { 'no' }

    $htmlBody += @"
<div class="mailbox">
  <h2>$([System.Web.HttpUtility]::HtmlEncode($item.DisplayName)) <small style="color:#666">($($item.SharedMailbox))</small></h2>

  <table>
    <tr><th>Forwarding Enabled</th><td class="$fwdClass">$($item.ForwardingEnabled)</td></tr>
    <tr><th>Forwarding To</th><td>$([System.Web.HttpUtility]::HtmlEncode($item.ForwardingTo))</td></tr>
  </table>
"@

    if ($item.GroupMembersHTML) {
        $htmlBody += "<strong>Distribution Group Members:</strong>$($item.GroupMembersHTML)"
    }

    $htmlBody += "<strong>Permissions:</strong>$($item.PermissionsHTML)</div>"
}

$htmlFooter = "<div class='footer'>Report generated on $($NowHKT.ToString('dddd, MMMM dd, yyyy HH:mm')) HKT | Microsoft 365 </div></body></html>"
$html = $htmlHead + $htmlBody + $htmlFooter

# =============================================
# SAVE & EMAIL
# =============================================
$htmlPath = "$env:USERPROFILE\Desktop\SharedMailbox_Audit_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
$html | Out-File -FilePath $htmlPath -Encoding UTF8

Write-Host "`nHTML report saved: $htmlPath" -ForegroundColor Green

# ---- Send Email ----
try {
    Send-MailMessage `
        -From $FromEmail `
        -To $ToEmail `
        -Subject "$ReportTitle $($NowHKT.ToString('yyyy-MM-dd'))" `
        -Body $html `
        -BodyAsHtml `
        -SmtpServer $SmtpServer `
        -Port $SmtpPort `
        -Attachments $htmlPath `
        -ErrorAction Stop

    Write-Host "Email sent to $ToEmail" -ForegroundColor Green
}
catch {
    Write-Warning "Email failed: $($_.Exception.Message)"
}

# ---- Disconnect ----
Write-Host "`nDisconnecting..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Audit complete." -ForegroundColor Green