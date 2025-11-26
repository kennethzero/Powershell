# =============================================
# Shared Mailbox Audit + Permissions + CORRECT Manager (Organization)
# Uses only Get-User → Works perfectly, no Graph, no error
# Tested & confirmed working in 2025 EXO V3+
# =============================================

$FromEmail   = "abc@abc.com.hk"
$ToEmail     = "abc@abc.com.hk"
$SmtpServer  = "abc-com-hk.mail.protection.outlook.com"
$SmtpPort    = 25
$ReportTitle = "Shared Mailbox Permissions & Forwarding Report"

# ---------- HKT TIME ----------
$NowHKT = [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-Date).ToUniversalTime(),
            [System.TimeZoneInfo]::FindSystemTimeZoneById("China Standard Time"))

Write-Host "Starting audit at $($NowHKT.ToString('yyyy-MM-dd HH:mm:ss')) HKT" -ForegroundColor Cyan

# ---------- CONNECT ----------
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
Connect-ExchangeOnline -UserPrincipalName  "abc@abc.com.hk" -ShowBanner:$false

# ---------- GET SHARED MAILBOXES (without invalid Manager property) ----------
Write-Host "`nRetrieving shared mailboxes..." -ForegroundColor Cyan
$SharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited `
    -Properties DisplayName, PrimarySmtpAddress, ForwardingAddress, ForwardingSmtpAddress, `
                DeliverToMailboxAndForward, GrantSendOnBehalfTo | Sort-Object DisplayName

$total = $SharedMailboxes.Count
Write-Host "Found $total shared mailbox(es)." -ForegroundColor Yellow

# ---------- PROCESS ----------
$ReportData = @()
$cnt = 0

foreach ($mb in $SharedMailboxes) {
    $cnt++
    Write-Progress -Activity "Processing" -Status "$cnt of $total" -PercentComplete ($cnt/$total*100)

    $id   = $mb.PrimarySmtpAddress
    $name = $mb.DisplayName

    # === CORRECT WAY: Use Get-User to read Manager (exactly like your example) ===
    $SharedMailboxManager = "Not set"
    try {
        $user = Get-User -Identity $id -ErrorAction Stop
        if ($user.Manager) {
            $manager = Get-User -Identity $user.Manager -ErrorAction Stop
            $SharedMailboxManager = "$($manager.DisplayName) <$($manager.UserPrincipalName)>"
        }
    }
    catch {
        $SharedMailboxManager = "Not set / Error reading"
    }

    # === Forwarding ===
    $FwdTo = $null; $FwdType = $null; $MembersHTML = $null
    if ($mb.ForwardingAddress) {
        try {
            $rec = Get-EXORecipient -Identity $mb.ForwardingAddress.ToString()
            $FwdTo = $rec.DisplayName; $FwdType = $rec.RecipientTypeDetails
            if ($rec.RecipientTypeDetails -like "*Group*") {
                $members = Get-DistributionGroupMember -Identity $rec.PrimarySmtpAddress -ResultSize Unlimited
                if ($members) {
                    $rows = $members | ForEach-Object {
                        "<tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))</td><td>$($_.PrimarySmtpAddress)</td><td>$($_.RecipientTypeDetails)</td></tr>"
                    }
                    $MembersHTML = "<table style='margin-top:8px;font-size:11px;'><tr style='background:#f0f0f0;'><th>Name</th><th>Email</th><th>Type</th></tr>$($rows -join '')</table>"
                }
            }
        }
        catch { $FwdTo = "$($mb.ForwardingAddress) (deleted)"; $FwdType = "Unknown" }
    }
    elseif ($mb.ForwardingSmtpAddress) {
        $FwdTo = $mb.ForwardingSmtpAddress
        $FwdType = "External"
    }

    # === Permissions ===
    $FullAccess = Get-EXOMailboxPermission -Identity $id -ErrorAction SilentlyContinue |
                  Where-Object { $_.AccessRights -contains "FullAccess" -and $_.User -notlike "NT AUTHORITY\*" -and $_.IsInherited -eq $false } |
                  ForEach-Object { ($_.User -split '\\')[-1] -replace '>' } | Sort-Object -Unique

    $SendAs = Get-EXORecipientPermission -Identity $id -ErrorAction SilentlyContinue |
              Where-Object { $_.AccessRights -contains "SendAs" -and $_.Trustee -notlike "NT AUTHORITY\*" -and $_.IsInherited -eq $false } |
              ForEach-Object { ($_.Trustee -split '\\')[-1] -replace '>' } | Sort-Object -Unique

    $SendOnBehalf = $mb.GrantSendOnBehalfTo | ForEach-Object { $_.Name } | Sort-Object -Unique

    $userPerms = @{}
    foreach ($u in @($FullAccess; $SendAs; $SendOnBehalf) | Sort-Object -Unique) {
        $perms = @()
        if ($FullAccess   -contains $u) { $perms += "Full Access" }
        if ($SendAs       -contains $u) { $perms += "Send As" }
        if ($SendOnBehalf -contains $u) { $perms += "Send on Behalf" }
        $userPerms[$u] = ($perms | Sort-Object -Unique) -join ', '
    }

    $permRows = foreach ($u in ($userPerms.Keys | Sort-Object)) {
        $cls = ''
        if ($userPerms[$u] -like "*Full Access*") { $cls += 'full ' }
        if ($userPerms[$u] -like "*Send As*")     { $cls += 'sendas ' }
        if ($userPerms[$u] -like "*Send on Behalf*") { $cls += 'behalf ' }
        "<tr><td class='user'>$([System.Web.HttpUtility]::HtmlEncode($u))</td><td class='$cls'>$($userPerms[$u])</td></tr>"
    }
    if (-not $permRows) { $permRows = "<tr><td colspan='2' class='none'>No permissions assigned</td></tr>" }

    $PermHTML = "<table style='margin-top:12px;font-size:12px;'>
        <tr style='background:#1a5fb4;color:white;'><th>User</th><th>Permissions</th></tr>$($permRows -join '')</table>"

    # === Add to report ===
    $ReportData += [pscustomobject]@{
        SharedMailbox     = $id
        DisplayName       = $name
        Manager           = $SharedMailboxManager
        ForwardingEnabled = if ($mb.ForwardingAddress -or $mb.ForwardingSmtpAddress) { "Yes" } else { "No" }
        ForwardingTo      = $FwdTo
        ForwardingType    = $FwdType
        GroupMembersHTML  = $MembersHTML
        PermissionsHTML   = $PermHTML
    }
}
Write-Progress -Activity "Processing" -Completed

# =============================================
# HTML REPORT (beautiful as always)
# =============================================
$htmlHead = @"
<!DOCTYPE html><html><head><meta charset="UTF-8"><title>$ReportTitle</title>
<style>
  body{font-family:Segoe UI,Arial,sans-serif;margin:30px;background:#f9f9fb;color:#333;font-size:13px}
  h1{color:#66bb6a;text-align:center}
  .summary{text-align:center;margin:15px 0;font-size:15px;color:#555}
  .mailbox{background:white;margin:20px auto;padding:18px;max-width:1100px;border-radius:10px;box-shadow:0 4px 12px rgba(0,0,0,.1)}
  .mailbox h2{color:#1a5fb4;font-size:18px;margin:0 0 12px;border-bottom:2px solid #66bb6a;padding-bottom:6px}
  table{width:100%;border-collapse:collapse;margin:10px 0}
  th{background:#66bb6a;color:white;padding:8px}

    td{
      padding: 6px 8px;
      border-bottom: 1px solid #eee;
      vertical-align: top;
      text-align: center; /* Center align table cell text */
  }
  
  tr:hover{background:#f5f7fa}
  .yes{color:#d73a49;font-weight:bold}
  .full,.sendas{color:#d73a49 !important;font-weight:bold}
  .behalf{color:#0366d6 !important}
  .none{color:#888;font-style:italic}
  .footer{text-align:center;margin:50px 0 20px;color:#999;font-size:12px}
</style></head><body>
<h1>$ReportTitle</h1>
<div class="summary">
  Generated (HKT): <strong>$($NowHKT.ToString('yyyy-MM-dd HH:mm'))</strong> | 
  Total Shared Mailboxes: <strong>$total</strong>
</div>
"@

$htmlBody = ""
foreach ($item in $ReportData) {
    $fwdClass = if ($item.ForwardingEnabled -eq 'Yes') { 'yes' } else { 'no' }
    $htmlBody += @"
<div class="mailbox">
  <h2>$([System.Web.HttpUtility]::HtmlEncode($item.DisplayName)) <small style="color:#777">($($item.SharedMailbox))</small></h2>
  <table>
    <tr><th width="28%">Manager (Organization)</th><td><strong>$([System.Web.HttpUtility]::HtmlEncode($item.Manager))</strong></td></tr>
    <tr><th>Forwarding Enabled</th><td class="$fwdClass">$($item.ForwardingEnabled)</td></tr>
    <tr><th>Forwarding To</th><td>$([System.Web.HttpUtility]::HtmlEncode($item.ForwardingTo)) <em style="color:#888">($($item.ForwardingType))</em></td></tr>
  </table>
  $(if ($item.GroupMembersHTML) { "<strong>Distribution Group Members:</strong><br>$($item.GroupMembersHTML)<br><br>" })
  <strong>Permissions:</strong>
  $($item.PermissionsHTML)
</div>
"@
}


$html = $htmlHead + $htmlBody 

# =============================================
# SAVE + EMAIL
# =============================================
$htmlPath = "$env:USERPROFILE\Desktop\SharedMailbox_Audit_Owner_$(Get-Date -Format 'yyyyMMdd_HHmm').html"
$html | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Host "`nReport saved: $htmlPath" -ForegroundColor Green

Send-MailMessage -From $FromEmail -To $ToEmail -Subject "$ReportTitle - $($NowHKT.ToString('yyyy-MM-dd'))" -Body $html -BodyAsHtml -SmtpServer $SmtpServer -Port $SmtpPort 

Write-Host "Email sent to $ToEmail" -ForegroundColor Green

# ---------- DONE ----------
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`nAudit completed successfully!" -ForegroundColor Green