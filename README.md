# PingCastle Automation Script

This PowerShell script automates the execution of [PingCastle](https://www.pingcastle.com/) for Active Directory health checks, saves the results, compares with the previous run, logs the changes, and emails a report.

---

## 📁 Folder Structure

Place all files in a base folder like:

<pre><code>C:\PingCastle-Automation\
│
├── PingCastle_AutoReport_FULL_EN.ps1
│
├── PingCastle\
│   └── PingCastle.exe
│
└── Reports\
    ├── ad_hc_<domain>_<yyyyMMdd>.xml
    ├── ad_hc_<domain>_<yyyyMMdd>.html
    └── Audit-PingCastle_Run-<yyyyMMdd>.log</code></pre>


- `Reports\` is created automatically if it doesn't exist.
- `PingCastle.exe` must be manually downloaded and placed inside the `PingCastle` subfolder.

---

## ⚙️ Parameters to Edit

At the top of the script:

```powershell
$smtpServer = "smtp.server.local"
$smtpPort = 25
$smtpFrom = "sender@domain.com"
$smtpTo = "recipient1@domain.com; recipient2@domain.com"
```
You can also customize the email subject and body at the bottom of the script.

## 📤 What the Script Does
Runs PingCastle with a full healthcheck against the local domain.

Renames and moves the generated XML/HTML reports to the Reports folder.

Compares the current report to the previous one.

Logs:

New risks

Removed risks

Risk detail changes

Sends an email with the current report and comparison log attached.


## ✅ Requirements
PowerShell 5.1 or later

PingCastle.exe (must be downloaded separately)

Outbound SMTP access

## 📌 Notes
Script must be executed with domain permissions.

No external dependencies required.

Works entirely offline after downloading PingCastle.

## 🛠️ Recommended Improvements
Add param() support for CLI execution.

Secure credentials handling.

Integrate optional webhook notification (Teams/Slack/Matter).

Add retention policy for report cleanup.

