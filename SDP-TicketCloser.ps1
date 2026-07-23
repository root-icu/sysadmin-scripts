#region Script Settings
#<ScriptSettings xmlns="http://tempuri.org/ScriptSettings.xsd">
#  <ScriptPackager>
#    <process>powershell.exe</process>
#    <arguments />
#    <extractdir>%TEMP%</extractdir>
#    <files />
#    <usedefaulticon>true</usedefaulticon>
#    <showinsystray>false</showinsystray>
#    <altcreds>false</altcreds>
#    <efs>true</efs>
#    <ntfs>true</ntfs>
#    <local>false</local>
#    <abortonfail>true</abortonfail>
#    <product />
#    <version>2.0.0.0</version>
#    <versionstring />
#    <comments>Rebuilt to use ServiceDesk Plus REST API v3 (on-premises) with API-key auth and error handling.</comments>
#    <company />
#    <includeinterpreter>false</includeinterpreter>
#    <forcecomregistration>false</forcecomregistration>
#    <consolemode>false</consolemode>
#    <EnableChangelog>false</EnableChangelog>
#    <AutoBackup>false</AutoBackup>
#    <snapinforce>false</snapinforce>
#    <snapinshowprogress>false</snapinshowprogress>
#    <snapinautoadd>2</snapinautoadd>
#    <snapinpermanentpath />
#    <cpumode>1</cpumode>
#    <hidepsconsole>false</hidepsconsole>
#  </ScriptPackager>
#</ScriptSettings>
#endregion

#region Constructor
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
#endregion

# ============================================================================
#  NOTES ON THIS VERSION
# ----------------------------------------------------------------------------
#  - Talks to ServiceDesk Plus REST API v3 (on-premises) instead of the old
#    Servlet form-post API. If you're on SDP Cloud (SDPOD) or MSP, the
#    endpoint paths are the same shape but the auth header differs slightly
#    (Cloud uses OAuth bearer tokens, not a static TECHNICIAN_KEY) - flag
#    that to me and I'll adjust Invoke-SdpApi accordingly.
#  - Auth: Generate a Technician Key from
#      Admin > Developer Space > API > Documentation  (on-prem)
#    and paste it into the "API Key" field. No username/password/domain
#    is transmitted anymore.
#  - Every ticket goes through: AddNote -> CloseRequest. Both calls are
#    wrapped in try/catch. If AddNote fails, the ticket is SKIPPED (not
#    closed) so you never end up with a silently-unclosed ticket that
#    looks handled. Every outcome (success or failure, with the actual
#    server error) is written to the on-screen log AND to a session log
#    file next to the script.
# ============================================================================

#region Form Creation
$frmMain = New-Object System.Windows.Forms.Form
$frmMain.ClientSize = New-Object System.Drawing.Size(560, 520)
$frmMain.Text = "Service Desk Plus Ticket Closer (REST API v3)"
$frmMain.StartPosition = "CenterScreen"
$frmMain.FormBorderStyle = "FixedDialog"
$frmMain.MaximizeBox = $false

# --- Connection setup group -------------------------------------------------
$gbxConnectionSetup = New-Object System.Windows.Forms.GroupBox
$gbxConnectionSetup.Location = New-Object System.Drawing.Point(12, 12)
$gbxConnectionSetup.Size = New-Object System.Drawing.Size(536, 90)
$gbxConnectionSetup.Text = "Service Desk Connection Setup"

$lblURL = New-Object System.Windows.Forms.Label
$lblURL.Location = New-Object System.Drawing.Point(10, 25)
$lblURL.Size = New-Object System.Drawing.Size(90, 20)
$lblURL.Text = "SDP Base URL:"

$txtURL = New-Object System.Windows.Forms.TextBox
$txtURL.Location = New-Object System.Drawing.Point(105, 22)
$txtURL.Size = New-Object System.Drawing.Size(420, 20)
$txtURL.Text = "https://"

$lblApiKey = New-Object System.Windows.Forms.Label
$lblApiKey.Location = New-Object System.Drawing.Point(10, 55)
$lblApiKey.Size = New-Object System.Drawing.Size(90, 20)
$lblApiKey.Text = "API Key:"

$txtApiKey = New-Object System.Windows.Forms.TextBox
$txtApiKey.Location = New-Object System.Drawing.Point(105, 52)
$txtApiKey.Size = New-Object System.Drawing.Size(420, 20)
$txtApiKey.UseSystemPasswordChar = $true

$gbxConnectionSetup.Controls.AddRange(@($lblURL, $txtURL, $lblApiKey, $txtApiKey))

# --- Ticket close info group -------------------------------------------------
$gbxTicketCloseInfo = New-Object System.Windows.Forms.GroupBox
$gbxTicketCloseInfo.Location = New-Object System.Drawing.Point(12, 110)
$gbxTicketCloseInfo.Size = New-Object System.Drawing.Size(536, 130)
$gbxTicketCloseInfo.Text = "Ticket Close Information"

$lblNotes = New-Object System.Windows.Forms.Label
$lblNotes.Location = New-Object System.Drawing.Point(10, 25)
$lblNotes.Size = New-Object System.Drawing.Size(90, 20)
$lblNotes.Text = "Note text:"

$txtNotes = New-Object System.Windows.Forms.TextBox
$txtNotes.Location = New-Object System.Drawing.Point(105, 22)
$txtNotes.Size = New-Object System.Drawing.Size(420, 20)
$txtNotes.Text = "Ticket was closed due to inactivity."

$lblResolution = New-Object System.Windows.Forms.Label
$lblResolution.Location = New-Object System.Drawing.Point(10, 52)
$lblResolution.Size = New-Object System.Drawing.Size(90, 20)
$lblResolution.Text = "Closure comment:"

$txtResolution = New-Object System.Windows.Forms.TextBox
$txtResolution.Location = New-Object System.Drawing.Point(105, 49)
$txtResolution.Size = New-Object System.Drawing.Size(420, 20)
$txtResolution.Text = "Ticket was closed due to inactivity."

$lblCSTicketNumbers = New-Object System.Windows.Forms.Label
$lblCSTicketNumbers.Location = New-Object System.Drawing.Point(10, 80)
$lblCSTicketNumbers.Size = New-Object System.Drawing.Size(420, 18)
$lblCSTicketNumbers.Text = "Comma-separated ticket (request) IDs:"

$txtTicketNumbers = New-Object System.Windows.Forms.TextBox
$txtTicketNumbers.Location = New-Object System.Drawing.Point(10, 100)
$txtTicketNumbers.Size = New-Object System.Drawing.Size(515, 20)

$gbxTicketCloseInfo.Controls.AddRange(@($lblNotes, $txtNotes, $lblResolution, $txtResolution, $lblCSTicketNumbers, $txtTicketNumbers))

# --- Log / results box -------------------------------------------------------
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Location = New-Object System.Drawing.Point(12, 245)
$lblLog.Size = New-Object System.Drawing.Size(200, 18)
$lblLog.Text = "Results:"

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(12, 265)
$txtLog.Size = New-Object System.Drawing.Size(536, 195)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(12, 465)
$progressBar.Size = New-Object System.Drawing.Size(536, 18)

# --- Buttons -------------------------------------------------------------
$butClose = New-Object System.Windows.Forms.Button
$butClose.Location = New-Object System.Drawing.Point(377, 490)
$butClose.Size = New-Object System.Drawing.Size(107, 25)
$butClose.Text = "Close Tickets"
$butClose.add_Click({ ButCloseClick })

$butExit = New-Object System.Windows.Forms.Button
$butExit.Location = New-Object System.Drawing.Point(490, 490)
$butExit.Size = New-Object System.Drawing.Size(58, 25)
$butExit.Text = "Exit"
$butExit.add_Click({ $frmMain.Close() })

$frmMain.Controls.AddRange(@($gbxConnectionSetup, $gbxTicketCloseInfo, $lblLog, $txtLog, $progressBar, $butClose, $butExit))
$frmMain.AcceptButton = $butClose
#endregion

#region Helper functions
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    $txtLog.AppendText("$line`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    if ($script:LogFilePath) {
        Add-Content -Path $script:LogFilePath -Value $line -ErrorAction SilentlyContinue
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Test-InputsValid {
    if ([string]::IsNullOrWhiteSpace($txtURL.Text) -or $txtURL.Text -eq "https://") {
        [System.Windows.Forms.MessageBox]::Show("Enter the SDP base URL.", "Missing info", "OK", "Warning") | Out-Null
        return $false
    }
    if ($txtURL.Text -notmatch '^https?://') {
        [System.Windows.Forms.MessageBox]::Show("The SDP URL must start with http:// or https://", "Invalid URL", "OK", "Warning") | Out-Null
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($txtApiKey.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Enter the API key (Technician Key).", "Missing info", "OK", "Warning") | Out-Null
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($txtTicketNumbers.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Enter at least one ticket ID.", "Missing info", "OK", "Warning") | Out-Null
        return $false
    }
    return $true
}

# Central REST call wrapper: returns @{ Success = $bool; Data = ...; Error = "..." }
function Invoke-SdpApi {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][ValidateSet("Get", "Post", "Put", "Delete")][string]$Method,
        [string]$InputDataJson
    )
    $headers = @{ "TECHNICIAN_KEY" = $txtApiKey.Text }
    $params = @{
        Uri             = $Uri
        Method          = $Method
        Headers         = $headers
        UseBasicParsing = $true
        ErrorAction     = "Stop"
    }
    if ($InputDataJson) {
        $body = "input_data=" + [System.Uri]::EscapeDataString($InputDataJson)
        $params["Body"] = $body
        $params["ContentType"] = "application/x-www-form-urlencoded"
    }
    try {
        $response = Invoke-RestMethod @params
        # SDP REST v3 wraps status in response_status
        $status = $response.response_status
        if ($status -and $status.status -eq "success") {
            return @{ Success = $true; Data = $response }
        }
        elseif ($status) {
            $msg = if ($status.messages) { ($status.messages | ForEach-Object { $_.message }) -join "; " } else { "Unknown API error" }
            return @{ Success = $false; Error = $msg; Data = $response }
        }
        else {
            return @{ Success = $true; Data = $response }
        }
    }
    catch {
        $errDetail = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            try {
                $parsed = $_.ErrorDetails.Message | ConvertFrom-Json
                if ($parsed.response_status.messages) {
                    $errDetail = ($parsed.response_status.messages | ForEach-Object { $_.message }) -join "; "
                }
            }
            catch { }
        }
        return @{ Success = $false; Error = $errDetail }
    }
}

function Add-SdpNote {
    param([string]$BaseUrl, [string]$TicketId, [string]$NoteText)
    $uri = "$BaseUrl/api/v3/requests/$TicketId/notes"
    $payload = @{
        request_note = @{
            description       = $NoteText
            show_to_requester = $false
        }
    } | ConvertTo-Json -Depth 5 -Compress
    return Invoke-SdpApi -Uri $uri -Method Post -InputDataJson $payload
}

function Close-SdpRequest {
    param([string]$BaseUrl, [string]$TicketId, [string]$ClosureComment)
    $uri = "$BaseUrl/api/v3/requests/$TicketId"
    $payload = @{
        request = @{
            status        = @{ name = "Closed" }
            closure_info  = @{ closure_comments = $ClosureComment }
        }
    } | ConvertTo-Json -Depth 5 -Compress
    return Invoke-SdpApi -Uri $uri -Method Put -InputDataJson $payload
}
#endregion

#region Event Handlers
function ButCloseClick {
    if (-not (Test-InputsValid)) { return }

    $baseUrl = $txtURL.Text.TrimEnd("/")
    $tickets = $txtTicketNumbers.Text -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

    if ($tickets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No valid ticket IDs found.", "Nothing to do", "OK", "Warning") | Out-Null
        return
    }

    $nonNumeric = $tickets | Where-Object { $_ -notmatch '^\d+$' }
    if ($nonNumeric.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("These entries are not valid ticket IDs: $($nonNumeric -join ', ')", "Invalid ticket ID", "OK", "Warning") | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will add a note and CLOSE $($tickets.Count) ticket(s)`r`n$($tickets -join ', ')`r`n`r`nContinue?",
        "Confirm", "YesNo", "Question")
    if ($confirm -ne "Yes") { return }

    # session log file next to the script
    $script:LogFilePath = Join-Path -Path $PSScriptRoot -ChildPath ("SDP-TicketCloser-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))

    $butClose.Enabled = $false
    $butExit.Enabled  = $false
    $progressBar.Minimum = 0
    $progressBar.Maximum = $tickets.Count
    $progressBar.Value = 0

    $successCount = 0
    $failCount = 0

    foreach ($ticket in $tickets) {
        Write-Log "Processing ticket $ticket ..."

        $noteResult = Add-SdpNote -BaseUrl $baseUrl -TicketId $ticket -NoteText $txtNotes.Text
        if (-not $noteResult.Success) {
            Write-Log "Ticket $ticket - FAILED to add note: $($noteResult.Error). Ticket NOT closed." "ERROR"
            $failCount++
            $progressBar.Value++
            continue
        }
        Write-Log "Ticket $ticket - note added." "OK"

        $closeResult = Close-SdpRequest -BaseUrl $baseUrl -TicketId $ticket -ClosureComment $txtResolution.Text
        if (-not $closeResult.Success) {
            Write-Log "Ticket $ticket - FAILED to close: $($closeResult.Error)" "ERROR"
            $failCount++
        }
        else {
            Write-Log "Ticket $ticket - closed successfully." "OK"
            $successCount++
        }
        $progressBar.Value++
    }

    Write-Log "Done. Success: $successCount, Failed: $failCount." "SUMMARY"
    [System.Windows.Forms.MessageBox]::Show(
        "Finished`r`nClosed: $successCount`r`nFailed: $failCount`r`n`r`nLog saved to`r`n$($script:LogFilePath)",
        "Ticket closing complete", "OK", "Information") | Out-Null

    $butClose.Enabled = $true
    $butExit.Enabled  = $true
}
#endregion

#region Event Loop
function Main {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::Run($frmMain)
}

Main
#endregion
