# Explicitly force-load native Windows UI frameworks at compilation
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# ======================
# USER CONFIGURATION - EDIT THESE VALUES
# ======================

# 1. Your OpenRouter API Key - Get this from https://openrouter.ai/keys
#    You need to create an account and generate an API key
$apiKey = "YOUR_OPENROUTER_API_KEY_HERE"

# 2. OpenRouter API URL - Do not change unless you know what you're doing
$apiURL = "https://openrouter.ai/api/v1"

# 3. Model to use - Choose from available models on OpenRouter
#    Popular free options: "google/gemma-2-9b-it:free", "microsoft/phi-3-mini-128k-instruct:free"
#    Paid options: "anthropic/claude-3.5-sonnet", "openai/gpt-4o", etc.
$modelId = "google/gemma-2-9b-it:free"  # Change as needed

# ======================
# END USER CONFIGURATION
# ======================

# Validate API key is set
if ([string]::IsNullOrEmpty($apiKey) -or $apiKey -eq "YOUR_OPENROUTER_API_KEY_HERE") {
    [System.Windows.MessageBox]::Show("Please set your OpenRouter API key in the script configuration section.", "Configuration Error", "OK", "Error")
    Exit
}

# 1. Hook Into the Running Desktop Outlook Instance
try {
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch {
    [System.Windows.MessageBox]::Show("Classic Outlook must be open and running.", "Outlook Not Found", "OK", "Warning")
    Exit
}

$explorer = $outlook.ActiveExplorer()

# 2. Clean Native Windows GUI Form for Tone Selection
$form = New-Object System.Windows.Forms.Form
$form.Text = "Select Response Tone Strategy"
$form.Size = New-Object System.Drawing.Size(350,220)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20,20)
$label.Size = New-Object System.Drawing.Size(300,20)
$label.Text = "Choose a reply style option:"
$form.Controls.Add($label)

$dropDown = New-Object System.Windows.Forms.ComboBox
$dropDown.Location = New-Object System.Drawing.Point(20,50)
$dropDown.Size = New-Object System.Drawing.Size(290,30)
$dropDown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$dropDown.MaxDropDownItems = 5
[void]$dropDown.Items.Add("1. Check last 24h unread emails")
[void]$dropDown.Items.Add("2. Summarize email")
[void]$dropDown.Items.Add("3. Acknowledge email received")
[void]$dropDown.Items.Add("4. Check last 2 day unread emails")
[void]$dropDown.Items.Add("5. Quick Reply")
$dropDown.SelectedIndex = 0
$form.Controls.Add($dropDown)

$btnOK = New-Object System.Windows.Forms.Button
$btnOK.Location = New-Object System.Drawing.Point(110,110)
$btnOK.Size = New-Object System.Drawing.Size(120,35)
$btnOK.Text = "Generate Draft"
$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $btnOK
$form.Controls.Add($btnOK)

# Display the interface window
$dialogResult = $form.ShowDialog()
if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) { Exit }

$isTriage = $dropDown.SelectedItem -like "1*" -or $dropDown.SelectedItem -like "4*"

# For option 5, collect Reply/ReplyAll/Forward before the progress form appears
$quickReplyMode = $null
if ($dropDown.SelectedItem -like "5*") {
    $subForm = New-Object System.Windows.Forms.Form
    $subForm.Text = "Quick Reply - Action"
    $subForm.Size = New-Object System.Drawing.Size(280,160)
    $subForm.StartPosition = "CenterScreen"
    $subForm.FormBorderStyle = "FixedDialog"
    $subForm.MaximizeBox = $false

    $subLabel = New-Object System.Windows.Forms.Label
    $subLabel.Location = New-Object System.Drawing.Point(20,20)
    $subLabel.Size = New-Object System.Drawing.Size(230,20)
    $subLabel.Text = "Choose action:"
    $subForm.Controls.Add($subLabel)

    $subDrop = New-Object System.Windows.Forms.ComboBox
    $subDrop.Location = New-Object System.Drawing.Point(20,48)
    $subDrop.Size = New-Object System.Drawing.Size(220,30)
    $subDrop.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$subDrop.Items.Add("Reply")
    [void]$subDrop.Items.Add("Reply All")
    [void]$subDrop.Items.Add("Forward")
    $subDrop.SelectedIndex = 0
    $subForm.Controls.Add($subDrop)

    $subBtn = New-Object System.Windows.Forms.Button
    $subBtn.Location = New-Object System.Drawing.Point(80,90)
    $subBtn.Size = New-Object System.Drawing.Size(100,30)
    $subBtn.Text = "Continue"
    $subBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $subForm.AcceptButton = $subBtn
    $subForm.Controls.Add($subBtn)

    if ($subForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { Exit }
    $quickReplyMode = $subDrop.SelectedItem
}

# Marquee progress popup shown while scanning Outlook and calling the gateway
$progressForm = New-Object System.Windows.Forms.Form
$progressForm.Text = "Working"
$progressForm.Size = New-Object System.Drawing.Size(300,110)
$progressForm.StartPosition = "CenterScreen"
$progressForm.FormBorderStyle = "FixedDialog"
$progressForm.ControlBox = $false
$progressForm.TopMost = $true

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Text = "Generating response..."
$progressLabel.AutoSize = $true
$progressLabel.Location = New-Object System.Drawing.Point(20,15)
$progressForm.Controls.Add($progressLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$progressBar.MarqueeAnimationSpeed = 30
$progressBar.Location = New-Object System.Drawing.Point(20,45)
$progressBar.Size = New-Object System.Drawing.Size(250,25)
$progressForm.Controls.Add($progressBar)

$progressForm.Show()
$progressForm.Refresh()
[System.Windows.Forms.Application]::DoEvents()

if ($isTriage) {
    $is24h = $dropDown.SelectedItem -like "1*"
    $inbox = $outlook.Session.GetDefaultFolder(6) # olFolderInbox
    $cutoff = if ($is24h) { (Get-Date).AddHours(-24).ToString("MM/dd/yyyy HH:mm") } else { (Get-Date).AddDays(-2).ToString("MM/dd/yyyy HH:mm") }
    $items = $inbox.Items.Restrict("[Unread] = True AND [ReceivedTime] >= '$cutoff'")

    $emailBlocks = foreach ($item in $items) {
        if ($item.MessageClass -ne "IPM.Note") { continue }
        "From: $($item.SenderName)`nSubject: $($item.Subject)`nReceived: $($item.ReceivedTime)`nBody: $($item.Body.Substring(0, [Math]::Min(1000, $item.Body.Length)))`n---"
    }

    $timeLabel = if ($is24h) { "last 24 hours" } else { "last 2 days" }
    if (-not $emailBlocks) {
        $progressForm.Close()
        [System.Windows.MessageBox]::Show("No unread emails found in the $timeLabel.", "Nothing to Triage", "OK", "Information")
        Exit
    }

    $tonePrompt = "Please scan my unread emails from the $timeLabel. Sort them into these exact priority buckets: Urgent - Needs a reply today (Direct questions, client blocks, time-sensitive tasks); Important - Needs attention this week (Updates, non-blocking requests); FYI / Low Priority - No action needed (Newsletters, notifications). List the Urgent section first. For each urgent email, provide the sender, subject line, a 1-sentence summary, and a suggested quick reply draft."
    $promptContent = "$tonePrompt`n`nEmails:`n$($emailBlocks -join "`n")"
    $maxTokens = 4096
} else {
    $explorer = $outlook.ActiveExplorer()
    if ($explorer.Selection.Count -eq 0) {
        $progressForm.Close()
        [System.Windows.MessageBox]::Show("Please select an email in Outlook first.", "No Selection", "OK", "Exclamation")
        Exit
    }
    $mailItem = $explorer.Selection.Item(1)
    if ($mailItem.MessageClass -ne "IPM.Note") {
        $progressForm.Close()
        [System.Windows.MessageBox]::Show("Selected item is not a standard email.", "Invalid Selection", "OK", "Exclamation")
        Exit
    }

    # Map selected string index directly to Claude prompts
    switch -Wildcard ($dropDown.SelectedItem) {
        "2*" { $tonePrompt = "Summarize this email concisely in bullet points" }
        "3*" { $tonePrompt = "Write a brief corporate acknowledgment message confirming you have received the email and are looking into it." }
        "5*" { $tonePrompt = "Write a short, direct reply in 2-3 sentences. Format it for easy reading: use short paragraphs (1-2 sentences each), a blank line between paragraphs, and a clear opening line that states the main point immediately. No filler, no pleasantries, no subject line, no sign-off." }
        Default { $tonePrompt = "Write a polite, professional corporate reply to the following email." }
    }

    $cleanBody = $mailItem.Body
    $promptContent = "$tonePrompt Do not include placeholders, signoffs, or subject lines. Just give me the text body. Email content: $cleanBody"
    $maxTokens = 1024
}

# 3. Construct OpenRouter API Payload
$jsonPayload = @{
    model = $modelId
    max_tokens = $maxTokens
    messages = @(
        @{
            role = "user"
            content = $promptContent
        }
    )
} | ConvertTo-Json -Depth 5 -Compress

# 4. Execute OpenRouter API Request
# WebClient with Encoding forced to UTF8: PS5.1's Invoke-RestMethod mis-decodes the
# response body (mangles bullets/em-dashes/curly quotes) when the server doesn't pin
# charset=utf-8 on Content-Type. Setting .Encoding fixes decoding for both request
# and response.
# UploadStringTaskAsync + a DoEvents poll loop (not the plain synchronous
# UploadString) — blocking the UI thread on the request leaves nobody to pump
# Windows messages, so the progress form and the whole hidden host report
# "Not Responding" until the call finishes.
$webClient = New-Object System.Net.WebClient
$webClient.Encoding = [System.Text.Encoding]::UTF8
$webClient.Headers.Add("Content-Type", "application/json")
$webClient.Headers.Add("Authorization", "Bearer $apiKey")
$webClient.Headers.Add("HTTP-Referer", "https://github.com/yourusername/claude-to-outlook")  # Optional but recommended
$webClient.Headers.Add("X-Title", "ClaudeToOutlook")  # Optional but recommended

try {
    $task = $webClient.UploadStringTaskAsync("$apiURL/chat/completions", "POST", $jsonPayload)
    while (-not $task.IsCompleted) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 50
    }

    if ($task.IsFaulted) { throw $task.Exception.InnerException }

    $responseText = $task.Result
    $response = $responseText | ConvertFrom-Json
    $generatedText = $response.choices[0].message.content
    $progressForm.Close()

    if ($isTriage) {
        # Show the triage report in a new mail window (nothing is sent)
        $reportMail = $outlook.CreateItem(0) # olMailItem
        $reportMail.Subject = "Inbox Triage - $((Get-Culture).TextInfo.ToTitleCase($timeLabel))"
        $reportMail.Body = $generatedText
        $reportMail.Display()
    } else {
        # Create the draft using the appropriate Outlook action
        $replyDraft = switch ($quickReplyMode) {
            "Reply All" { $mailItem.ReplyAll() }
            "Forward"   { $mailItem.Forward() }
            default     { $mailItem.Reply() }
        }

        # Convert plain text to HTML with Aptos Body 12pt, insert at top of body.
        # Outlook already provides its own separator line between the reply area and
        # the quoted thread, so no extra <hr> is needed here.
        $htmlText = ($generatedText -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;') `
                    -replace "`r?`n`r?`n", '</p><p>' `
                    -replace "`r?`n", '<br>'
        $htmlWrapped = "<div style='font-family:Aptos,Calibri,sans-serif;font-size:12pt'><p>$htmlText</p></div>"
        $replyDraft.HTMLBody = $replyDraft.HTMLBody -replace '(<body[^>]*>)', "`$1$htmlWrapped"
        $replyDraft.Display()
    }
} catch [System.Net.WebException] {
    $progressForm.Close()
    $statusCode = [int]$_.Exception.Response.StatusCode
    $streamReader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream(), $webClient.Encoding)
    $errorDetails = $streamReader.ReadToEnd()
    if ([string]::IsNullOrWhiteSpace($errorDetails)) { $errorDetails = "(empty response body)" }
    [System.Windows.MessageBox]::Show("HTTP $statusCode`nAPI Error Details:`n$errorDetails", "Network Connection Error", "OK", "Error")
} catch {
    $progressForm.Close()
    [System.Windows.MessageBox]::Show("Network Error Details:`n$_", "Network Connection Error", "OK", "Error")
}