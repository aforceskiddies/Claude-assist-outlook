# Claude-assist-outlook
powershell script to use claude code to read email and write summary, draft reply, etc

# ClaudeToOutlook Public Edition

A PowerShell script that integrates with Microsoft Outlook to generate AI-powered email replies, summaries, and triage reports using OpenRouter's API.

## Features

- **Email Triage**: Scan and prioritize unread emails from the last 24 hours or 2 days
- **Email Summarization**: Get concise bullet-point summaries of selected emails
- **Acknowledgment Drafts**: Generate brief corporate acknowledgment messages
- **Smart Replies**: Create contextual replies in 2-3 sentences with clean formatting
- **Multiple Response Actions**: Choose Reply, Reply All, or Forward for quick replies
- **Outlook Integration**: Works directly with classic Outlook desktop application
- **Progress Indicators**: Visual feedback during API calls
- **Privacy Focused**: Uses your own OpenRouter API key - no data sharing

## Prerequisites

- Windows operating system
- Microsoft Outlook (classic version, not New Outlook/WebView2)
- PowerShell 5.1 or later
- An OpenRouter API key (free tier available at [openrouter.ai](https://openrouter.ai))

## Installation

1. **Get an OpenRouter API Key**:
   - Visit [https://openrouter.ai/keys](https://openrouter.ai/keys)
   - Sign up for a free account
   - Generate a new API key

2. **Download the Script**:
   - Download `claudeToOutlookPublic.ps1` from this repository

3. **Configure the Script**:
   - Open `claudeToOutlookPublic.ps1` in a text editor
   - Replace `YOUR_OPENROUTER_API_KEY_HERE` with your actual OpenRouter API key
   - Optionally change the `$modelId` variable to use a different AI model
   - Save the file

4. **Create a Desktop Shortcut (Optional)**:
   - Right-click on your Desktop → New → Shortcut
   - Set the target to:
     ```
     powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "FULL_PATH_TO\claudeToOutlookPublic.ps1"
     ```
   - Name the shortcut (e.g., "ClaudeToOutlook")
   - Right-click the shortcut → Properties → Shortcut key
   - Set your preferred hotkey (e.g., Ctrl+Alt+M)

## Usage

1. Ensure classic Outlook is running
2. Select an email in Outlook (not required for triage options 1 & 4)
3. Run the script (via double-click, shortcut, or PowerShell)
4. Choose an option from the dropdown:
   - **Option 1**: Check last 24h unread emails (triage report)
   - **Option 2**: Summarize selected email
   - **Option 3**: Acknowledge selected email received
   - **Option 4**: Check last 2 day unread emails (triage report)
   - **Option 5**: Quick Reply (select Reply/Reply All/Forward in secondary dialog)
5. Click "Generate Draft"
6. Wait for the "Working..." progress popup to complete
7. Review the generated draft in Outlook and send when ready

## Available Models

You can change the `$modelId` variable in the script to use different AI models:

**Free Models** (good for testing):
- `google/gemma-2-9b-it:free`
- `microsoft/phi-3-mini-128k-instruct:free`
- `meta-llama/llama-3-8b-instruct:free`

**Paid Models** (higher quality):
- `anthropic/claude-3.5-sonnet` (default)
- `openai/gpt-4o`
- `anthropic/claude-3-opus`
- `google/gemini-pro-1.5`

Visit [openrouter.ai/models](https://openrouter.ai/models) for the complete list.

## How It Works

1. **Outlook Integration**: Uses COM automation to connect to the running Outlook instance
2. **Email Processing**: Extracts email content based on your selection or triage criteria
3. **Prompt Engineering**: Creates specific prompts for each function (summarize, acknowledge, reply, triage)
4. **API Call**: Sends request to OpenRouter with your API key
5. **Response Handling**: Formats the AI response and inserts it into an Outlook draft
6. **HTML Conversion**: Converts plain text to properly formatted HTML for Outlook

## Privacy & Security

- Your OpenRouter API key is stored only in the script file on your local machine
- Email content is sent only to OpenRouter for processing
- No data is stored, logged, or shared with third parties
- The script does not send emails automatically - only creates drafts for your review
- Consider securing the script file if your API key is sensitive

## Troubleshooting

**Common Issues:**

1. **"Outlook Not Found" Error**:
   - Make sure classic Outlook is running (not New Outlook)
   - The script requires the desktop Outlook application

2. **API Authentication Errors**:
   - Verify your OpenRouter API key is correctly entered
   - Check that you have sufficient credits/quota on OpenRouter
   - Ensure the key hasn't been expired or revoked

3. **Network/Connection Errors**:
   - Verify your internet connection
   - Check if your firewall is blocking the script
   - OpenRouter API endpoint: `https://openrouter.ai/api/v1`

4. **Script Execution Issues**:
   - Make sure PowerShell execution policy allows script running
   - Try running PowerShell as Administrator if needed
   - Check for any typos in the script after editing

5. **Blank or Unexpected Responses**:
   - Try a different model (some free models have usage limits)
   - Check your OpenRouter dashboard for usage statistics
   - Verify the prompt content is being generated correctly

## Customization

Feel free to modify the script to suit your needs:

- Adjust prompt text in the switch statement for different response styles
- Change timeout values or progress bar settings
- Modify the HTML formatting for replies
- Add additional tone options or email processing features

## License

This script is provided as-is for educational and personal use.
By using this script, you agree to comply with OpenRouter's terms of service and
Microsoft's software licensing terms.

## Acknowledgments

- Based on the original ClaudeReply.ps1 for internal TI use
- Uses OpenRouter for access to various AI models
- Built with PowerShell and Windows Forms for native Outlook integration

## Support

For issues or questions, please open an issue in the GitHub repository.
Check the OpenRouter documentation at [https://openrouter.ai/docs](https://openrouter.ai/docs) for API details.
