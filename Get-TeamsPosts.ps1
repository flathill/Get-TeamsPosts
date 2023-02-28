#####
#
# By executing the command with a user name (UPN),
# the contents of the postings of the team to which
# the user belongs are retrieved and output to an HTML file.
# When executing the command, a pop-up will ask for authentication,
# so authenticate with the Teams global administrator.
# It works with PowerShell 7.
#
# Code Based on: https://qiita.com/seilian/items/225b1fe012d502bd4172
# CSS  Based on: https://nelog.jp/line-bolloon-css
#
# Get-TeamsPosts.ps1
#   Seiichirou Hiraoka <seiichirou.hiraoka@gmail.com>
#     Initial Version: 2023/02/24
#
#   Usage: Get-TeamsPosts.ps1 -UserName UPN [-Verbose] [-Debug]
#
# Preparing the script to work
#   Install-Module Microsoft.Graph.Authentication
#   Install-Module Microsoft.Graph.Teams
#   Install-Module MicrosoftTeams
#
#####

# Accepts UserName(UPN) as a parameter
# Debug, Verbose if necessary
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$UserName
)
# Function to check if an e-mail address is in the correct format
function Validate-UserName($UserName) {
    return $UserName -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
}

# Function to take the result of Get-Date and convert it to JST time zone
function Change-Timezone($DateTime){
    # Convert format by adding 9 hours JST to DateTime
    $JST = (Get-Date $DateTime.AddHours(9) -format "yyyy/MM/dd HH:mm:ss")
    
    Write-Output $JST
}

# Determine if the email address is correct
$isValid = Validate-UserName $UserName

# Output the contents of a variable for debugging
Write-Verbose "UserName: $UserName"
Write-Verbose "IsValid: $isValid"

# If there is no UPN
if([string]::IsNullOrEmpty($UserName)) {
    Write-Host "Execute with UPN." 
    Exit 1
}

# Import Microsoft Teams Module
Import-Module Microsoft.Graph.Teams

# Connect to Teams Admin Center
Connect-MicrosoftTeams

# Connet to Microsoft Graph
Connect-MgGraph -Scopes @("ChannelMessage.Read.All")

# Get the ID and name of the team to which the user belongs
$Teams = Get-Team -User $UserName | Select-Object GroupId, DisplayName

# Processed for each team
foreach ($Team in $Teams) {

    # Show team ID and name
    Write-Verbose "Team Name: $($Team.DisplayName)" -Verbose
    Write-Verbose "Team ID: $($Team.GroupId)" -Verbose

    # Get IDs and names of all channels on the team
    $Channels = Get-TeamChannel -GroupId $Team.GroupId | Select-Object Id, DisplayName

    # Proceeded for each channels
    foreach ($Channel in $Channels) {

        # Show channel ID and name
        Write-Verbose "Channel Name: $($Channel.DisplayName)" -Verbose
        Write-Verbose "Channel ID: $($Channel.Id)" -Verbose

        # Create variables to output postings to a file in HTML format
        $Html = "<html>"
        $Html += "<head>"
        $Html += "<title>$($Team.DisplayName) - $($Channel.DisplayName)</title>"
        $Html += "<link rel=""stylesheet"" href=""style.css"" type=""text/css"">"
        $Html += "</head>"
        $Html += "<body>"
        $Html += "<div style=""background-color: #7897C5;"">"

        # Output team and channel names and IDs
        $Html += "<p>"
        $Html += "<b>Team Name: $($Team.DisplayName)</b><br>"
        $Html += "<b>Team ID: $($Team.GroupId)</b><br>"
        $Html += "<b>Channel Name: $($Channel.DisplayName)</b><br>"
        $Html += "<b>Channel ID: $($Channel.Id)</b><br>"
        $Html += "</p>"
            
        # Retrieve channel postings
        $Messages = Get-MgTeamChannelMessage -TeamId $Team.GroupId -ChannelId $Channel.Id

        # Process each submission
        foreach ($Message in $Messages) {

            # Change timezone
            $JST = Change-Timezone($Message.CreatedDateTime)

            # Add the date and time of the post, the name of the submitter, and the text (HTML) to the variable
            $Html += "<hr>"
            $Html += "<div class=""left_balloon"">"
            $Html += "<p>"
            $Html += "<b>件名:$($Message.Subject)</b><br>"
            $Html += "$($JST)) by <i>$($Message.From.User.DisplayName)</i><br>"
            $Html += "</p>"
            $Html += "$($Message.Body.Content)"

            # Process attachments
            # $Message.attachments is not Null
            if(-not [string]::IsNullOrEmpty($Message.attachments)) {
                $Message.attachments | ConvertTo-Json

                $Html += "<p>"

                foreach ($attachment in $Message.attachments) {
                    $Html += "<a href=""$($attachment.ContentUrl)"">$($attachment.name)</a><br>"
                }

                $Html += "</p>"
            }

            # Process reactions
            if(-not [string]::IsNullOrEmpty($Message.Reactions)) {
                $Message.Reactions | ConvertTo-Json

                $Html += "<p>"
                $Html += "<b>リアクション</b><br>"

                foreach ($reaction in $Message.Reactions) {
                    # Change timezone
                    $JST = Change-Timezone($reaction.CreatedDateTime)

                    $Html += "$($JST) $($reaction.ReactionType) by $($reaction.User.DisplayName)<br>"
                }
                
                $Html += "</p>"
            }

            $Html += "</div>"
            $Html += "<br class=""clear_balloon""/>"

            # Get replies
            $replies = Get-MgTeamChannelMessageReply -TeamId $Team.GroupId -ChannelId $Channel.Id -ChatMessageId $Message.Id -All -PageSize 50
            
            # Process replies
            # $replies is not Null)
            
            if([string]::IsNullOrEmpty($replies)) {
                continue
            }

            Write-Host "Dump replies"
            # Start Reply
            
            # Sort array in reverse order
            [Array]::Reverse($replies)

            # Process each reply content
            foreach ($reply in $replies) {
                if ($reply.messageType -ne "message") {
                    continue
                }
                
                $reply | ConvertTo-Json
                $Html += "<div class=""right_balloon"">"
                
                # Change timezone
                $JST = Change-Timezone($reply.CreatedDateTime)

                # Add the date and time of the post, the name of the submitter, and the text (HTML) to the variable
                $Html += "<p><b>$($JST)</b> by <i>$($reply.From.User.DisplayName)</i></p>"
                $Html += "$($reply.Body.Content)"
                
                # Process attachments
                
                if(-not [string]::IsNullOrEmpty($reply.attachments)) {
                    $reply.attachments | ConvertTo-Json
                    
                    $Html += "<p>"
                    foreach ($attachment in $reply.attachments) {
                        $Html += "<a href=""$($attachment.ContentUrl)"">$($attachment.name)</a><br>"
                    }
                    $Html += "</p>"
                }
                
                # Process reactions
                
                if(-not [string]::IsNullOrEmpty($reply.Reactions)) {
                    $Html += "<p>"
                    $Html += "<b>リアクション</b><br>"
                    $reply.Reactions | ConvertTo-Json
                    
                    foreach ($reaction in $reply.Reactions) {
                        # Change timezone
                        $JST = Change-Timezone($reaction.CreatedDateTime)

                        $Html += "$($JST) $($reaction.ReactionType) by $($reaction.User.DisplayName)<br>"
                    }

                    $Html += "</p>"
                }
                $Html += "</div>"
                $Html += "<br class=""clear_balloon""/>"
            }
            # End of Reply
        }
        # Added end tag to variable for output to file in HTML format
        $Html += "</div>"
        $Html += "</body></html>"
        
        # Output to file in HTML format (file name is team_channel_name.html)
        Out-File -FilePath "$($Team.DisplayName)_$($Channel.DisplayName).html" -InputObject $Html -Encoding UTF8
    }
}

# Disconnect from Microsoft Graph API
Disconnect-MgGraph

# Disconnect from Teams Admin Center
Disconnect-MicrosoftTeams
