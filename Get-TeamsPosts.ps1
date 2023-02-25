#####
#
# By executing the command with a user name (UPN),
# the contents of the postings of the team to which
# the user belongs are retrieved and output to an HTML file.
# When executing the command, a pop-up will ask for authentication,
# so authenticate with the Teams global administrator.
# Please execute the command with PowerShell 7
# because PowerShell 5 will cause an error.
#
# Based on: https://qiita.com/seilian/items/225b1fe012d502bd4172
#
# Get-TeamsPosts.ps1
#   Seiichirou Hiraoka <seiichirou.hiraoka@gmail.com>
#     Initial Version: 2023/02/24
#
#   Usage: Get-TeamsPosts.ps1 UPN
#
#####

# パラメーターとしてUserName(UPN)を受け取る
# 必要に応じてDebug, Verboseをつける
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$UserName
)
# メールアドレスが正しい形式かチェックする関数
function Validate-UserName($UserName) {
    return $UserName -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
}

# メールアドレスが正しいかどうか判定する
$isValid = Validate-UserName $UserName

# デバッグ用に変数の中身を出力する
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
        $Html = "<html><head><title>$($Team.DisplayName) - $($Channel.DisplayName)</title></head><body>"

        # Output team and channel names and IDs
        $Html += "<p><b>Team Name: $($Team.DisplayName)</b></p>"
        $Html += "<p><b>Team ID: $($Team.GroupId)</b></p>"
        $Html += "<p><b>Channel Name: $($Channel.DisplayName)</b></p>"
        $Html += "<p><b>Channel ID: $($Channel.Id)</b></p>"
            
        # Retrieve channel postings
        $Messages = Get-MgTeamChannelMessage -TeamId $Team.GroupId -ChannelId $Channel.Id

        # Process each submission
        foreach ($Message in $Messages) {
            
            # Add the date and time of the post, the name of the submitter, and the text (HTML) to the variable
            $Html += "<hr>"
            $Html += "<p><b>件名:$($Message.Subject)</b></p>"
            $Html += "<p><b>$($Message.CreatedDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff"))</b> by <i>$($Message.From.User.DisplayName)</i></p>"
            $Html += "$($Message.Body.Content)"

            # Process attachments
            # $Message.attachments is not Null
            if(-not [string]::IsNullOrEmpty($Message.attachments)) {
                $Html += "<p><b>添付ファイル</b></p>"
                Write-Output $Message.attachments

                foreach ($attachment in $Message.attachments) {
                    $Html += "<p>$($attachment.ContentUrl)</p>"
                }
            }

            # Process reactions
            if(-not [string]::IsNullOrEmpty($Message.Reactions)) {
                $Html += "<p><b>リアクション</b></p>"
                Write-Output $Message.Reactions

                foreach ($reaction in $Message.Reactions) {
                    $Html += "<p>$($reaction.CreatedDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff")) $($reaction.ReactionType) by $($reaction.User.DisplayName)</p>"
                }
            }

            # Get replies
            $replies = Get-MgTeamChannelMessageReply -TeamId $Team.GroupId -ChannelId $Channel.Id -ChatMessageId $Message.Id -All -PageSize 50

            # Process replies
            # $replies is not Null)
        }
        if([string]::IsNullOrEmpty($replies)) {
            continue
        }
        
        Write-Output $replies        
        # Start Reply
        $Html += "<div><p><b>返信</b></p>"
        
        # Process each reply content
        foreach ($reply in $replies) {
            if ($reply.messageType -ne "message") {
                continue
            }
            
            $replyData = New-Object PSObject | Select-Object DateTime, From, Type, Content, Attachments
            $replyData.DateTime = $reply.createdDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
            $replyData.From = $reply.from.user.displayName
            $replyData.Type = "Reply"
            $replyData.Content = $reply.body.content
            
            # Add the date and time of the post, the name of the submitter, and the text (HTML) to the variable
            $Html += "<p><b>$($replyData.DateTime)</b> by <i>$($replyData.From)</i></p>"
            $Html += "$($replyData.Content)"
            
            # Process attachments
            
            if(-not [string]::IsNullOrEmpty($reply.attachments)) {
                $Html += "<p><b>添付ファイル</b></p>"
                Write-Output $reply.attachments
                
                foreach ($attachment in $reply.attachments) {
                    $Html += "<p>$($attachment.ContentUrl)</p>"
                }
            }
            
            # Process reactions
            
            if(-not [string]::IsNullOrEmpty($reply.Reactions)) {
                $Html += "<p><b>リアクション</b></p>"
                Write-Output $reply.Reactions
                
                foreach ($reaction in $reply.Reactions) {
                    $Html += "<p>$($reaction.CreatedDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff")) $($reaction.ReactionType) by $($reaction.User.DisplayName)</p>"
                }
            }
        }
        # End of Reply
        $Html += "</div>"
    }
    # Added end tag to variable for output to file in HTML format
    $Html += "</body></html>"
    
    # Output to file in HTML format (file name is team_channel_name.html)
    Out-File -FilePath "$($Team.DisplayName)_$($Channel.DisplayName).html" -InputObject $Html -Encoding UTF8
}

# Disconnect from Microsoft Graph API
Disconnect-MgGraph

# Disconnect from Teams Admin Center
Disconnect-MicrosoftTeams
