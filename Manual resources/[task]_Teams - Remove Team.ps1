# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$baseGraphUri = "https://graph.microsoft.com/"

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# variables configured in form
$groupid = $form.teams.GroupId
$team = $form.teams.DisplayName
$description = $form.teams.Description

# Create authorization token and add to headers
try{
    Write-Information "Generating Microsoft Graph API Access Token"

    $baseUri = "https://login.microsoftonline.com/"
    $authUri = $baseUri + "$AADTenantID/oauth2/token"

    $body = @{
        grant_type    = "client_credentials"
        client_id     = "$AADAppId"
        client_secret = "$AADAppSecret"
        resource      = "https://graph.microsoft.com"
    }

    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token;

    #Add the authorization header to the request
    $authorization = @{
        Authorization  = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept         = "application/json";
    }
}
catch{
    throw "Could not generate Microsoft Graph API Access Token. Error: $($_.Exception.Message)"    
}

try {
    Write-Information "Deleting Team [$team] with description [$description]."

    $deleteTeamUri = $baseGraphUri + "v1.0/groups/$groupid"

    $deleteTeam = Invoke-RestMethod -Method DELETE -Uri $deleteTeamUri -Headers $authorization -Verbose:$false
    
    Write-Information "Successfully deleted team [$team] with description [$description]."
    $Log = @{
        Action            = "DeleteResource" # optional. ENUM (undefined = default) 
        System            = "MicrosoftTeams" # optional (free format text) 
        Message           = "Successfully deleted team [$team] with description [$description]." # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $team # optional (free format text)
        TargetIdentifier  = $groupid # optional (free format text)
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
catch
{
    Write-Error "Failed to delete team [$team] with description [$description]. Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "DeleteResource" # optional. ENUM (undefined = default) 
        System            = "MicrosoftTeams" # optional (free format text) 
        Message           = "Failed to delete team [$team] with description [$description]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $team # optional (free format text)
        TargetIdentifier  = $groupid # optional (free format text)
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}

