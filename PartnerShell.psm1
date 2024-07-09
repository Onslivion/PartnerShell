$MODULES_AVAILABLE = @(
    "Connect-ExchangeOnline",
    "Connect-IPPSSession",
    "Connect-MicrosoftTeams",
    "Connect-MgGraph"
)

function Get-TenantID { # Credit to Daniel Kåven | https://teams.se/powershell-script-find-a-microsoft-365-tenantid/
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0, HelpMessage="The domain name of the tenant")]
        [String]$domain
    )
    $request = Invoke-WebRequest -Uri https://login.windows.net/$domain/.well-known/openid-configuration -UseBasicParsing
    $data = ConvertFrom-Json $request.Content
    return $Data.token_endpoint.split('/')[3]
}

function Connect-ExchangeOnline-ViaPartner {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Client Tenant ID")]
        [String] $Tenant
    )
    Process {
        $token = New-PartnerAccessToken -ApplicationId "fb78d390-0c51-40cd-8e17-fdbfab77341b" -Scopes "https://outlook.office365.com/powershell-liveid" -Tenant $Tenant
        Connect-ExchangeOnline -AccessToken $token.AccessToken -DelegatedOrganization $Tenant
    }
}

function Connect-IPPSSession-ViaPartner {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Client Tenant ID")]
        [String] $Tenant
    )
    Process {
        $token = New-PartnerAccessToken -ApplicationId "fb78d390-0c51-40cd-8e17-fdbfab77341b" -Scopes "https://ps.compliance.protection.outlook.com/PowerShell-LiveId" -Tenant $Tenant
        Connect-ExchangeOnline -ConnectionUri "https://ps.compliance.protection.outlook.com/PowerShell-LiveId" -AccessToken $token.AccessToken -DelegatedOrganization $Tenant
    }
}

function Connect-MicrosoftTeams-ViaPartner {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Client Tenant ID")]
        [String] $Tenant
    )
    Process {
        $graphToken = New-PartnerAccessToken -ApplicationId "12128f48-ec9e-42f0-b203-ea49fb6af367" -Scopes "https://graph.microsoft.com/.default" -Tenant $Tenant
        $teamsToken = New-PartnerAccessToken -ApplicationId "12128f48-ec9e-42f0-b203-ea49fb6af367" -RefreshToken $graphToken.RefreshToken -Scopes "48ac35b8-9aa8-4d74-927d-1f4a14a0b239/.default" -Tenant $Tenant
        Connect-MicrosoftTeams -AccessTokens [$graphToken.AccessToken, $teamsToken.AccessToken]
    }
}

function Connect-MgGraph-ViaPartner { 
    param (
        [Parameter(Mandatory=$true, HelpMessage="Client Tenant ID")]
        [String] $Tenant
    )
    Process {
        Write-Host @"
        Due to the nature of the Micrsoft Graph plugin and the functionality of Graph itself,
        An application with a set of API permissions must be instantiated in the client tenant.
"@
        do {
            $scopes = Read-Host "Please enter a list of API permissions, separated by commas"
            $scopesList = $scopes.Split(",")
            Write-Output "List of scopes: $scopes"
            $Confirmation = Read-Host "Confirm? (y/N)"
        } while ($Confirmation.ToLower -ne "y")

        Connect-PartnerCenter -AccessToken $authResult.AccessToken

        $newGrant = New-Object -TypeName Microsoft.Store.PartnerCenter.Models.ApplicationConsents.ApplicationGrant
        $newGrant.EnterpriseApplicationId = "00000002-0000-0000-c000-000000000000"
        $newGrant.Scope = $scopesList

        New-PartnerCustomerApplicationConsent -ApplicationId "14d82eec-204b-4c2f-b7e8-296a70dab67e" -CustomerId $Tenant -ApplicationGrants @($newGrant) -DisplayName "Microsoft Graph Command Line Tools"

        $token = New-PartnerAccessToken -ApplicationId "14d82eec-204b-4c2f-b7e8-296a70dab67e" -Scopes "https://graph.microsoft.com/.default" -Tenant $Tenant

        Connect-MgGraph -AccessToken $token.AccessToken
    }
}

function Get-Choice {
    param (
        [Parameter(Mandatory=$true, HelpMessage="Input data")]
        [Array] $In,
        [Parameter(Mandatory=$false, HelpMessage="Relevant parameters")]
        [String[]] $Params,
        [Parameter(Mandatory=$false, HelpMessage="Allow selection of multiple values")]
        [switch] $MultipleChoice = $false,
        [Parameter(Mandatory=$false, HelpMessage="Page size")]
        [int] $pageSize = 8
    )
    Process {
        $currentPage = 0
        do {
            for ($i = $currentPage * $pageSize; $i -lt ($currentPage + 1) * $pageSize; $i++) {
                if ($i -ge $In.Length) {
                    break;
                }
                $outStr = $i.ToString() + ". "
                if ($PSBoundParameters.ContainsKey("Params")) {
                    for ($j = 0; $j -lt $Params.Length; $j++) {
                        $outStr += $In[$i].$($Params[$j]) + " | "
                    }
                }
                else {
                    $outStr += $In[$i]
                }
                Write-Host $outStr
            }

            $selectStr = "Pick the relevant option. "
            if ($MultipleChoice) {
                $selectStr += "Select several options by separating via commas. "
            }
            if ($pageSize -le $In.Length) {
                if (($currentPage + 1) * $pageSize -lt $In.Length) {
                    $selectStr += "(N) Next Page "
                }
                if ($currentPage -ne 0) {
                    $selectStr += "(P) Previous Page "
                }
            }

            $pick = Read-Host $selectStr

            switch ($pick) {
                {$_.ToLower() -eq "n"} { 
                    if ((($currentPage + 1) * $pageSize) -gt $In.Length) {
                        Write-Host "Cannot go further: this is the last page."
                        continue
                    }
                    else {
                        $currentPage++
                    }
                }
                {$_.ToLower() -eq "p"} { 
                    if ($currentPage -eq 0) {
                        Write-Host "Cannot go back any further: this is the first page."
                        continue
                    }
                    else {
                        $currentPage--
                    }
                }
                default { 
                    try {
                        if ($MultipleChoice) {
                            $pickMultiple = $pick.split(",")
                            $choices = foreach ($num in $pickMultiple) { ([int]::parse($num)) }
                        }
                        else {
                            $choices = @([int]::parse($pick))
                        }
                    }
                    catch {
                        Write-Host "Invalid input."
                        continue
                    }
            
                    $valid = $true
            
                    foreach ($option in $choices) {
                        if (!(($option -ge 0) -and ($option -lt $In.Length))) {
                            $valid = $false
                            Write-Host "Invalid option: $($option)"
                        }
                    }
            
                    if (!$valid) {
                        Write-Host "Invalid input."
                        continue
                    }
            
                    do {
                        foreach ($option in $choices) {
                            $confStr = ""
                            if ($PSBoundParameters.ContainsKey("Params")) {
                                for ($i = 0; $i -lt $Params.Length; $i++){
                                    $confStr += $In[$option].$($Params[$i]) + " | "
                                }
                            }
                            else {
                                $confStr += $In[$option]
                            }
                            Write-Host $confStr
                        }
                        $Confirmation = Read-Host "Confirm? (y/N)"
                        if (($null -eq $Confirmation) -or ($Confirmation.ToLower() -eq "n")) {
                            break
                        }
                    } while ($Confirmation.ToLower() -ne "y")
            
                    if ($Confirmation.ToLower() -eq "y") {
                        $choice = foreach ($option in $choices) { $In[$option] }
                    }
                }
            }

        } while ($null -eq $choice)

        return $choice
    }
}

function PartnerShell {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false, HelpMessage="Tenant ID / Domain'")]
        [String] $Tenant
    )
    Process {

        # Authenticate to Partner Center and store token information
        Connect-PartnerCenter

        # Acquire all customers associated to partner
        $customers = Get-PartnerCustomer
        
        # Determine client tenant to be authenticated into
        if ($PSBoundParameters.ContainsKey("Tenant")) {
            try {
                $tenantId = Get-TenantID $Tenant
                $inPartner = $false
                for ($i = 0; $i -lt $customers.Length; i++) {
                    if ($tenantId -eq $customers[$i].CustomerId) {
                        $inPartner = $true
                        break
                    }
                }
                if (!$inPartner) {
                    throw "Specified tenant " + $tenantId + " is not in available customers."
                }
            }
            catch {
                Write-Output "The specified tenant ID / domain is not acceptable."
                Write-Error -Message $_ -ErrorAction Stop
            }
        }
        else {
            $tenantId = (Get-Choice -In $customers -Params "Name","Domain").CustomerId
        }

        # List available modules to authenticate into and prompt for selected modules
        $actions = Get-Choice -In $MODULES_AVAILABLE -MultipleChoice

        foreach ($action in $actions) {
            $partnerAction = $action + "-ViaPartner"
            & $partnerAction -Tenant $tenantId
        }

    }
}