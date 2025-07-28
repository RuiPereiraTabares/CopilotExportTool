#------------------------------------------------------------------------------
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# Author Rui Pereira Tabares
# Modified for MFA support
#------------------------------------------------------------------------------

# Ensure ExchangeOnlineManagement module is installed
if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
    Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
}
Import-Module ExchangeOnlineManagement

Write-Host -ForegroundColor Red "If you have never run this tool before, verify your admin has the following permissions:"
Write-Host -ForegroundColor Red "Go to https://protection.office.com/permissions"
Write-Host -ForegroundColor Red "Add your admin account to eDiscovery Manager and Compliance Administrator roles"
Write-Host -ForegroundColor Red "Wait 30-40 minutes for permissions to take effect"

$Loop = $true
While ($Loop) {
    Write-Host
    Write-Host "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    Write-Host -BackgroundColor Magenta "Copilot Chat Export Tool"
    Write-Host "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    Write-Host
    Write-Host -ForegroundColor White '----------------------------------------------------------------------------------------------'
    Write-Host -ForegroundColor White -BackgroundColor Green 'Please select your option'
    Write-Host -ForegroundColor White '----------------------------------------------------------------------------------------------'
    Write-Host ' 1) Export Copilot data from a user mailbox hosted online'
    Write-Host ' 2) Export Copilot data from a user mailbox hosted on-premises'
    Write-Host
    Write-Host -ForegroundColor White '----------------------------------------------------------------------------------------------'
    Write-Host -ForegroundColor White -BackgroundColor Red 'End of PowerShell - Script menu'
    Write-Host -ForegroundColor White '----------------------------------------------------------------------------------------------'
    Write-Host -ForegroundColor Yellow "3) Exit the PowerShell script menu"
    Write-Host
    $opt = Read-Host "Select an option [1-3]"
    Write-Host $opt

    switch ($opt) {
        1 {
            # Connect to Security & Compliance Center with MFA
            Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Connecting to Security & Compliance Center..." -ForegroundColor Green
            $Sessions = Get-PSSession
            if (-not ($Sessions.ComputerName -match "ps.compliance.protection.outlook.com")) {
                Connect-IPPSSession -ShowBanner:$False
            }

            # Connect to Exchange Online with MFA
            Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Connecting to Exchange Online..." -ForegroundColor Green
            if (-not ($Sessions.ComputerName -contains "outlook.office365.com")) {
                Connect-ExchangeOnline -ShowBanner:$False
            }

            # Validate online mailbox
            do {
                $email = Read-Host "Enter an email address (Note: for groups, obtain from https://admin.microsoft.com/Adminportal/Home?source=applauncher#/groups)"
                $emailAddress = $email
                $userMailbox = Get-Mailbox -Identity $email -ErrorAction SilentlyContinue
                if ($null -eq $userMailbox) {
                    Write-Host "Error: Please enter a valid online mailbox" -ForegroundColor Red
                }
            } while ($null -eq $userMailbox)

            $DisplayName = $userMailbox.DisplayName
            Write-Host "Performing Search on $DisplayName Copilot chat history"

            # Create and start compliance search
            $searchName = ((Get-Date).ToString("HH:mm:ss")) + $DisplayName
            $copilotQuery = @'
ItemClass:"IPM.SkypeTeams.Message.Copilot.Fabric." OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Studio." OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Excel" OR
ItemClass:"IPM.SkypeTeams.Message" OR
ItemClass:"IPM.SkypeTeams.Message.TeamCopilot.AiNotes.Teams" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Loop" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.M365App" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot." OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.BizChat" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Forms" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Security.SecurityCopilot" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.OneNote" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Outlook" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Powerpoint" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.SharePoint" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Teams" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.WebChat" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Whiteboard" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Word"
'@ -replace "`r`n", " "

            $complianceSearch = New-ComplianceSearch -Name $searchName -ContentMatchQuery $copilotQuery -ExchangeLocation $emailAddress
            Start-ComplianceSearch $searchName
            Write-Host "A Search from Copilot chat history has started with the Search Name: $searchName"

            # Wait for search to complete
            do {
                Write-Host "Waiting for search to complete..."
                Start-Sleep -Seconds 5
                $complianceSearch = Get-ComplianceSearch $searchName
            } while ($complianceSearch.Status -ne 'Completed')

            if ($complianceSearch.Items -gt 0) {
                $complianceSearchAction = New-ComplianceSearchAction -SearchName $searchName -Preview
                do {
                    Write-Host "Waiting for search action to complete..."
                    Start-Sleep -Seconds 5
                    $complianceSearchAction = Get-ComplianceSearchAction -Identity "$searchName`_Preview"
                } while ($complianceSearchAction.Status -ne 'Completed')

                $results = Get-ComplianceSearch -Identity $searchName | Select-Object -ExpandProperty SuccessResults
                $results = $results -replace "@{SuccessResults={", "" -replace "}}", ""
                $results -match "size:(\d+)"
                $match = $matches[1]
                $matchMb = $match / 1MB
                $matchGb = $match / 1GB
                Write-Host "------------------------"
                Write-Host "Results"
                Write-Host "------------------------"
                Write-Host "$results"
                Write-Host "------------------------"
                Write-Host "Found Size"
                Write-Host "$matchMb Mb"
                Write-Host "$matchGb Gb"
                Write-Host "________________________"
                Write-Host -ForegroundColor Green "Success"
                Write-Host "________________________"
                Write-Host "Step 1: Go to Office 365 Security & Compliance -> Searches"
                Write-Host "Step 2: Check the search with name: $searchName"
                Write-Host "Step 3: Verify you can view chat messages by clicking View Results"
                Write-Host "Step 4: Click 'Export results', use default options, and click Export"
                Write-Host "Step 5: Go to Office 365 Security & Compliance -> Exports and click Refresh"
                Write-Host "Step 6: Click the export with name: $searchName"
                Write-Host "Step 7: Copy the Export Key, click Download results, paste the Export Key, and specify the download location"
                Write-Host "Step 8: Click Start"
                Write-Host "Step 9: When download completes, click the link under Export location to access the exported PST file"
                Write-Host "Step 10: Open Outlook, go to File -> Open & Export -> Open Outlook Data File (.pst), select the PST file, locate the folder <$emailAddress> -> TeamsMessagesData to view exported chat messages"
            }
        }

        2 {
            # Connect to MSOnline for on-premises user validation
            if (-not (Get-Module MSOnline -ListAvailable)) {
                Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Installing MSOnline module..." -ForegroundColor Yellow
                Install-Module MSOnline -Force -ErrorAction Stop
            }
            Import-Module MSOnline
            Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Connecting to MSOnline..." -ForegroundColor Green
            Connect-MsolService

            # Connect to Security & Compliance Center with MFA
            Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Connecting to Security & Compliance Center..." -ForegroundColor Green
            $Sessions = Get-PSSession
            if (-not ($Sessions.ComputerName -match "ps.compliance.protection.outlook.com")) {
                Connect-IPPSSession -ShowBanner:$False
            }

            # Connect to Exchange Online with MFA
            Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] Connecting to Exchange Online..." -ForegroundColor Green
            if (-not ($Sessions.ComputerName -contains "outlook.office365.com")) {
                Connect-ExchangeOnline -ShowBanner:$False
            }

            # Validate on-premises user
            do {
                $UserPrincipalName = Read-Host "Enter User Principal Name (UPN)"
                $User = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue
                if ($null -eq $User) {
                    Write-Host "Error: User must be synced to the cloud or there is a typo in the UserPrincipalName" -ForegroundColor Red
                }
            } while ($null -eq $User)

            $ValidateExoLicenseE = $User | Where-Object { $_.Licenses.ServiceStatus | Where-Object { $_.ServicePlan.ServiceName -eq "EXCHANGE_S_ENTERPRISE" -and $_.ProvisioningStatus -eq "Success" } }
            $ValidateExoLicenseS = $User | Where-Object { $_.Licenses.ServiceStatus | Where-Object { $_.ServicePlan.ServiceName -eq "EXCHANGE_S_STANDARD" -and $_.ProvisioningStatus -eq "Success" } }

            if (-not ($ValidateExoLicenseE.IsLicensed -or $ValidateExoLicenseS.IsLicensed)) {
                Write-Host "Error: User does not have an Exchange Online license" -ForegroundColor Red
                Write-Host "See requirements at https://docs.microsoft.com/en-us/microsoft-365/compliance/search-cloud-based-mailboxes-for-on-premises-users" -ForegroundColor Red
                Exit
            }

            $OnpremValidation = $User | Select-Object -Property DisplayName, UserPrincipalName, IsLicensed, @{label='MailboxLocation';expression={switch ($_.MSExchRecipientTypeDetails) {1 {'Onprem'; break} 2147483648 {'Office365'; break} default {'Unknown'}}}}
            $ValidationDup = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue

            if (-not ($OnpremValidation.MailboxLocation -eq "Onprem" -and $null -eq $ValidationDup)) {
                Write-Host "WARNING: MSExchRecipientTypeDetails is not 1, an Exchange Online mailbox exists for this user" -ForegroundColor Yellow
                Write-Host "Verify there is no duplicate mailbox in Exchange Online or you selected the wrong option" -ForegroundColor Yellow
                Exit
            }

            $PrimarySmtp = (Get-Recipient -Identity $UserPrincipalName).PrimarySmtpAddress
            $DisplayName = (Get-Recipient -Identity $UserPrincipalName).Name
            $searchName = ((Get-Date).ToString("HH:mm:ss")) + $DisplayName

            # Create and start compliance search
            $copilotQuery = @'
ItemClass:"IPM.SkypeTeams.Message.Copilot.Fabric." OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Studio." OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Excel" OR
ItemClass:"IPM.SkypeTeams.Message" OR
ItemClass:"IPM.SkypeTeams.Message.TeamCopilot.AiNotes.Teams" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Loop" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.M365App" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot." OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.BizChat" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Forms" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Security.SecurityCopilot" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.OneNote" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Outlook" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Powerpoint" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.SharePoint" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Teams" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.WebChat" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Whiteboard" OR
ItemClass:"IPM.SkypeTeams.Message.Copilot.Word"
'@ -replace "`r`n", " "

            $complianceSearch = New-ComplianceSearch -Name $searchName -ContentMatchQuery $copilotQuery -ExchangeLocation $PrimarySmtp -IncludeUserAppContent $true -AllowNotFoundExchangeLocationsEnabled $true
            Start-ComplianceSearch $searchName
            Write-Host "A Search from Copilot chat history has started with the Search Name: $searchName"

            # Wait for search to complete
            do {
                Write-Host "Waiting for search to complete..."
                Start-Sleep -Seconds 5
                $complianceSearch = Get-ComplianceSearch $searchName
            } while ($complianceSearch.Status -ne 'Completed')

            if ($complianceSearch.Items -gt 0) {
                $complianceSearchAction = New-ComplianceSearchAction -SearchName $searchName -Preview
                do {
                    Write-Host "Waiting for search action to complete..."
                    Start-Sleep -Seconds 5
                    $complianceSearchAction = Get-ComplianceSearchAction -Identity "$searchName`_Preview"
                } while ($complianceSearchAction.Status -ne 'Completed')

                $results = Get-ComplianceSearch -Identity $searchName | Select-Object -ExpandProperty SuccessResults
                $results = $results -replace "@{SuccessResults={", "" -replace "}}", ""
                $results -match "size:(\d+)"
                $match = $matches[1]
                $matchMb = $match / 1MB
                $matchGb = $match / 1GB
                Write-Host "------------------------"
                Write-Host "Results"
                Write-Host "------------------------"
                Write-Host "$results"
                Write-Host "------------------------"
                Write-Host "Found Size"
                Write-Host "$matchMb Mb"
                Write-Host "$matchGb Gb"
                Write-Host "________________________"
                Write-Host -ForegroundColor Green "Success"
                Write-Host "________________________"
                Write-Host "Step 1: Go to Office 365 Security & Compliance -> Searches"
                Write-Host "Step 2: Check the search with name: $searchName"
                Write-Host "Step 3: Verify you can view chat messages by clicking View Results"
                Write-Host "Step 4: Click 'Export results', use default options, and click Export"
                Write-Host "Step 5: Go to Office 365 Security & Compliance -> Exports and click Refresh"
                Write-Host "Step 6: Click the export with name: $searchName"
                Write-Host "Step 7: Copy the Export Key, click Download results, paste the Export Key, and specify the download location"
                Write-Host "Step 8: Click Start"
                Write-Host "Step 9: When download completes, click the link under Export location to access the exported PST file"
                Write-Host "Step 10: Open Outlook, go to File -> Open & Export -> Open Outlook Data File (.pst), select the PST file, locate the folder <$PrimarySmtp> -> TeamsMessagesData to view exported chat messages"
            }
        }

        3 {
            $Loop = $false
            Exit
        }
    }
}
