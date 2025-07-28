#------------------------------------------------------------------------------
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# Author Rui Pereira Tabares
# Modified to support MFA
#------------------------------------------------------------------------------

# PowerShell Functions
#------------------------------------------------------------------------------

write-host -ForegroundColor Red  "If you have never run this tool before, please verify if your admin has the following permissions:" 
Write-host -ForegroundColor Red "Go to https://protection.office.com/permissions"
Write-host -ForegroundColor Red "Add your admin account into the eDiscovery Manager and Compliance Administrator Permissions"
write-host -ForegroundColor Red "After adding your admin into those permissions, wait 30 to 40 minutes for them to take effect"

$Loop = $true
While ($Loop)
{
    write-host 
    write-host ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    write-host -BackgroundColor Magenta  "Copilot chat Export Tool"
    write-host ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    write-host
    write-host -ForegroundColor white  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Green   'Please select your option           ' 
    write-host -ForegroundColor white '----------------------------------------------------------------------------------------------' 
    write-host                                              ' 1)  Export Copilot data from a user mailbox hosted online'
    write-host                                              ' 2)  Export Copilot data from a user mailbox hosted onprem  '
    write-host
    write-host -ForegroundColor white  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Red 'End of PowerShell - Script menu ' 
    write-host -ForegroundColor  white  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor Yellow            "3)  Exit the PowerShell script menu" 
    write-host
    $opt = Read-Host "Select an option [1-2]"
    write-host $opt
    switch ($opt) 
    {
        1
        {  
            #### Getting PowerShell sessions ########
            $Sessions = Get-PSSession
            #region Connecting to Security & Compliance Center
            write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Security & Compliance Center if not already connected..." -foregroundColor Green
            if ( -not ($Sessions.ComputerName -match "ps.compliance.protection.outlook.com") ) {
                write-host
                if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) {
                    Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
                }
                Import-Module ExchangeOnlineManagement
                Connect-IPPSSession -ShowBanner:$false
            }
            #endregion

            #region Connecting to Exchange Online
            write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online if not already connected..." -foregroundColor Green
            if ( $Sessions.ComputerName -notcontains "outlook.office365.com" ) {
                write-host
                Connect-ExchangeOnline -ShowBanner:$false
            }
            #endregion

            ##### Online mailbox validation also works for O365 groups #####
            do { 
                $email = Read-Host "Enter an email address (Note: for groups, you can obtain this from https://admin.microsoft.com/Adminportal/Home?source=applauncher#/groups)"
                $emailAddress = $email

                $userMailbox = Get-Mailbox -Identity $email -ErrorAction SilentlyContinue
          
                if (($userMailbox -eq $null)) {
                    write-host "Error: Please enter a valid online mailbox" -ForegroundColor Red 
                }
            } while (($userMailbox -eq $null))

            #### Saving searched user display name
            $DisplayName = $userMailbox.DisplayName 

            write-host "Performing Search on $DisplayName Teams chat history"
            ## Starting search from user
            $searchName = ((Get-Date).ToString("HH:mm:ss")) + $DisplayName
            # Create the query using OR statements
            $copilotQuery = @"
ItemClass:IPM.SkypeTeams.Message.Copilot.Fabric. OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Studio. OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Excel OR
 ItemClass:IPM.SkypeTeams.Message.TeamCopilot.AiNotes.Teams OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Loop OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.M365App OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.BizChat OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Forms OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Security.SecurityCopilot OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.OneNote OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Outlook OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Powerpoint OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.SharePoint OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Teams OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.WebChat OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Whiteboard OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Word OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.
"@
            $complianceSearch = New-ComplianceSearch -Name $searchName -ContentMatchQuery $copilotQuery -ExchangeLocation $emailAddress
            Start-ComplianceSearch $searchName
            write-host "A Search from Copilot chat history has started with the Search Name" $searchName
            ## Loop until search completes
            do {
                Write-host "Waiting for search to complete..."
                Start-Sleep -s 5
                $complianceSearch = Get-ComplianceSearch $searchName
            } while ($complianceSearch.Status -ne 'Completed')

            if ($complianceSearch.Items -gt 0) {
                # Create a Compliance Search Action and wait for it to complete
                $complianceSearchAction = New-ComplianceSearchAction -SearchName $searchName -Preview
                do {
                    Write-host "Waiting for search action to complete..."
                    Start-Sleep -s 5
                    $complianceSearchAction = Get-ComplianceSearchAction -SearchName $searchName -Details
                } while ($complianceSearchAction.Status -ne 'Completed')
                
                $results = Get-ComplianceSearch -Identity $searchName | Select-Object SuccessResults
                $results = $results.SuccessResults -replace "@{SuccessResults={", "" -replace "}}",""
                if ($results -match "size:(\d+)") {
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
                    Write-Host 
                    Write-Host "Step 1: Go to Office 365 Security & Compliance->Searches"
                    Write-Host "Step 2: Check the search with name $searchName"
                    Write-Host "Step 3: Verify that you can view chat messages exported by clicking View Results"
                    Write-Host "Step 4: Click 'Export results', use the default options and click Export"
                    Write-Host "Step 5: Now go to Office 365 Security & Compliance->Exports and click Refresh"
                    Write-Host "Step 6: Click the export with name"
                    Write-Host "Step 7: Copy the Export Key and then click Download results and paste the Export key and specify the location where you want to download the exported chat messages."
                    Write-Host "Step 8: Click Start"
                    Write-Host "Step 9: When you see message download completed successfully then click on the link under Export location: to get to the exported files. Here you will find the exported PST of chat messages under Exchange folder <user or group email>.pst"
                    Write-Host "Step 10: Open Outlook application on your Windows PC and click File>Open&Export>Open Outlook data file (.pst) and provide location of the PST file from Step 9. In Outlook you will need to locate the folder <user or group email> and under that TeamsMessagesData. You will see all the chat messages exported here now."
                }
            }
        }

        2
        {   
            $Module = Get-Module
            if ( -not ($Module.Name -match "MSOnline") ) {
                write-host
                write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to MSOnline Service" -foregroundColor Green
                if ( !(Get-Module MSOnline -ListAvailable) -and !(Get-Module MSOnline) ) {
                    Install-Module MSOnline -Force -ErrorAction Stop
                }
                Import-Module MSOnline
                Connect-MsolService
            }

            #### Getting PowerShell sessions ########
            $Sessions = Get-PSSession
            #region Connecting to Security & Compliance Center
            write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Security & Compliance Center if not already connected..." -foregroundColor Green
            if ( -not ($Sessions.ComputerName -match "ps.compliance.protection.outlook.com") ) {
                write-host
                if ( !(Get-Module ExchangeOnlineManagement -ListAvailable) -and !(Get-Module ExchangeOnlineManagement) ) {
                    Install-Module ExchangeOnlineManagement -Force -ErrorAction Stop
                }
                Import-Module ExchangeOnlineManagement
                Connect-IPPSSession -ShowBanner:$false
            }
            #endregion

            #region Connecting to Exchange Online
            write-host "[$((Get-Date).ToString("HH:mm:ss"))] Connecting to Exchange Online if not already connected..." -foregroundColor Green
            if ( $Sessions.ComputerName -notcontains "outlook.office365.com" ) {
                write-host
                Connect-ExchangeOnline -ShowBanner:$false
            }
            #endregion

            do {
                $UserPrincipalName = Read-Host "Enter User Principal Name (UPN)"
                $User = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue
                if ($User -eq $null) {
                    write-host "User must be synced to the cloud or there is a typo in the UserPrincipalName" -ForegroundColor Red 
                }
            } while ($User -eq $null)

            $ValidateExoLicenseE = $User | Where-Object { $_.Licenses.ServiceStatus | Where-Object { $_.ServicePlan.ServiceName -eq "EXCHANGE_S_ENTERPRISE" -and $_.ProvisioningStatus -eq "Success" } }
            $ValidateExoLicenseS = $User | Where-Object { $_.Licenses.ServiceStatus | Where-Object { $_.ServicePlan.ServiceName -eq "EXCHANGE_S_STANDARD" -and $_.ProvisioningStatus -eq "Success" } }
                
            if ( -not (($ValidateExoLicenseE -or $ValidateExoLicenseS))) {
                write-host "Error: User does not have an Exchange Online license" -ForegroundColor Red 
                write-host "See requirements on https://docs.microsoft.com/en-us/microsoft-365/compliance/search-cloud-based-mailboxes-for-on-premises-users" -ForegroundColor Red 
                Exit
            }

            $OnpremValidation = $User | Select-Object -Property DisplayName, UserPrincipalName, isLicensed, @{label='MailboxLocation';expression={switch ($_.MSExchRecipientTypeDetails) {1 {'Onprem'; break} 2147483648 {'Office365'; break} default {'Unknown'}}}}
            $ValidationDup = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue

            if ( -not (($OnpremValidation.MailboxLocation -eq "Onprem") -and ($ValidationDup -eq $null))) {
                write-host "WARNING!!: MSExchRecipientTypeDetails is not in value 1, there is an Exchange Online mailbox for this user" -ForegroundColor Yellow 
                write-host "Please verify if there is not a duplicate mailbox in online for this user or you selected the wrong option" -ForegroundColor Yellow 
                Exit
            }

            $PrimarySmtp = (Get-Recipient -Identity $UserPrincipalName).PrimarySmtpAddress

            ## Starting search from user
            $DisplayName = (Get-Recipient -Identity $UserPrincipalName).Name
            $searchName = ((Get-Date).ToString("HH:mm:ss")) + $DisplayName
                
            # Create the query using OR statements
           $copilotQuery = @"
ItemClass:IPM.SkypeTeams.Message.Copilot.Fabric. OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Studio. OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Excel OR
 ItemClass:IPM.SkypeTeams.Message.TeamCopilot.AiNotes.Teams OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Loop OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.M365App OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.BizChat OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Forms OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Security.SecurityCopilot OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.OneNote OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Outlook OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Powerpoint OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.SharePoint OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Teams OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.WebChat OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Whiteboard OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.Word OR
 ItemClass:IPM.SkypeTeams.Message.Copilot.
"@
            $complianceSearch = New-ComplianceSearch -Name $searchName -ContentMatchQuery $copilotQuery -ExchangeLocation $PrimarySmtp -IncludeUserAppContent $true -AllowNotFoundExchangeLocationsEnabled $true
            Start-ComplianceSearch $searchName
            write-host "A Search from Copilot chat history has started with the Search Name" $searchName
            ## Loop until search completes
            do {
                Write-host "Waiting for search to complete..."
                Start-Sleep -s 5
                $complianceSearch = Get-ComplianceSearch $searchName
            } while ($complianceSearch.Status -ne 'Completed')

            if ($complianceSearch.Items -gt 0) {
                # Create a Compliance Search Action and wait for it to complete
                $complianceSearchAction = New-ComplianceSearchAction -SearchName $searchName -Preview
                do {
                    Write-host "Waiting for search action to complete..."
                    Start-Sleep -s 5
                    $complianceSearchAction = Get-ComplianceSearchAction -SearchName $searchName -Details
                } while ($complianceSearchAction.Status -ne 'Completed')
                
                $results = Get-ComplianceSearch -Identity $searchName | Select-Object SuccessResults
                $results = $results.SuccessResults -replace "@{SuccessResults={", "" -replace "}}",""
                if ($results -match "size:(\d+)") {
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
                    Write-Host "Step 1: Go to Office 365 Security & Compliance->Searches"
                    Write-Host "Step 2: Check the search with name $searchName"
                    Write-Host "Step 3: Verify that you can view chat messages exported by clicking View Results"
                    Write-Host "Step 4: Click 'Export results', use the default options and click Export"
                    Write-Host "Step 5: Now go to Office 365 Security & Compliance->Exports and click Refresh"
                    Write-Host "Step 6: Click the export with name"
                    Write-Host "Step 7: Copy the Export Key and then click Download results and paste the Export key and specify the location where you want to download the exported chat messages."
                    Write-Host "Step 8: Click Start"
                    Write-Host "Step 9: When you see message download completed successfully then click on the link under Export location: to get to the exported files. Here you will find the exported PST of chat messages under Exchange folder <user or group email>.pst"
                    Write-Host "Step 10: Open Outlook application on your Windows PC and click File>Open&Export>Open Outlook data file (.pst) and provide location of the PST file from Step 9. In Outlook you will need to locate the folder <user or group email> and under that TeamsMessagesData. You will see all the chat messages exported here now."
                }
            }
        }
        
        3
        {
            $Loop = $false
            Exit
        } 
    }
}
