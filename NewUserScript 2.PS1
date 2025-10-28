$DebugPreference = "Continue" #comment this line out when you don't want to enable the debugging output.
#$VerbosePreference = "Continue"
$ErrorActionPreference = "Continue"

$LogFolder = "$env:userprofile\desktop\logs" #log file location.
$TranscriptLog = -join($LogFolder,"\transcript.log")
Start-Transcript -Path $TranscriptLog -Force
$csvPath = 'C:\Users\sa.ps\OneDrive - domain.com\Documents\ExternalContacts' #changeme - Location where the website is delivering the CVS files.  Only a directory path is needed, do not enter a full path to a specific CSV file.



function Format-CsvValue {
    [CmdletBinding()]
    param (
        #Sets whether or not we want to format the provided string into 'title' (aka Proper) case when using named values.
        #When isTitleCase = $true the function will take the input string ($sValue) and format it to proper(title) case and will also remove leading and trailing whitespaces.  Example; "JoHN SmITH" will return "John Smith" or "   JaNE " will return "Jane" (removed whitespaces and set to title case).
        [Parameter(Mandatory = $false)]
        [bool]
        $isTitleCase = $false,
        #The string value that's passed into the function to properly format.
        #Example: Format-CsvValue -isTitleCase $true -sValue $mvar
        #Example: To only remove whitespace from a string-> Format-CsvValue -sValue $myvar
        [Parameter(Mandatory = $false)]  # was originally true
        [string]
        $sValue
    
    ) #=>Params
  
    begin {
        #no variables or intitializations to declare.
    } #=>begin
  
    process {
        if ($isTitleCase) {
            #isTitleCase is set to true so let's format it...
            $rValue = $((Get-Culture).TextInfo.ToTitleCase($sValue.ToLower())).Trim() #trim leading/trailing whitespace AND convert to title case string format.
        }
        else {
            #only whitespace trim is required, process that request.
            $rValue = $sValue.Trim() #Remove leading/trailing whitespace.
        }#=>if/isTitleCase
    }#=>process
  
    end {
        #return the value through the function.
        $rValue
    }
} #=>Format-CsvValue



Function Write-CustomEventLog {
    [CmdletBinding()]
    param(
        # What message to write to the event viewer.
        [Parameter(Mandatory=$true)]
        [string]
        $message,
        # Type
        [Parameter(Mandatory=$true)]
        [ValidateSet('Information','Warning','Error')]
        [string]
        $entryType
    )

    Begin {
        $eventSourceExists = [System.Diagnostics.EventLog]::SourceExists("Update-domainAD")
        if(-not($eventSourceExists)) {
            try {
                New-EventLog -LogName Application -Source 'Update-domainAD'
            }
            catch {
                Write-Debug 'Unable to create new application source.'
            }
        }#=>if not $eventSourceExists
    }#=>Begin

    Process {
        switch ($entryType) {
            "Information" { [int]$EventID = 1000 }
            "Warning" { [int]$EventID = 2000 }
            "Error" { [int]$EventID = 3000}
        }
        Write-EventLog -LogName Application -Source 'Update-domainAD' -EntryType $entryType -EventId $EventID -Message $message
    }
}

#Import-Module ActiveDirectory

if (!($isScheduled)) {
    Write-Debug "This is not a scheduled task so we can safely assume this is an initial read of a CSV file. Looking for all CSV files in $($csvPath) that are NOT readonly."
    #since we are anticipating *dynamically* named CSV files let's find all CSV files we have yet to process.
    $csvFiles = Get-ChildItem -Path $csvPath -Filter "NS*.csv" -Attributes !readonly+!directory
    $csvCount = ($csvFiles | Measure-Object).Count
    Write-Debug "Found $($csvCount) CSV files in $($csvPath) to process: `n`n `$csvFiles: $($csvFiles)"
    if ($csvFiles) {
        Write-Debug "Found unprocessed CSV files..."
        foreach ($csvFile in $csvFiles) {
            Write-Debug "Processing CSV file $($csvFile.FullName)"
            try {
                $Users = Import-CSV $csvFile.FullName
            }
            catch {
                
                
                #We need to check if the csvFiles count is greater than 1. If it is, we can move to the next file. If it's not, we need to throw an error and exit this script.
                if ($csvCount -gt '1') {
                    Write-CustomEventLog -message "Unable to import CSV file: $($csvFile.FullName). This is a fatal error for this csv file. Continuing to next file. Error message is: `n`n $($Error[0].Exception.Message)" -entryType "Warning"
                    Write-Debug "Unable to import our CSV file: $($csvFile.FullName). This is a fatal error for this CSV file.  Continuing to next file. Error message is: $Error[0].Exception.Message"
                    Continue
                } else {
                    Write-CustomEventLog -message "Unable to import CSV file: $($csvFile.FullName). This is a fatal error for this csv file and this script. Exiting script. Error message is: `n`n $($Error[0].Exception.Message)" -entryType "Error"
                    Write-Debug "Unable to import our CSV file: $($csvFile.FullName). This is a fatal error for this CSV file and this script. Exiting script. Error message is: $Error[0].Exception.Message"
                    Throw $csvFile.FullName
                }
                

            }#=> try $Users
        
            #imported our CSV file properly.  Let's process the file for new users...
            ForEach ($User in $Users){
                #debugging purposes...
                Write-Debug "First Name: $($User.Firstname)"
                Write-Debug "Last Name : $($User.Lastname)"
                Write-Debug "StartDate : $($User.startdate)"
                #Write-Debug "Email : $($User.Email)"
                Write-Debug "Title : $($User.Title)"
                Write-Debug "Department :$($User.Department)"
                Write-Debug "CopyUser : $($User.copyUser)"
                Write-Debug "Mobile     :$($user.Mobile)"
                Write-Debug "Manager    :$($user.Manager)"
                Write-Debug "BusinessUnit :$($user.BusinessUnit)"
                Write-Debug "OfficeLocation : $($User.OfficeLocation)"
                Write-Debug "Country :$($User.Country)"

        
                #Let's properly format all the values in this *ROW* of the CSV. Trim() where necessary and change to Title Case where necessary - also create a new variable so we can use it later when creating the user in AD using the New-ADuser cmdlet.
                $FirstName = Format-CsvValue -isTitleCase $true -sValue $User.FirstName #trim and title case
                $LastName = Format-CsvValue -isTitleCase $true -sValue $User.LastName #trim and title case.
                #$Email = Format-CsvValue -sValue $User.Email #trim only.
                $StartDate = Format-CsvValue -sValue $User.startdate #trim only.
                $copyUser = $User.copyUser
                $Title = $User.Title
                $PreferredLanguage = $User.PreferredLanguage
                $EmployeeType = $User.EmployeeType
                $Department = $User.Department
                $Mobile = $User.Mobile
                $Manager = $User.Manager
                $BusinessUnit = $User.BusinessUnit
                $officeLocation = $User.OfficeLocation.split('(')[0]
                $Country = $User.Country
                #$EndDate = Format-CsvValue -sValue $User.enddate #trim only.
                #$Company = Format-CsvValue -sValue $User.company #trim only since company names are rather specific on how they're spelled out.
                if ($csvFile.Name -like "NS*") {
                    #This csvFile that we're working on seems to be a New User request as defined by the "NU" in the CSV file name so we add more details.
                    $copyUser = $User.copyUser #We need the fullname of the user we're copying from.
                }
                #=> End of CSV values.

                #Let's build other necessary variables that we want to use as parameters for the New-ADuser cmdlet out of the information provided by the CSV file or other sources...
                $FullName = -join($($FirstName)," ",$($LastName)) #join $Firstname and $Lastname and a space to get FullName
                $SAM = -join($($FirstName.Replace(' ','')),'.',$($LastName.Replace(' ',''))) #this assumes that your SAM naming convention is first.LASTNAME and makes everything lowercase.
                if($SAM.Length -gt 20){$SAM = $SAM.Substring(0,20)}
                $DisplayName = -join($($FirstName),' ',$($LastName))
                $Username = (-join($FirstName,".",$LastName)).ToLower() #this assumes that your usernames have a naming convention of firstname.lastname and makes everything lowercase.
                $DNSroot = "@$((Get-ADDomain).dnsroot)"
                $UPN = -join($SAM,'@domain.com')
                $Email = -join($SAM,'@domain.com')
                $Title = $Title
                $Manager = $Manager
                $Company = "domain NV"
                $Password = (ConvertTo-SecureString -AsPlainText '#jiKlp@n2!' -Force)
                $oStartDate = $User.StartDate #This converts the CSV "startdate" field from a string to a datetime object so we can use it for comparison.
                #$oEndDate = [datetime]::ParseExact(($User.EndDate).Trim(), "dd/MM/yyyy", $null) #This conerts to CSV 'EndDate' field from a string to a datetime object which is required for the New-AdUser cmdlet 'AccountExpirationDate' parameter.
                $today=(get-date).ToString('yyyy-MM-dd')
                

                # This function can be called to send an mail in case of conflict in creating an account
                function Send-Mail{
                    [CmdletBinding()]
                    param ()
                    $cred = Get-Credential (Get-Secret -Vault KeyVaultStore -Name DecoAutomate)
                    $sendMbx = "Automate.domain@domain.com"
                    $recMbx = "smith.smith2@domain.com"
                    $body   = "User with SamAccountName '$SAM' already exist, this account will have to be created manually."
                
                    $mailProps = @{
                            
                        SMTPServer = "smtp.office365.com"
                        From = $sendMbx
                        To = $recMbx
                        Cc = "smith3.smith4@domain.com"
                        Subject = "SamAccountName Conflict($SAM)"
                        UseSSL = $true
                        Port = 587
                        Credential = $cred
                    }
                    Send-MailMessage @mailProps $body


                }
                function informsmith6 {
                    [CmdletBinding()]
                    param ()
                    $cred = Get-Credential (Get-Secret -Vault KeyVaultStore -Name DecoAutomate)
                    $sendMbx = "Automate.domain@domain.com"
                    $recMbx = "smith6.smith7@domain.com"
                    $body = "Hallo smith6,`nMailBox/agenda is klaar voor deze gebruiker om trainingen te kunnen plannen. : '$($SAM)' "

                    $mailProps = @{
                            
                        SMTPServer = "smtp.office365.com"
                        From = $sendMbx
                        To = $recMbx
                        Cc = "smith3.smith4@domain.com"
                        Subject = "Account is aangemaakt ($SAM)"
                        UseSSL = $true
                        Port = 587
                        Credential = $cred
                    }
                    Send-MailMessage @mailProps $body

                }
                #debugging purposes...
                Write-Debug "`$FirstName:  $($FirstName)"
                Write-Debug "`$LastName: $($LastName)"
                Write-Debug "`$Email: $($Email)"
                Write-Debug "`$StartDate: $($StartDate)"
                Write-Debug "`$EndDate: $($EndDate)"
                Write-Debug "`$copyUser: $($copyUser)"
                Write-Debug "`$FullName: $($FullName)"
                Write-Debug "`$SAM: $($SAM)"
                Write-Debug "`$Manager: $($Manager)"
                Write-Debug "`$Username: $($Username)"
                Write-Debug "`$DNSRoot: $($DNSroot)"
                Write-Debug "`$UPN: $($UPN)"
                Write-Debug "`$oStartDate: $($oStartDate)"

                #=>debugging puproses

                # Now, let's check the user's startdate as listed in the CSV file.  If startdate is within 48 hours of today's (Get-Date) date we'll create the user directly in AD.  Otherwise, we'll schedule a task to create the user at a later date.
                # First, we need to check if this is a New User request, 'startdate' only applies to new users...
                if ($csvFile.name -like "NS*") {
                    if ( $today -eq $oStartDate) {
                        Write-Debug "$(Get-Date) is due so we are creating the user immediately."

                        #Checking to see if a user already exists in AD with the same email address...
                        if (Get-AdUser -Filter {mail -eq $Email}) {
                            Send-Mail
                            Rename-Item -Path $csvFile.FullName -NewName "$($csvFile.FullName).done" -Force
                            Rename-Item -Path "C:\Users\sa.ps\Desktop\logs\transcript.log" -NewName "$($LastName).log" -Force
                            Set-ItemProperty -Path "$($csvFile.FullName).done" -name IsReadOnly -Value $true


                            Write-Debug "A user with email address $($email) already exists in AD.  We cannot add this user."
                            $failedUsers+= -join($Fullname,",",$SAM,",","A user with email address $($email) already exists in AD. Skipping this user.")
                            Write-CustomEventLog -message "When attempting to create user $($FullName) [SAM: $($SAM)] we found another user that exists in AD using the same email address of $($email). We have to skip this user." -entryType "Warning"
                            Continue #go to next csv record.
                        }#=if get-aduser
                        else {
                            Write-Debug "No existing user in AD with email address $($email) so we can create our user."

                            $newUserAD = @{
                                'SamAccountName'            = $SAM
                                'UserPrincipalName'         = $UPN
                                'Name'                      = $FullName
                                'DisplayName'               = $DisplayName
                                'Company'                   = $Company
                                'GivenName'                 = $FirstName
                                'Surname'                   = $LastName
                                'Title'                     = $Title
                                'Department'                = $Department
                                'Mobile'                    = $Mobile
                                'AccountPassword'           = $Password
                                'AccountExpirationDate'     = $oEndDate
                                'ChangePasswordAtLogon'     = $true
                                'Enabled'                   = $true
                                'PasswordNeverExpires'      = $false  
                                #'Country'                   = 'BE'
                                'EmailAddress'              = $UPN
                            }#=>$newUserAD

                            Write-Debug "Attempting to get properties of our user to copy from..."
                            #$templateUser = Get-ADUser -Identity $copyUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Add-ADGroupMember -Members $SAM
                            $templateUser = Get-ADUser -filter {SamAccountName -eq $copyUser} -Properties MemberOf
                            if (-not($templateUser)) {
                                Write-Debug "We were unable to find the template user $($copyUser) so we have to skip this new AD user and go to the next row in the CSV file."
                                #$failedUsers+= -join($Fullname,",",$SAM,",","We were unable to find the template user $($copyUser) so we have to skip creating new user $($FullName) and go to the next row in the CSV file.")
                                Write-CustomEventLog -message "We were unable to find the template user $($copyUser) when attempting to create new user $($FullName) with SAM $($SAM).  Skipping the creation of this user." -entryType "Warning"
                                continue #move to next CSV row.
                            } else {
                                #Let's get the OU that our template user belongs to and apply that to our new user...
                                $OU = ($templateUser.DistinguishedName).Substring(($templateUser.DistinguishedName).IndexOf(",")+1)
                                Write-Debug "Our OU for new user $($FullName) is $($OU) from copy of our template user $($copyUser) with OU of $($templateUser.DistinguishedName)"
                                #Let's update our $newUserAD properties with this OU...
                                $newUserAD['Path'] = $OU
                            }#=>if/else $templateuser

                            try {
                                Write-Debug "Adding user $($FullName) to AD with the following paramaters; `n $($newUserAD | Out-String)"
                                $oNewADUser = New-ADUser @newUserAD                                
                            }
                            catch {
                                Write-Debug "Unable to create new user $($FullName) to AD.  Error message `n`n $Error"
                                if(-not($oNewADUser)) {
                                    Write-Debug "Something went wrong with adding our new $($FullName) user to AD. `n`n $error"
                                    $failedUsers+= -join($Fullname,",",$SAM,",","We were unable to add our new user $($FullName) to AD. `n`n $error `n`n Moving to next user...")
                                    Write-CustomEventLog -message "We were unable to add our new user $($FullName) to AD. Skipping this user.  Full error details below; `n`n $($Error)." -entryType "Warning"
                                    continue
                                }
                            }
                            Start-Sleep -Seconds 10
                            #Adding user went well..
                            Get-ADUser -Identity $copyUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Add-ADGroupMember -Members $SAM
                            $templateManager = Get-ADUser -filter {UserPrincipalName -like $Manager} -Properties * | Select-Object SamAccountName
                            Get-ADUser -Identity $SAM | Set-ADUser -Add @{domainBusinessUnit = $BusinessUnit}
                            #Get-ADUser -Identity $SAM | Set-ADUser -Add @{
                                #co = 'Belgium'
                                #countryCode = 56
                            #}
                            Get-ADUser -Identity $SAM | Set-ADUser -Manager $templateManager -Fax '+32 56 528 800' 
                            Start-Sleep -Seconds 5
                            Get-ADUser -Identity $SAM | Set-ADUser -Add @{'PreferredLanguage' = $PreferredLanguage}
                            Start-Sleep -Seconds 5
                            Get-ADUser -Identity $SAM | Set-ADUser -Add @{'EmployeeType' = $EmployeeType}
                            Start-Sleep -Seconds 5
                            Get-ADUser -Identity $SAM | Set-ADUser -Add @{'extensionattribute2' = $EmployeeType}
                            Start-Sleep -Seconds 5
                            Get-ADUser -Identity $SAM | Set-ADUser -Add @{'extensionattribute3' = $BusinessUnit}
                            Start-Sleep -Seconds 5
                            Get-ADUser -Identity $SAM | Set-ADUser -Add @{'physicalDeliveryOfficeName' = $officeLocation}
                            Start-Sleep -Seconds 5
                            Get-ADUser -Identity $SAM | Set-ADUser -Add @{'co'= $Country}
                            
                            Add-ADGroupMember -Identity MS365 -Members $SAM
                            #Add-ADGroupMember -Identity MyITHubApp -Members $SAM
                            #Add-ADGroupMember -Identity HRPortalGrp -Members $SAM
                            Write-Debug "We created our new user $($FullName) in AD."

                            # Sleep for 45 minutes to allow the sync to complete
                            Write-Host "Waiting for 35 minutes to ensure AD sync completes..."

                            Start-Sleep -Seconds 2100  # 35 minutes

                            # Azure DevOps Integration (add user to DevOps)
                            $AzureDevOpsPAT = "thispersonalaccesstokencanbefoundindevopssettings"  # Replace with actual PAT
                            $OrganizationName = "domainGroup"
                            $ProjectId = "-4281-2c864654-2f38gtghyefrgthggbe49-77b5c7cce178"  # Replace with actual project ID
                            $AzureDevOpsAuthenicationHeader = @{
                                Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($AzureDevOpsPAT)"))
                            }

                            $UserBody = @{
                                accessLevel = @{
                                    accountLicenseType = "stakeholder"
                                };
                                extensions = @();
                                user = @{
                                    principalName = "$SAM@domain.com";
                                    subjectKind = "user"
                                };
                                projectEntitlements = @(
                                    @{
                                        group = @{
                                            groupType = "projectContributor"
                                        };
                                        projectRef = @{
                                            id = $ProjectId
                                        }
                                    }
                                )
                            } | ConvertTo-Json -Depth 10

                            $AddUserParameters = @{
                                Method = "POST"
                                Headers = $AzureDevOpsAuthenicationHeader
                                Uri = "https://vsaex.dev.azure.com/$OrganizationName/_apis/userentitlements?api-version=7.1"
                                Body = $UserBody
                                ContentType = "application/json"
                            }

                            try {
                                $AddUserResponse = Invoke-RestMethod @AddUserParameters
                                Write-CustomEventLog -message "User $SAM added to Azure DevOps project successfully." -entryType "Information"
                            } catch {
                                Write-CustomEventLog -message "Failed to add user $SAM to Azure DevOps." -entryType "Error"
                            }
                            $successUsers += -join($FullName,",",$SAM,",","Successfully created new AD user.")
                            Write-CustomEventLog -message "Successfully created new AD User $($FullName).  AD Details included below; `n`n $($newuserAD | Out-String)" -entryType "Information"
                            Write-Debug "Creating user mailbox..."

                            informsmith6
                            
			                # Connect to Mg-graph. The tokens and creds needed for authentication
                            $appid = 'hghghfhfhfhdhsgfygrfrf529525454'
                            $secret = 'lglgoflfjdfkfjfjfjfjfjfjfjfjf88f8f8f8f8f8f8'
                            $tenantid = 'jgjgjgjgjgkfkfkfkfjgjgjgkfkfjf969696'
                            
                            
                            $body =  @{
                                Grant_Type    = "client_credentials"
                                Scope         = "https://graph.microsoft.com/.default"
                                Client_Id     = $appid
                                Client_Secret = $secret
                            }
                            
                            $connection = Invoke-RestMethod `
                                -Uri https://login.microsoftonline.com/$tenantid/oauth2/v2.0/token `
                                -Method POST `
                                -Body $body
                            $token = $connection.access_token
                            
                            
                            Connect-MgGraph -AccessToken ($token |ConvertTo-SecureString -AsPlainText -Force) -NoWelcome
                            
                            Update-MgUser -UserId $UPN -AdditionalProperties @{'extension_akhkhkhkhkfkfhfkkdhdh0_UserEmployeeCategory' = $EmployeeType}
                            
                            
                            Disconnect-MgGraph	


                            # This function connects to ExchangeOnPrem to mirror the mailbox
                            function Connect-ExchangeOnPrem {
                                [CmdletBinding()]
                                Param(
                                    $Prefix = $null
                                )
                                $Sessions = Get-PSSession | Where-Object { $_.ComputerName -match "BEMENVMBX01" }
                                If (-Not ($Sessions)) {
                                    Function Get-MhExchangeServerInAdSite {
                                        $ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]
                                        $siteDN = $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName
                                        $configNC = ([ADSI]"LDAP://RootDse").configurationNamingContext
                                        $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC")
                                        $objectClass = "objectClass=msExchExchangeServer"
                                        $version = "versionNumber>=1937801568"
                                        $site = "msExchServerSite=$siteDN"
                                        $search.Filter = "(&($objectClass)($version)($site))"
                                        $search.PageSize = 1000
                                        [void] $search.PropertiesToLoad.Add("name")
                                        [void] $search.PropertiesToLoad.Add("networkaddress")
                                        $search.FindAll() | Foreach-Object {
                                            [PSCustomObject]@{
                                                Name = $_.Properties.name[0]
                                                FQDN = $_.Properties.networkaddress | ForEach-Object { 
                                                    if ($_ -match "ncacn_ip_tcp") {
                                                        $_.split(":")[1]
                                                    } 
                                                }
                                            }
                                        }
                                    }
                                    
                                    #add all servers in the local site to an array
                                    $ExchSrvrs = Get-MhExchangeServerInAdSite
                                    
                                    #select a random server from the current site
                                    If ($ExchSrvrs.count -gt 1) {
                                        $random = Get-Random -Minimum 0 -Maximum $ExchSrvrs.count
                                        $ExchSrvrFqdn = $ExchSrvrs[$random]
                                    } Else {
                                        $ExchSrvrFqdn = $ExchSrvrs[0]
                                    }
                            
                                    $URI = "http://$($ExchSrvrFqdn.Fqdn)/PowerShell/?SerializationLevel=Full"
                                    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URI -Authentication Kerberos -ErrorAction Stop
                                    If ($Prefix) {
                                        Import-PSSession $Session -DisableNameChecking -Prefix $Prefix | Out-Null    
                                    } Else {
                                        Import-PSSession $Session -DisableNameChecking -AllowClobber | Out-Null
                                    }
                                    Write-Verbose "Session connected to $($ExchSrvrFqdn.Fqdn)"
                                } Else {
                                    Write-Warning "Already connected to Exchange"
                                }
                            }
                            
                        
                            #Then you would be able to connect to an Exchange server via:
                            
                            #Connect-ExchangeOnPrem
                            #Enable-RemoteMailbox $sam -RemoteRoutingAddress "$($_.SAMAccountName)@domain.mail.onmicrosoft.com"
                            <#And if you wanted there to be no conflict with the EXO module:#>
                            
                            Connect-ExchangeOnPrem -Prefix 'OP' # As in "on-premises"
                            Enable-OPRemoteMailbox $SAM -RemoteRoutingAddress "$SAM@domain0.mail.onmicrosoft.com"
                             
                            
                            <#Edit: Forgot you would need to update EXCHSRV value with the naming standard (or list of names if disparate)
                            for your Exchange servers in your org if you care to not try re-connecting when you already have an established connection.
                            Line I'm referencing:
                            #>
                            
                            $Sessions = Get-PSSession | Where-Object { $_.ComputerName -match "EXCHSRV" }
                        }#=>else get-aduser


                    } else {
                        write-debug "Start date beyond 72hrs, leaving file unchanged!"
                        break
       
            }#=>ForEach $user !$isScheduled
            
            Write-Debug "Renaming our current csv file $($csvFile.FullName) and addding a .done extension. Also making the file read-only."
            Stop-Transcript
            # Rename the CSV file with the .done extension
            $NewFileName = Join-Path -Path $Csvfile.Directory.FullName -ChildPath ($Csvfile.BaseName + ".done")
            Rename-Item -Path $Csvfile.FullName -NewName $NewFileName -Force

            # Set the renamed CSV file as read-only
            $NewFileFullPath = Join-Path -Path $Csvfile.Directory.FullName -ChildPath ($Csvfile.BaseName + ".done")
            Set-ItemProperty -Path $NewFileFullPath -Name IsReadOnly -Value $true

 
                }
            }
        }
    }
}

Disconnect-ExchangeOnline -Confirm:$false