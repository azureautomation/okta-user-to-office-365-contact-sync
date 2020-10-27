#Author: Jos Lieben (OGD)
#Script home: www.lieben.nu
#Copyright: MIT
#Manual: http://www.lieben.nu/liebensraum/2017/12/setting-up-okta-user-office-365-contact-synchronisation/
#Purpose: Sync users from Okta to Office 365 and import them as contacts if their email address does not already exist in Office 365
#Requires –Version 4

Param(
    [Parameter(Mandatory=$true)][String]$oktaURL, #URL to your okta tenant e.g. https://lieben.okta.com
    [Parameter(Mandatory=$true)][String]$o365credentialName, #name of the credential object in Azure with O365 credentials
    [Parameter(Mandatory=$true)][String]$oktaTokenName = "OKTA-TOKEN" #name of the credential object in Azure with the okta token information
)

$VerbosePreference = "SilentlyContinue"

function validateExOConnection{
    Param(
        [Parameter(Mandatory=$true)]$o365Creds,
        [switch]$retry
    )
    if($script:Session -eq $Null -or $script:Session.State -ne "Opened"){
        #There is no session, or it has gone stale
        write-verbose "Exchange Online connection not available or has gone stale, attempting to connect"
        $failed = $False
        try {
            $a = New-PSSessionOption
            $a.IdleTimeout = 432000000000
            $script:Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $o365Creds -Authentication Basic -AllowRedirection -SessionOption $a
            $res = Import-PSSession $Session -AllowClobber -DisableNameChecking -WarningAction SilentlyContinue
            write-verbose "Connected to Exchange Online"
            return $Null
        }catch{
            $failed = $True
            write-verbose "Failed to connect to Exchange Online! $($Error[0])"
        }
        if($failed -and !$retry){
            validateExoConnection -o365Creds $o365Creds -retry
            return $Null
        }
        if($failed -and $retry){
            Throw "Failed to connect to Exchange Online twice! Aborting $($Error[0])"
        }
    }       
}

function Invoke-PagedMethod($url) {
    try{
        $response = Invoke-WebRequest $url -Method GET -Headers $authHeader -ErrorAction Stop -UseBasicParsing
        $links = @{}
        if ($response.Headers.Link) { # Some searches (eg List Users with Search) do not support pagination.
            foreach ($header in $response.Headers.Link.split(",")) {
                if ($header -match '<(.*)>; rel="(.*)"') {
                    $links[$matches[2]] = $matches[1]
                }
            }
        }
        @{objects = ConvertFrom-Json $response.content; nextUrl = $links.next}
    }catch{
        Throw $_
    }
}

#build authentication header for API requests to Okta
function buildOktaAuthHeader{
    Param(
        [Parameter(Mandatory=$true)]$oktaToken
    )    
    $authHeader = @{
    'Content-Type'='application/json'
    'Accept'='application/json'
    'Authorization'= 'SSWS '+$oktaToken
    }
    return $authHeader
}

#retrieve all active users from Okta
function retrieveAllOktaUsers{
    Param(
        [Parameter(Mandatory=$true)]$oktaToken,
        [Parameter(Mandatory=$true)]$oktaURL
    )  
    $authHeader = buildOktaAuthHeader -oktaToken $oktaToken  
    $users = @()
    $res = @{"nextURL"="$oktaURL/api/v1/users?limit=200&filter=status eq `"ACTIVE`""}
    while($true){
        try{
            if($res.nextURL -eq $Null){
                write-verbose "finished retrieveAllOktaUsers function"
                break;
            }
            $res = Invoke-PagedMethod –Url $res.nextURL -Method GET
            if($res.objects){
                $users += $res.objects
            }
            write-verbose "retrieved $($users.Count) users..."
        }catch{
            Throw "failed to retrieve users from Okta: $_"
        }
    }    
    return $users  
}

function cacheO365Objects{
    Param(
        [Parameter(Mandatory=$true)]$o365Creds
    )    
    validateExOConnection -o365Creds $o365Creds
    $O365Cache = @{}
    $O365Cache.FastSearch = @{}
    $script:WarningActionPreference = "SilentlyContinue"
    $script:ErrorActionPreference = "SilentlyContinue"
    Write-Verbose "Retrieving O365 Mailboxes....."
    #first look for mailboxes
    try{
        [Array]$O365Cache.Mailboxes = @()
        get-mailbox -ResultSize Unlimited -erroraction SilentlyContinue -warningAction SilentlyContinue | Select-Object RecipientType,IsDirSynced,Identity,Guid,Alias,DisplayName,primarySmtpAddress,EmailAddresses | % {
            if($_){
                $O365Cache.Mailboxes += $_
                $O365Cache.FastSearch."$($_.primarySmtpAddress)" = "mbx_$($O365Cache.Mailboxes.Count-1)"
            }
        }
    }catch{$Null}
    Write-Verbose "Retrieving O365 Contacts....."
    #now look for contacts
    try{
        [Array]$O365Cache.MailContacts = @()
        get-mailcontact -ResultSize Unlimited -erroraction SilentlyContinue -warningAction SilentlyContinue | Select-Object RecipientType,IsDirSynced,Identity,Guid,Alias,DisplayName,primarySmtpAddress,EmailAddresses | % {
            if($_){
                $O365Cache.MailContacts += $_
                $O365Cache.FastSearch."$($_.primarySmtpAddress)" = "con_$($O365Cache.MailContacts.Count-1)"
            }
        }
    }catch{$Null}
    Write-Verbose "Retrieving O365 MailUsers....."
    #now look for mailusers
    try{
        [Array]$O365Cache.MailUsers = @()
        get-mailuser -ResultSize Unlimited -erroraction SilentlyContinue -warningAction SilentlyContinue | Select-Object RecipientType,IsDirSynced,Identity,Guid,Alias,DisplayName,primarySmtpAddress,EmailAddresses | % {
            if($_){
                $O365Cache.MailUsers += $_
                $O365Cache.FastSearch."$($_.primarySmtpAddress)" = "usr_$($O365Cache.MailUsers.Count-1)"
            }
        }
    }catch{$Null}
    Write-Verbose "Retrieving O365 Groups....."
    #now look for groups
    try{
        [Array]$O365Cache.Groups = @()
        Get-distributiongroup -ResultSize Unlimited -erroraction SilentlyContinue -warningAction SilentlyContinue | Select-Object RecipientType,IsDirSynced,Identity,Guid,Alias,DisplayName,primarySmtpAddress,EmailAddresses | % {
            if($_){
                $O365Cache.Groups += $_
                $O365Cache.FastSearch."$($_.primarySmtpAddress)" = "grp_$($O365Cache.Groups.Count-1)"
            }
        }
    }catch{$Null}
    Write-Verbose "All O365 objects cached"
    $script:WarningActionPreference = "Stop"
    $script:ErrorActionPreference = "Stop" 
    return $O365Cache
}

function searchExOForUserOrGroup{#returns Exchange Online object, when searching based on an email address, also finds if alias, but much slower
    Param(
        [Parameter(Mandatory=$true)]$O365ObjectCache,
        [Parameter(Mandatory=$true)]$searchQuery, #defaults to email address, but switches can modify this
        [Switch]$byAlias,
        [Switch]$byDisplayName
    )    
    if($searchQuery.Length -lt 2){Throw "Invalid query: $searchQuery"}
    if(!$byAlias -and !$byDisplayName) {
        if($searchQuery.IndexOf("smtp:",[System.StringComparison]::CurrentCultureIgnoreCase) -eq 0){
            $searchQuery = $searchQuery.SubString(5)
        }
        try{
            $res = $O365ObjectCache.FastSearch.$searchQuery
            switch($res.Split("_")[0]){
                "mbx"{
                    $retVal = $O365ObjectCache.Mailboxes[$res.Split("_")[1]]
                    if($retVal){return $retVal}
                }
                "con"{
                    $retVal = $O365ObjectCache.MailContacts[$res.Split("_")[1]]
                    if($retVal){return $retVal}
                }
                "usr"{
                    $retVal = $O365ObjectCache.MailUsers[$res.Split("_")[1]]
                    if($retVal){return $retVal}
                }
                "grp"{
                    $retVal = $O365ObjectCache.Groups[$res.Split("_")[1]]
                    if($retVal){return $retVal}
                }
                default{
                    Throw
                }
            }
        }catch{$Null}
        $searchQuery = "smtp:$searchQuery" #default to searching by email, prepend smtp: as this is how O365 knows them
    }
    #first look for a mailbox
    try{
        if($byAlias){[Array]$res = @($O365ObjectCache.Mailboxes | where {$_.alias -eq $searchQuery -and $_})}
        elseif($byDisplayName){[Array]$res = @($O365ObjectCache.Mailboxes | where {$_.DisplayName -eq $searchQuery -and $_})}
        else{[Array]$res = @($O365ObjectCache.Mailboxes | where {$_.EmailAddresses -Contains $searchQuery -and $_})}
        if($res.Count -eq 0){Throw "No Mailbox Found"}
    }catch{$res = $Null}
    #if none is found, look for a contact
    if($res -eq $Null){
        try{
            if($byAlias){[Array]$res = @($O365ObjectCache.MailContacts | where {$_.alias -eq $searchQuery -and $_})}
            elseif($byDisplayName){[Array]$res = @($O365ObjectCache.MailContacts | where {$_.DisplayName -eq $searchQuery -and $_})}
            else{[Array]$res = @($O365ObjectCache.MailContacts | where {$_.EmailAddresses -Contains $searchQuery -and $_})}
            if($res.Count -eq 0){Throw "No Contact Found"}
        }catch{$res = $Null}
    }
    #if none is found, look for a mailuser
    if($res -eq $Null){
        try{
            if($byAlias){[Array]$res = @($O365ObjectCache.MailUsers | where {$_.alias -eq $searchQuery -and $_})}
            elseif($byDisplayName){[Array]$res = @($O365ObjectCache.MailUsers | where {$_.DisplayName -eq $searchQuery -and $_})}
            else{[Array]$res = @($O365ObjectCache.MailUsers | where {$_.EmailAddresses -Contains $searchQuery -and $_})}
            if($res.Count -eq 0){Throw "No MailUser Found"}
        }catch{$res = $Null}
    }
    #if none is found, look for a Group
    if($res -eq $Null){
        try{
            if($byAlias){[Array]$res = @($O365ObjectCache.Groups | where {$_.alias -eq $searchQuery -and $_})}
            elseif($byDisplayName){[Array]$res = @($O365ObjectCache.Groups | where {$_.DisplayName -eq $searchQuery -and $_})}
            else{[Array]$res = @($O365ObjectCache.Groups | where {$_.EmailAddresses -Contains $searchQuery -and $_})}
            if($res.Count -eq 0){Throw "No DistributionGroup Found"}
        }catch{$res = $Null}
    }
    #return false if nothing found, otherwise return object
    $script:WarningActionPreference = "Stop"
    $script:ErrorActionPreference = "Stop" 
    if($res){
        return $res[0]
    }else{
        Throw "Could not find $searchQuery in Exchange Online"       
    }
}

function validateOktaUserProfile{
    Param(
        [Parameter(Mandatory=$true)]$oktaUserObject
    )  
    $validUser = $True
    if($oktaUserObject.profile.email -notmatch "\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,7}\b"){
        write-verbose "user $($oktaUserObject.profile.login) does not have a valid email address set: $($oktaUserObject.profile.email)"
        $validUser=$False
    }
    if($oktaUserObject.profile.lastName.length -le 1 -or $oktaUserObject.profile.firstName -le 1){
        write-verbose "user $($oktaUserObject.profile.login) does not have a valid first or lastname configured"
        $validUser=$False
    }
    if(!$validUser){
        return $False 
    }else{
        return $True
    }
}

function validateExistingContact{
    Param(
        [Parameter(Mandatory=$true)]$existingObject
    )  
    #if the object is not a contact, it is not in our sphere of influence
    if($existingObject.RecipientType -ne "MailContact"){
        write-verbose "$($existingObject.DisplayName) found in O365 but is not a contact, skipping"
        return $False
    }
    #Object is a contact object, check if it is being synchronized (read-only), if so, we cannot modify the contact
    if($existingObject.IsDirSynced -eq $True){
        write-verbose "$($existingObject.DisplayName) found in O365 but is read-only (synced through other source), skipping"
        return $False
    }
    return $True
}

function updateExistingContact{
    Param(
        [Parameter(Mandatory=$true)]$existingObject,
        [Parameter(Mandatory=$true)]$oktaUserObject
    )  
    #retrieve the matching O365 contact
    try{
        $o365Contact = Get-Contact -Identity $existingObject.guid.guid -ErrorAction Stop
    }catch{
        Throw "$($existingObject.identity) does not have a matching O365 (MSOL) contact and only exists in Exchange Online!"
    }

	#check if the first and lastname combined match the objects displayName
    $requiredDisplayName = "$($oktaUserObject.profile.firstName) $($oktaUserObject.profile.lastName)"
    if($existingObject.displayName -ne $requiredDisplayName){
        try{
            write-verbose "updating display name for $($existingObject.identity)"
            Set-MailContact -DisplayName $requiredDisplayName -Confirm:$False -Identity $existingObject.guid.guid -ErrorAction Stop
        }catch{Throw $_}
    }
    #check if the primary email address matches the one in Okta
    if($existingObject.PrimarySmtpAddress -ne $oktaUserObject.profile.email){
        try{
            write-verbose "updating email address for $($existingObject.identity)"
            Set-MailContact -WindowsEmailAddress $oktaUserObject.profile.email -Confirm:$False -Identity $existingObject.guid.guid -ErrorAction Stop
        }catch{Throw $_}
    }
    #in case the okta contact has an address, update this for the contact in O365 as well
    if($oktaUserObject.profile.streetAddress -and $o365Contact.StreetAddress -ne $oktaUserObject.profile.streetAddress){
        try{
            write-verbose "updating street address for $($existingObject.identity)"
            Set-Contact -Identity $existingObject.guid.guid -StreetAddress $oktaUserObject.profile.streetAddress -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
    #in case the okta contact has a zipcode, update this for the contact in O365 as well
    if($oktaUserObject.profile.zipCode -and $o365Contact.PostalCode -ne $oktaUserObject.profile.streetAddress){
        try{
            write-verbose "updating zipcode for $($existingObject.identity)"
            Set-Contact -Identity $existingObject.guid.guid -PostalCode $oktaUserObject.profile.zipCode -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
    #in case the okta contact has a city, update this for the contact in O365 as well
    if($oktaUserObject.profile.city -and $o365Contact.city -ne $oktaUserObject.profile.city){
        try{
            write-verbose "updating city for $($existingObject.identity)"
            Set-Contact -Identity $existingObject.guid.guid -City $oktaUserObject.profile.city -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
    #in case the okta contact has a department, update this for the contact in O365 as well
    if($oktaUserObject.profile.department -and $o365Contact.Department -ne $oktaUserObject.profile.department){
        try{
            write-verbose "updating department for $($existingObject.identity)"
            Set-Contact -Identity $existingObject.guid.guid -Department $oktaUserObject.profile.department -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
    #in case the okta contact has a title, update this for the contact in O365 as well
    if($oktaUserObject.profile.title -and $o365Contact.Title -ne $oktaUserObject.profile.title){
        try{
            write-verbose "updating Title for $($existingObject.identity)"
            Set-Contact -Identity $existingObject.guid.guid -Title $oktaUserObject.profile.title -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
    #in case the okta contact has a country set, update this for the contact in O365 as well
    if($oktaUserObject.profile.countryCode -and $o365Contact.CountryOrRegion -ne $oktaUserObject.profile.countryCode -and $oktaUserObject.profile.countryCode -ne 0){
        try{
            write-verbose "updating country for $($existingObject.identity)"
            Set-Contact -Identity $existingObject.guid.guid -CountryOrRegion $oktaUserObject.profile.CountryOrRegion -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
}

function createNewContact{
    Param(
        [Parameter(Mandatory=$true)]$oktaUserObject
    )
    try{
        $res = $Null
        Write-verbose "creating new contact object for $($oktaUserObject.profile.email)"
        $res = new-mailcontact -LastName $oktaUserObject.profile.lastName -FirstName $oktaUserObject.profile.firstName -DisplayName "$($oktaUserObject.profile.firstName) $($oktaUserObject.profile.lastName)" -ExternalEmailAddress $oktaUserObject.profile.email -Name "$($oktaUserObject.profile.firstName) $($oktaUserObject.profile.lastName) oktaContactSync"
    }catch{
        write-verbose "failed to create contact!"
        Throw $_
    } 	

    #in case the okta contact has an address, update this for the contact in O365 as well
    if($oktaUserObject.profile.streetAddress){
        try{
            write-verbose "updating street address"
            Set-Contact -Identity $res.guid.guid -StreetAddress $oktaUserObject.profile.streetAddress -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
    #in case the okta contact has an address, update this for the contact in O365 as well
    if($oktaUserObject.profile.zipCode){
        try{
            write-verbose "updating zipcode address"
            Set-Contact -Identity $res.guid.guid -PostalCode $oktaUserObject.profile.zipCode -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
    #in case the okta contact has a city, update this for the contact in O365 as well
    if($oktaUserObject.profile.city){
        try{
            write-verbose "updating city address"
            Set-Contact -Identity $res.guid.guid -City $oktaUserObject.profile.city -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
    #in case the okta contact has an department, update this for the contact in O365 as well
    if($oktaUserObject.profile.department){
        try{
            write-verbose "updating department"
            Set-Contact -Identity $res.guid.guid -Department $oktaUserObject.profile.department -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
    #in case the okta contact has an title, update this for the contact in O365 as well
    if($oktaUserObject.profile.title){
        try{
            write-verbose "updating title"
            Set-Contact -Identity $res.guid.guid -Title $oktaUserObject.profile.title -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
    #in case the okta contact has an country set, update this for the contact in O365 as well
    if($oktaUserObject.profile.countryCode -and $oktaUserObject.profile.countryCode -ne 0){
        try{
            write-verbose "updating country"
            Set-Contact -Identity $res.guid.guid -CountryOrRegion $oktaUserObject.profile.countryCode -ErrorAction Stop -Confirm:$False
        }catch{Throw $_}
    }
}

#lets get started
$WarningActionPreference = "Stop"
$ErrorActionPreference = "Stop"  

#try to set TLS to v1.2, Powershell defaults to v1.0
try{
    $res = [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    Write-Verbose "Set TLS protocol version to prefer v1.2"
}catch{
    Write-Error "Failed to set TLS protocol to prefer v1.2, script may fail" -ErrorAction Continue
    Write-Error $_ -ErrorAction Continue
}

#get OKTA token and O365 Creds
try{
    $oktaToken = Get-AutomationPSCredential -Name $oktaTokenName -ErrorAction Stop
    $o365credentials = Get-AutomationPSCredential -Name $o365credentialName
    $oktaToken = $oktaToken.GetNetworkCredential().Password
    Write-Output "Okta token and O365 credentials loaded"
}catch{
    Write-Error "Failed to load okta token and/or credentials"
    Exit
}

try{
    Write-Output "Open connection to Exchange Online..."
    validateExOConnection -o365Creds $o365credentials
}catch{
    write-output "Failed to connect to Exchange Online, aborting"
    Throw $_
}

try{
    Write-Output "Cache current Office 365 Objects..."
    $O365Cache = cacheO365Objects -o365Creds $o365credentials
}catch{
    write-output "Failed to cache objects in O365"
    Throw $_
}

Write-Output "Will retrieve all Okta users now"
$users = retrieveAllOktaUsers -oktaToken $oktaToken -oktaURL $oktaURL
#loop over all users returned returned by Okta
if($users.count -le 1){
    Throw "Less than 1 user returned by Okta, expected more..."
}

$failedUsers = 0
$failedUserObjects = @()
$updatedUsers = 0
$createdUsers = 0
$unchangedUsers = 0
$uniqueMails = @()
foreach($user in $users){
    $VerbosePreference = "SilentlyContinue"
    try{
        validateExOConnection -o365Creds $o365credentials
    }catch{
        write-output "Failed to connect to Exchange Online, aborting"
        Throw $_
    }
    $VerbosePreference = "Continue"
    $userValidity = validateOktaUserProfile -oktaUserObject $user
    if(!$userValidity){
        $failedUserObjects += $user
        $failedusers++
        continue
    }
    #add mail to array
    if($user.profile.email){
        $uniqueMails += $user.profile.email
    }

    #check if this object exists in Office 365
    $exists = $False
    try{
        $res = searchExOForUserOrGroup -O365ObjectCache $O365Cache -searchQuery $user.profile.email
        $exists = $True
    }catch{$res=$Null;$Null}
    
    if($exists){
        #check if the object is not read-only in Office 365 and that it is a MailContact
        if((validateExistingContact -existingObject $res) -eq $True){
            #update the existing object
            try{
                updateExistingContact -existingObject $res -oktaUserObject $user
                $updatedUsers++
            }catch{
                write-error "failed to update existing contact" -erroraction Continue
                write-error $_ -erroraction Continue
                $failedUsers++
            }
        }else{
            $unchangedUsers++
        }
    }else{
        #create new contact
        try{
            createNewContact -oktaUserObject $user
            $createdUsers++
        }catch{
            write-error "failed to create new contact" -erroraction Continue
            write-error $_ -erroraction Continue
            $failedUsers++
        }
    }
}

write-output "Users failed: $failedUsers"
write-output "Users created: $createdUsers"
write-output "Users changed: $updatedUsers"
write-output "Users unchanged: $unchangedUsers"

write-output "will now remove contacts which are no longer present in Okta, processing $($O365Cache.MailContacts.Count) contacts.."

if($uniqueMails.Count -gt 1){
    $O365Cache.MailContacts | Where{$_.Alias.EndsWith('oktaContactSync')} | % {
        if($uniqueMails -notcontains $_.PrimarySmtpAddress){
            #user doesn't exist in Okta, thus should be deleted.
            try{
                write-output "removing $($_.Alias) because it is no longer present in Okta"
                Remove-MailContact -Identity $_.Alias -ErrorAction Stop -Confirm:$False
            }catch{
                write-error "Failed to remove $($_.Alias)" -ErrorAction Continue
                write-error $_ -ErrorAction Continue
            }
        }
    }
}
