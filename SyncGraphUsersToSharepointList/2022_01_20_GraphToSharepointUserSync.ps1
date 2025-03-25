### Graph API  ### ################## ################## ################## ##################

#Generate Access token for Oath/graph access
$timer = Get-Date -Format "MM/dd/yyyy HH:mm"
Write-Host "$timer - Generate Access token for Oath/graph access" -ForegroundColor Yellow
try 
{    
    $uri="https://login.microsoftonline.com/(tenantID)/oauth2/token"

    $body=@{
        scope='https://graph.microsoft.com'
        client_id="(clientID)"
        client_secret="(secret)"
        grant_type="client_credentials"
        resource='https://graph.microsoft.com'
    }

    $Tokenresponse=Invoke-RestMethod -Uri $uri -Method Post -Body $body 

    #Write-Host $Tokenresponse
    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "$timer - Tokenresponse Got" -ForegroundColor Green         
}
catch {
    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "`n$timer - Token Reqsponse : Something went wrong`n"  -ForegroundColor Red 
    Write-Host $_.Exception.Message -ForegroundColor Red 
    exit
}
# User Entity
# Reference: https://docs.microsoft.com/en-us/previous-versions/azure/ad/graph/api/entity-and-complex-type-reference#user-entity
# Get user data from the Graph API
$timer = Get-Date -Format "MM/dd/yyyy HH:mm"
Write-Host "$timer - Get user data from the Graph API" -ForegroundColor Yellow
try {    
    $select = "`$select=displayName,givenName,surname,jobTitle,businessPhones,mail,userPrincipalName,officeLocation,department,country,id,showInAddressList"
    $filter = "`$filter=accountEnabled eq true"
    $expand = "`$expand=manager(`$select=displayName;`$levels=1)"
    $top = "`$top=100"
    $apiUrl = "https://graph.microsoft.com/v1.0/users?$top&$select&$filter&$expand"
    
    $graphUsers = @()
    $pageCounter = 0
    do{
        $pageCounter += 1
        $results = Invoke-RestMethod `
            -Headers @{ Authorization = "Bearer $($Tokenresponse.access_token)" } `
            -Uri $apiUrl `
            -Method Get
        if ($results.value){
            $graphUsers += $results.value
        }
        else {
            $graphUsers += $results
        }
        $apiUrl = $results.'@odata.nextlink'  
        $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
        Write-Host "Page $pageCounter - $timer" -ForegroundColor Yellow
    }
    #while ($apiUrl)
    while ($apiUrl -and $pageCounter -lt 1)
    #while ($apiUrl -and $pageCounter -lt 200)

    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "$timer - Graph Data Got " -ForegroundColor Green     
}
catch {
    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "`n$timer - Data Response : Something went wrong`n"  -ForegroundColor Red 
    Write-Host $_.Exception.Message -ForegroundColor Red 
    exit
}

#data check
#Write-Host ($graphUsers | Format-Table | Out-String)
#foreach($user in $graphUsers){Write-Host $user.displayname $user.businessPhones[0]}
#exit


### Sharepoint ### ################## ################## ################## ##################

#Connect to the sharepoint list
$timer = Get-Date -Format "MM/dd/yyyy HH:mm"
Write-Host "$timer - Connect to the sharepoint list" -ForegroundColor Yellow

try {    
    
    Connect-PnPOnline `
        -Url "https://chemonics.sharepoint.com/sites/app/dir/" `
        -ClientId (clientID) `
        -Thumbprint (thumbprint) `
        -Tenant (tenant).onmicrosoft.com
        
    $web = Get-PnPWeb

    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "$timer - Connected to" $web.Title "Set" -ForegroundColor Green  
}
catch {
    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "`n$timer - Something went wrong`n"  -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red 
    exit
}



# create a hash table so we can check who's in the list already
$timer = Get-Date -Format "MM/dd/yyyy HH:mm"
Write-Host "$timer - create a hash table so we can check who's in the list already" -ForegroundColor Yellow
try{    
    
    $listItems = (Get-PnPListItem -List "GraphUserData" -PageSize 1000 -Fields "id","graphID")
    Write-host "Total number of Graph items found:"$listItems.count -ForegroundColor Yellow
    $hash = @{}
    $listItems | ForEach-Object {$hash.Add($_["graphID"],$_.id)}

    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "Total number of Hash items added : " $hash.Count -ForegroundColor Yellow
    Write-Host "$timer - Got hash of SP data" -ForegroundColor Green  
}
catch {
    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "$timer - Something went wrong"  -ForegroundColor Red 
    Write-Host $_.Exception.Message -ForegroundColor Red 
    exit
}



# check if Graph Data Users are in the Hash.
$timer = Get-Date -Format "MM/dd/yyyy HH:mm"
Write-Host "$timer - check if Graph Data Users are in the Hash" -ForegroundColor Yellow
try {
    
    $userCounter = 0
    foreach($user in $graphUsers)    
    {
        
        # user exists / Update        
        if($hash.ContainsKey($user.id))
        {
            #Write-Host "Update User " $user.id 
            #Write-Host $hash[$user.id]
            Set-PnPListItem `
            -List "GraphUserData" `
            -Identity $hash[$user.id] `
            -Values @{`
                "graphID"=$user.id;`
                "displayName"=$user.displayName;`
                "givenName"=$user.givenName;`
                "surname"=$user.surname;`
                "jobTitle"=$user.jobTitle;`
                "businessPhones"=$user.businessPhones[0];`
                "mail"=$user.mail;`
                "userPrincipalName"=$user.userPrincipalName;`
                "officeLocation"=$user.officeLocation;`
                "department"=$user.department;`
                "country"=$user.country;`
                "showInAddressList"=$user.showInAddressList;`
                "manager"=$user.manager.displayName;`
            } > $null            
            $hash.Remove($user.id)
        }
        else {        
            # user does not exist / Create
            #Write-Host "Create User " $user.id             
            Add-PnPListItem `
            -List "GraphUserData" `
            -Values @{`
                "graphID"=$user.id;`
                "displayName"=$user.displayName;`
                "givenName"=$user.givenName;`
                "surname"=$user.surname;`
                "jobTitle"=$user.jobTitle;`
                "businessPhones"=$user.businessPhones[0];`
                "mail"=$user.mail;`
                "userPrincipalName"=$user.userPrincipalName;`
                "officeLocation"=$user.officeLocation;`
                "department"=$user.department;`
                "country"=$user.country;`
                "showInAddressList"=$user.showInAddressList;`
                "manager"=$user.manager.displayName;`
            } > $null
            $hash.Remove($user.id)
        }
        
        $userCounter = $userCounter + 1
        if($userCounter%100 -eq 0)
        {
            $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
            Write-Host "$timer - Users Sync'd : $userCounter" -ForegroundColor Yellow
        }

    } 

    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "Hash items NOT in the Graph data : " $hash.Count -ForegroundColor Yellow
    Write-Host "$timer - Removing records not found in Hash" -ForegroundColor Yellow
    
 
    foreach ($hashItem in $hash.GetEnumerator())
    {
        # https://stackoverflow.com/questions/9015138/looping-through-a-hash-or-using-an-array-in-powershell
        #Write-Host ($hashItem | Format-Table | Out-String)
        #Write-Host "$($hashItem.Name): $($hashItem.Value)"
        Remove-PnPListItem -List "GraphUserData" -Identity $hashItem.Value -Recycle -Force
        Write-Host "Removing SP List ID : $($hashItem.Value)"
    }
    exit

    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "Hash table count : " $hash.Count
    Write-Host "$timer - Wrote Graph data to SP" -ForegroundColor Green  
}
catch {
    $timer = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "$timer - Something went wrong`n"  -ForegroundColor Red 
    Write-Host $_.Exception.Message -ForegroundColor Red 
    exit
}



$timer = Get-Date -Format "MM/dd/yyyy HH:mm"
Write-Host "$timer - End Script" -ForegroundColor Green



################## ################## ################## ################## ##################
### References ### ################## ################## ################## ##################
################## ################## ################## ################## ##################

# https://docs.microsoft.com/en-us/powershell/module/pki/import-pfxcertificate?view=windowsserver2019-ps
# https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/new-pnpazurecertificate?view=sharepoint-ps
# Import-PfxCertificate -Exportable -CertStoreLocation Cert:\LocalMachine\My -FilePath .\pnp.pfx
# Import-Certificate -FilePath "C:\Users\evanh\Downloads\newcert.cer" -CertStoreLocation 'Cert:\LocalMachine\My' -Verbose
# Import-PfxCertificate -Exportable -CertStoreLocation Cert:\LocalMachine\My -FilePath .\pnp.pfx #Install certificate
# Import-PfxCertificate -Exportable -CertStoreLocation Cert:\LocalMachine\My -FilePath "X:\My Drive\Code\PowerShell\2021_11_03_sharepoint_pnp.pfx"

# Delete by thumbprint
# Get-ChildItem Cert:\LocalMachine\My\(thumbprint) | Remove-Item

# Graph Paging
# https://danielchronlund.com/2018/11/19/fetch-data-from-microsoft-graph-with-powershell-paging-support/

# Documentation for updating list items
#https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/set-pnplistitem?view=sharepoint-ps

# .ps1 is not digitally signed.
# Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Uninstall-Module -Name SharePointPnPPowerShellOnline -AllVersions -Force                                                                                    
# Install-Module -Name PnP.PowerShell     
