Param
(

[cmdletbinding()]
    [Parameter(Mandatory= $true, HelpMessage="Enter your ApplicationId from the Secure Application Model https://github.com/KelvinTegelaar/SecureAppModel/blob/master/Create-SecureAppModel.ps1")]
    [string]$ApplicationId,
    [Parameter(Mandatory= $true, HelpMessage="Enter your ApplicationSecret from the Secure Application Model")]
    [string]$ApplicationSecret,
    [Parameter(Mandatory= $true, HelpMessage="Enter your Partner Tenantid")]
    [string]$tenantID,
    [Parameter(Mandatory= $true, HelpMessage="Enter your refreshToken from the Secure Application Model")]
    [string]$refreshToken,
    [Parameter(Mandatory= $true, HelpMessage="Enter your Exchange refreshToken from the Secure Application Model")]
    [string]$ExchangeRefreshToken,
    [Parameter(Mandatory= $true, HelpMessage="Enter the UPN of a global admin in partner center")]
    [string]$upn

)

# Check if the MSOnline PowerShell module has already been loaded.
if ( ! ( Get-Module MSOnline) ) {
    # Check if the MSOnline PowerShell module is installed.
    if ( Get-Module -ListAvailable -Name MSOnline ) {
        Write-Host -ForegroundColor Green "Loading the Azure AD PowerShell module..."
        Import-Module MsOnline
    } else {
        Install-Module MsOnline
    }
}

###MICROSOFT SECRETS#####

$ApplicationId = $ApplicationId
$ApplicationSecret = $ApplicationSecret
$tenantID = $tenantID
$refreshToken = $refreshToken
$ExchangeRefreshToken = $ExchangeRefreshToken
$upn = $upn
$secPas = $ApplicationSecret| ConvertTo-SecureString -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $secPas)
 
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $tenantID
$graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $tenantID

Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken
 
$customers = Get-MsolPartnerContract -All
 
Write-Host "Found $($customers.Count) customers for $((Get-MsolCompanyInformation).displayname)." -ForegroundColor DarkGreen

#Define CSV Path 
$path = echo ([Environment]::GetFolderPath("Desktop")+"\ATPSettings")
New-Item -ItemType Directory -Force -Path $path
$ATPSettings = echo ([Environment]::GetFolderPath("Desktop")+"\ATPSettings\ATPCustomerList.csv")
 
foreach ($customer in $customers) {
    #Dispaly customer name#
    Write-Host "Checking ATP settings for $($Customer.Name)" -ForegroundColor Green
    #Establish Token for Exchange Online
    $token = New-PartnerAccessToken -ApplicationId 'a0c73c16-a7e3-4564-9a95-2bdf47383716'-RefreshToken $ExchangeRefreshToken -Scopes 'https://outlook.office365.com/.default' -Tenant $customer.TenantId
    $tokenValue = ConvertTo-SecureString "Bearer $($token.AccessToken)" -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($upn, $tokenValue)
    $InitialDomain = Get-MsolDomain -TenantId $customer.TenantId | Where-Object {$_.IsInitial -eq $true}
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell-liveid?DelegatedOrg=$($InitialDomain)&BasicAuthToOAuthConversion=true" -Credential $credential -Authentication Basic -AllowRedirection 
    try{
    Import-PSSession $session -DisableNameChecking -ErrorAction Ignore
    } catch{}
            #Check Authsettings
    if($session){
    $SafeLinks = ""
    $SafeLinks = Get-SafeLinksPolicy
    $SafeAttachments = ""
    $SafeAttachments = Get-SafeAttachmentPolicy
    if($SafeLinks ){ 
        write-host "Safe Links policy exist" -ForegroundColor Yellow
        $safelinkpolicy = "true"
	 } else {
        write-host "Safe Links policy does not exist" -ForegroundColor Red
        $safelinkpolicy = "fasle"
     }
 
     if($SafeAttachments ){ 
        write-host "Safe Attachments policy exist" -ForegroundColor Yellow
        $safeattachpolicy = "true"
	  }else {
        write-host "Safe Attachments policy does not exist" -ForegroundColor Red
         $safeattachpolicy = "fasle"
     }
   
    Remove-PSSession $session
    Write-Host "Removed PS Session"
  } else{
    $safelinkpolicy = "fasle"
    $safeattachpolicy = "fasle"
  }
       $properties = @{
                    'Company Name' = $customer.Name
                    'Safe Link Policy Exist' =  $safelinkpolicy
	                'Safe Attachment Policy Exist' = $safeattachpolicy
                    }  
    
    $PropsObject = New-Object -TypeName PSObject -Property $Properties
    $PropsObject | Select-Object  "Company Name", "Safe Link Policy Exist", "Safe Attachment Policy Exist"  | Export-CSV -Path $ATPSettings -NoTypeInformation -Append   
}