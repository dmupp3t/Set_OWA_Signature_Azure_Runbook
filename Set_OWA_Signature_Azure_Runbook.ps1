<#
    .SYNOPSIS
    Set_OWA_Signature_Azure_Runbook.ps1

    .DESCRIPTION
    A powershell script to pull enabled user information using MSGraph and set their OWA Signature via an HTML template variable and ExchangeOnline module. 
   
    .NOTES
    Written by: James Monroe
    Website:    www.jmonroeiv.com
    LinkedIn:   linkedin.com/in/jmonroeiv

    .CHANGELOG
    V1.01, 06/12/2024 - Initial version

    .CREDITS
    Signature templated generated with https://www.mail-signatures.com
    
#>

## Connect to Microsoft Graph API via App Registration - ClientId/TenantId/Thumbprint are saved as Global Runbook Variables
$ClientId = Get-AutomationVariable -Name 'ClientId' 
$TenantId = Get-AutomationVariable -Name 'TenantId'
#Thumbprint is needed because app authorizes using self signed certificate added to Application and Azure Runbook 
$Thumbprint = Get-AutomationVariable -Name 'Thumbprint' 
Connect-MgGraph -clientId $ClientId -tenantId $TenantId -certificatethumbprint $Thumbprint
 
## Example of how to filter single user for testing.
# $users = Get-MgUser -Filter "DisplayName eq '<Display Name>'"

# Get all enabled users that have an email address that ends in 'yourdomain.com' (Replace with your domain)
$users = Get-MgBetaUser -All -Filter "endsWith(mail,'yourdomain.com') and accountEnabled eq true" -Sort "displayName" -ConsistencyLevel eventual -CountVariable CountVar


# HTML Signature Template - Relies on placeholders - For example %%FirstName%% %%LastName%% - Later in the script the for loop will replace the placeholders with information pulled from MSGraph
# Replace block of HTML with your signature template and placeholders you want included with signature
$HTMLsig = @" 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>Email Signature</title>
    <meta content="text/html; charset=utf-8" http-equiv="Content-Type">
  </head>
  <body>
    <table style="width:530px;font-size:9pt;font-family:Arial,sans-serif;line-height:normal;background:0 0!important" cellpadding="0" cellspacing="0">
      <tbody>
        <tr>
          <td style="width:86px;vertical-align:top" valign="top">
            <img border="0" height="86" width="86" style="width:200px;height:86px;border:0" src="https://thumbs.dreamstime.com/b/abstract-vector-logo-your-company-colorful-crossing-orange-red-lines-generic-template-84604765.jpg">
          </td>
          <td style="width:200px;text-align:center;vertical-align:top" valign="top">
            <img border="0" width="11" style="width:11px;height:85px;border:0" src="https://www.mail-signatures.com/signature-generator/img/templates//medium-banner/line.png">
          </td>
          <td style="width:350px;vertical-align:top" valign="top">
            <table cellpadding="0" cellspacing="0" style="background:0 0!important">
              <tbody>
                <tr>
                  <td style="font-size:10pt;font-family:Arial,sans-serif;font-weight:700;color:#3c3c3b;padding-bottom:5px">
                    <span style="font-family:Arial,sans-serif;color:#3c3c3b">%%FirstName%% %%LastName%%</span>
                  </td>
                </tr>
                <tr>
                  <td style="font-size:10pt;font-family:Arial,sans-serif;font-weight:700;color:#3c3c3b;padding-bottom:5px">
                    <span style="font-family:Arial,sans-serif;color:#3c3c3b">%%Title%%</span>
                  </td>
                </tr>
                <tr>
                  <td style="font-size:9pt;font-family:Arial,sans-serif;color:#3c3c3b;padding-bottom:1px">
                    <span style="font-family:Arial,sans-serif;color:#3c3c3b">
                      <span style="font-weight:700">e. </span>%%Email%% </span>
                  </td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>
  </body>
</html>
"@

# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process
# Connect to Azure with system-assigned managed identity
$AzureContext = (Connect-AzAccount -Identity).context
# set and store context
$AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext

# Import the ExchangeOnlineManagement Module we imported into the Automation Account
Import-Module ExchangeOnlineManagement
# Connect to Exchange Online using the Certificate Thumbprint of the Certificate imported into the Automation Account - Replace -Organization with your domain. 
Connect-ExchangeOnline -CertificateThumbPrint $ThumbPrint -AppID $ClientID -Organization "yourdomain.onmicrosoft.com"

# For loop to replace placeholders with information pulled from MSGraph
foreach($user in $users){
    #Temp Variable
    $HTMLSigX = "";
    #Build signature HTML replacing data pulled from MSGraph 
    $HTMLSigX = $HTMLsig.replace('%%FirstName%%', $user.GivenName).replace('%%LastName%%', $user.Surname).replace('%%Title%%', $user.JobTitle).replace('%%PhoneNumber%%', $user.TelephoneNumber).replace('%%MobileNumber%%', $user.Mobile).replace('%%Email%%', $user.Mail).replace('%%Company%%', $user.CompanyName).replace('%%Street%%', $user.StreetAddress).replace('%%City%%', $user.City).replace('%%ZipCode%%', $user.PostalCode).replace('%%State%%', $user.State).replace('%%Country%%', $user.Country) 
    #Set the HTML Signature (NOTE: Roaming Signatures must be disabled or this will have no effect.)
    Set-MailboxMessageConfiguration $user.Mail -SignatureHTML $HTMLSigX -AutoAddSignature $true -AutoAddSignatureOnReply $true 

}




