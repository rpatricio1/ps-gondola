$username = "rvs@ameri-kleen.com"
$password = "M0unt@1nV13w"
$securestring = New-Object -TypeName System.Security.SecureString
$password.ToCharArray() | ForEach-Object {$securestring.AppendChar($_)}
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securestring

$Sku = @{ 
    "MCOIMP" = "Lync Online (Plan 1)"
    "MCOSTANDARD" = "Lync Online (Plan 2)"
    "MCOVOICECONF" = "Lync Online (Plan 3)"
    "OFFICESUBSCRIPTION" = "Office Professional Plus"
    "DESKLESSPACK" = "Office 365 (Plan K1)" 
    "DESKLESSWOFFPACK" = "Office 365 (Plan K2)" 
    "LITEPACK" = "Office 365 (Plan P1)" 
    "EXCHANGESTANDARD" = "Exchange Online (Plan 1)"
    "EXCHANGEENTERPRISE" = "Exchange Online (Plan 2)"
    "EXCHANGEARCHIVE" = "Exchange Online Archiving"
    "EXCHANGEDESKLESS" = "Exchange Online Kiosk"
    "EXCHANGETELCO" = "Exchange Online POP"
    "STANDARDPACK" = "Microsoft Office 365 (Plan E1)"
    "STANDARDWOFFPACK" = "Microsoft Office 365 (Plan E2)"
    "ENTERPRISEPACK" = "Microsoft Office 365 (Plan E3)"
    "ENTERPRISEPACKLRG" = "Microsoft Office 365 (Plan E3)"
    "ENTERPRISEWITHSCAL" = "Microsoft Office 365 (Plan E4)"
    "STANDARDPACK_STUDENT" = "Microsoft Office 365 (Plan A1) for Students"
    "STANDARDWOFFPACKPACK_STUDENT" = "Microsoft Office 365 (Plan A2) for Students"
    "ENTERPRISEPACK_STUDENT" = "Microsoft Office 365 (Plan A3) for Students"
    "ENTERPRISEWITHSCAL_STUDENT" = "Microsoft Office 365 (Plan A4) for Students"
    "STANDARDPACK_FACULTY" = "Microsoft Office 365 (Plan A1) for Faculty"
    "STANDARDWOFFPACKPACK_FACULTY" = "Microsoft Office 365 (Plan A2) for Faculty"
    "ENTERPRISEPACK_FACULTY" = "Microsoft Office 365 (Plan A3) for Faculty"
    "ENTERPRISEWITHSCAL_FACULTY" = "Microsoft Office 365 (Plan A4) for Faculty"
    "ENTERPRISEPACK_B_PILOT" = "Microsoft Office 365 (Enterprise Preview)"
    "STANDARD_B_PILOT" = "Microsoft Office 365 (Small Business Preview)"
    "DESKLESSPACK_GOV" = "Microsoft Office 365 (Plan K1) for Government"
    "DESKLESSWOFFPACK_GOV" = "Microsoft Office 365 (Plan K2) for Government" 
    "OFFICESUBSCRIPTION_GOV" = "Office Professional Plus for Government"
    "EXCHANGESTANDARD_GOV" = "Microsoft Office 365 Exchange Online (Plan 1) only for Government"
    "EXCHANGEENTERPRISE_GOV" = "Microsoft Office 365 Exchange Online (Plan 2) only for Government"
    "STANDARDPACK_GOV" = "Microsoft Office 365 (Plan G1) for Government"
    "STANDARDWOFFPACK_GOV" = "Microsoft Office 365 (Plan G2) for Government"
    "ENTERPRISEPACK_GOV" = "Microsoft Office 365 (Plan G3) for Government"
    "ENTERPRISEWITHSCAL_GOV" = "Microsoft Office 365 (Plan G4) for Government"
    "VISIOCLIENT" = "Visio Pro for Office 365"
    "INTUNE_A" = "Intune for Office 365"
    "PROJECTONLINE_PLAN_1" = "Project Online Plan 1"
    "PROJECTONLINE_PLAN_2" = "Project Online Plan 2"
    "PROJECTCLIENT" = "Project Pro for Office 365"
    "WACSHAREPOINTSTD" = "Office Web Apps with SharePoint Plan 1"
    "WACSHAREPOINTENT" = "Office Web Apps with SharePoint Plan 2"
    "SHAREPOINTSTANDARD" = "SharePoint Online (Plan 1)"
    "SHAREPOINTENTERPRISE" = "SharePoint Online (Plan 2)"
    "SHAREPOINTDESKLESS" = "SharePoint Online Kiosk"
    "SHAREPOINTPARTNER" = "SharePoint Online Partner Access"
    "SHAREPOINTSTORAGE" = "SharePoint Online Storage" 
    "POWER_BI_STANDALONE" = "Power BI for Office 365"
} 
$allUsers = @()

Connect-MsolService -Credential $credential
$AccountSkus = Get-MsolAccountSku
$MSOLUsers = Get-MsolUser | Where-Object {$_.isLicensed -eq "TRUE"} 
foreach ($MSOLUser in $MSOLUsers)
{
    Write-Progress -activity "Gathering license statistics...." -status $MSOLUser.Displayname 
    $AccountSkus = $MSOLUser.Licenses 
    foreach ($AccountSku in $AccountSkus)
    {
        $obj = New-Object PSObject -Property @{"DisplayName" = $MSOLUser.displayname; "Title" = $MSOLUser.Title; "License" = $Sku.Item($AccountSku.AccountSkuId.Split(":")[1]) }
        $allUsers = $allUsers += $obj
    }
}
$allUsers | Sort-Object DisplayName, License

