# Ask for credentials
$UserCredential = Get-Credential

# Initilize result variable as array
$result = @()

# Connect to Exchange Online
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

#Get Distribution Groups and its members
$groupType = "Distribution List"
$groups = (Get-DistributionGroup | Sort-Object DisplayName)
foreach($group in $groups){
    $groupMembers = Get-DistributionGroupMember -Identity $group.PrimarySmtpAddress
    foreach ($groupMember in $groupMembers){
        $properties = @{
            "GroupType" = $groupType;
            "GroupName" = $group.DisplayName;
            "GroupEmail" = $group.PrimarySmtpAddress;
            "GroupAccess" = "";
            "MemberName" = $groupMember.DisplayName;
            "MemberEmail" = $groupMember.PrimarySmtpAddress;
            "MemberType" = $groupMember.RecipientType;
        }
        $memberObject = New-Object -TypeName PSObject -Property $properties
        $result += $memberObject
    }
}

#Get Dynamic Distribution Groups and its members
$groupType = "Dynamic Distribution List"
$groups = (Get-DynamicDistributionGroup | Sort-Object DisplayName)
foreach($group in $groups){
    $groupMembers = Get-Recipient -RecipientPreviewFilter $group.RecipientFilter -OrganizationalUnit $group.RecipientContainer
    foreach ($groupMember in $groupMembers){
        $properties = @{
            "GroupType" = $groupType;
            "GroupName" = $group.DisplayName;
            "GroupEmail" = $group.PrimarySmtpAddress;
            "GroupAccess" = "";
            "MemberName" = $groupMember.DisplayName;
            "MemberEmail" = $groupMember.PrimarySmtpAddress;
            "MemberType" = $groupMember.RecipientType;
        }
        $memberObject = New-Object -TypeName PSObject -Property $properties
        $result += $memberObject
    }
}

#Get Office 365 Groups
$groupType = "Office 365 Group"
$groups = (Get-UnifiedGroup | Sort-Object DisplayName)
foreach($group in $groups){
    $groupMembers = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members
    foreach ($groupMember in $groupMembers){
        $properties = @{
            "GroupType" = $groupType;
            "GroupName" = $group.DisplayName;
            "GroupEmail" = $group.PrimarySmtpAddress;
            "GroupAccess" = $group.AccessType;
            "MemberName" = $groupMember.DisplayName;
            "MemberEmail" = $groupMember.PrimarySmtpAddress;
            "MemberType" = $groupMember.RecipientType;
        }
        $memberObject = New-Object -TypeName PSObject -Property $properties
        $result += $memberObject
    }
}

# Disconnect from Exchange Online
Remove-PSSession $Session

# Return data
$result



