function Get-scADUserPasswordInfo {
    param($SamAccountName, $RootDN, $DomainController)

$filter = "samaccountName -eq '$SamAccountName'"
$now = get-date

get-aduser -filter $filter -server $DomainController -SearchBase $RootDN –Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed", passwordlastset, passwordneverexpires  | `
            Select-Object -Property SamAccountName,DisplayName,  PasswordNeverExpires , PasswordLastSet `
            ,@{Name="PasswordExpirationDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} `
            ,@{Name="ExpireInXDays";Expression={ ( New-TimeSpan –Start $now –End ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed"))).days}}
}


#require modules activedirectory


$domainName =  $env:USERDNSDOMAIN

[STRING]$DomainController = (Get-ADDomainController -DomainName $DomainName -discover).hostname
[String]$Dn = (get-addomain $DomainName -Server $DomainController).DistinguishedName

Get-scADUserPasswordInfo -SamAccountName LESIRESY -RootDN $Dn -DomainController $DomainController