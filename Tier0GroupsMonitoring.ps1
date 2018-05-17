param (
    [Parameter(Mandatory=$false)][string]$zabbix_url="https://evdetect.cyber.net/ZBX/api/sender.php",
    [Parameter(Mandatory=$false)][string]$domain = "cyber.se",
    [Parameter(Mandatory=$false)][string]$zabbixhost = "InfraDir-ADcore-test-cyber"
)


function TrustCert{
add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
    $AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
    [System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}

if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type)
{
    TrustCert
}

function send-zabbix {

    param(
        [Parameter(Mandatory=$true)]
        [string]$TargetHost,
        [Parameter(Mandatory=$true)]
        $data,
        [Parameter(Mandatory=$false)]
        [string]$Url
    ) 
    $body = ($data.Keys | % { @{host=$TargetHost; key=$_ ; value=$data[$_]} } | ConvertTo-Json)

    if (!$Url) {
        Write-Host "Zabbix API endpoint not specified...`n$($body | Out-String)"
        return
    }

    try {
        Invoke-WebRequest -Uri $Url -Method POST -Body $body
    } catch [exception] {
        $Host.UI.WriteErrorLine("ERROR: $($_.Exception.Message)")
    }
}


function getMember($groupName){    
    $member = @(get-ADGroup -Identity $groupName -Properties * -server "cyber.se"| select -ExpandProperty member)
    return ($member|sort) -join ", "
}

function domaincontroller{
   
    
    $groupsData =@(
        "Account Operators Test"
        "Administrators Test"
        "Backup Operators Test"
        "Domain Admins Test"
        "Enterprise Admins Test"
        "Print Operators Test"
        "Schema Admins Test" 
        "Server Operators Test"
        "Allowed RODC Password Replication Group Test",
        "Certificate Service DCOM Access Test",
        "Cert Publishers Test",
        "Distributed COM Users Test",
        "DnsAdmins Test",
        "Event Log Readers Test",
        "Hyper-V Administrators Test",
        "Protected Users Test",
        "Remote Desktop Users Test",
        "WinRMRemoteWMIUsers__ Test"
    )

    return @{        
        "account-operators-test" = getMember $groupsData[0]
        "administrators-test"= getMember $groupsData[1]
        "backup-operators-test" = getMember $groupsData[2]
        "domain-admins-test"= getMember $groupsData[3]
        "enterprise-admins-test"=getMember $groupsData[4]
        "print-operators-test"=getMember $groupsData[5]
        "schema-admins-test"=getMember $groupsData[6]
        "server-operators-test"=getMember $groupsData[7]
        "allowed-rodc-password-replication-group-test"=getMember $groupsData[8]
        "certificate-service-dcom-access-test"=getMember $groupsData[9]
        "cert-publishers-test"=getMember $groupsData[10]
        "distributed-com-users-test"=getMember $groupsData[11]
        "dnsadmins-test"=getMember $groupsData[12]
        "event-log-readers-test"=getMember $groupsData[13]
        "hyper-v-administrators-test"=getMember $groupsData[14]
        "protected-users-test"=getMember $groupsData[15]
        "remote-desktop-users-test"=getMember $groupsData[16]
        "winrmremotewmiusers__-test"=getMember $groupsData[17]
    }      
}

function checkConnection{
    param([Parameter(Mandatory=$true)][String]$dc)

    $checkConnection = Test-Connection $dc -Quiet -Count 1

    return $checkConnection
}


if($(checkConnection -dc $domain) -eq $true){             
    send-zabbix -TargetHost $zabbixhost -Data $(domaincontroller) -Url $zabbix_url   
}
else{
    send-zabbix -TargetHost $zabbixhost -Data @{"Connection"=$(checkConnection)|Out-String} -Url $zabbix_url
}
