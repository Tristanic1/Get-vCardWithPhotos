#Requires -Modules ActiveDirectory

# Connect to Exchange
$ExchangeServer = 'Exchange01'
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
Import-PSSession $session -CommandName Get-Userphoto

$filename = 'C:\Temp\Contoso.vcf'

#test to see if vcard file already exists
If (!(Test-Path $filename)){
    #if not then create the file
    New-Item -Path $filename -ItemType File -Force -Encoding ([System.Text.Encoding]::UTF8)
}

# Get AD Users from specified OU and build vCards
Get-ADUser -SearchBase "OU=AD Users,DC=contoso,DC=com" -Filter {(Enabled -eq $true)} -ResultSetSize $null -Properties sn,givenName,displayName,company,department,title,telephoneNumber,mobile,mail | ForEach-Object {
            
            if ($_.telephoneNumber) {
                
                # Generate vCard
                Add-Content -Path $filename "BEGIN:VCARD"
                Add-Content -Path $filename "VERSION:2.1"
                Add-Content -Path $filename ("N;LANGUAGE=en-us:" + $_.sn + ";" + $_.givenName)
                Add-Content -Path $filename ("FN:" + $_.displayName)
                Add-Content -Path $filename ("ORG:" + $_.company + ";" + $_.department)
                Add-Content -Path $filename ("TITLE:" + $_.title)
                Add-Content -Path $filename ("TEL;WORK;VOICE:" + $_.telephoneNumber)
                Add-Content -Path $filename ("TEL;CELL;VOICE:" + $_.mobile)
                Add-Content -Path $filename ("EMAIL;PREF;INTERNET:" + $_.mail)

                # Get high quality photo from Exchange and convert to Base64
                $photo = Get-UserPhoto -Identity $_.Name -ErrorAction SilentlyContinue
                if ($photo -ne $null) {$photo = [convert]::ToBase64String($photo.PictureData)}
                Add-Content -Path $filename ("PHOTO;ENCODING=b;TYPE=JPEG:" + $photo)

                Add-Content -Path $filename "END:VCARD"
            }
        }

Remove-PSSession $session