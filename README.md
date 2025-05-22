# Set-OutlookSignature.ps1
$UserName = $env:USERNAME
$SignaturePath = "$env:APPDATA\Microsoft\Signatures"
$SignatureName = "CompanySignature"

# Create HTML signature content
$SignatureHtml = @"
<html>
<body>
<p>Best regards,</p>
<p><b>$UserName</b><br>
Company Name<br>
<font color='gray'>www.company.com</font></p>
</body>
</html>
"@

# Ensure signature folder exists
if (!(Test-Path $SignaturePath)) {
    New-Item -Path $SignaturePath -ItemType Directory | Out-Null
}

# Remove old signature files if they exist
$extensions = @("htm", "rtf", "txt")
foreach ($ext in $extensions) {
    $file = "$SignaturePath\$SignatureName.$ext"
    if (Test-Path $file) {
        Remove-Item $file -Force
    }
}

# Write new signature files
$SignatureHtml | Out-File "$SignaturePath\$SignatureName.htm" -Encoding ASCII
$SignatureHtml | Out-File "$SignaturePath\$SignatureName.rtf" -Encoding ASCII
$SignatureHtml | Out-File "$SignaturePath\$SignatureName.txt" -Encoding ASCII

# Set as default signature in Outlook (Office 365/2019/2016 uses version 16.0)
$OutlookVersion = "16.0"
$OutlookSignatureKey = "HKCU:\Software\Microsoft\Office\$OutlookVersion\Common\MailSettings"

if (!(Test-Path $OutlookSignatureKey)) {
    New-Item -Path $OutlookSignatureKey -Force | Out-Null
}

Set-ItemProperty -Path $OutlookSignatureKey -Name "NewSignature" -Value $SignatureName
Set-ItemProperty -Path $OutlookSignatureKey -Name "ReplySignature" -Value $SignatureName
