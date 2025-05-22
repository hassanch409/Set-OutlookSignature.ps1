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

# Create signature folder if it doesn't exist
if (!(Test-Path $SignaturePath)) {
    New-Item -Path $SignaturePath -ItemType Directory | Out-Null
}

# Write signature files
$SignatureHtml | Out-File "$SignaturePath\$SignatureName.htm" -Encoding ASCII
$SignatureHtml | Out-File "$SignaturePath\$SignatureName.rtf" -Encoding ASCII
$SignatureHtml | Out-File "$SignaturePath\$SignatureName.txt" -Encoding ASCII
