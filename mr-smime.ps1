#This script can be used to decrypt "smime.p7m" messages, like those produced by 
# Outlook webmail when you don't have the S/MIME extension installed. 

# The script can be run by right clicking on the file and choosing "Run with PowerShell". 

# In order for the script to work, your computer needs to have all the certificate stuff that it would 
#  normally need in order to access wembail on your computer. 


# Importing crypto stuff
Add-Type -AssemblyName System.Security
# for file picking and saving dialogs
Add-type -AssemblyName System.Windows.Forms

# this "envelope" is a wrapper that handles the decrypting
[System.Security.Cryptography.Pkcs.EnvelopedCms] $envelope = [System.Security.Cryptography.Pkcs.EnvelopedCms]::new()        

Write-Host "Prompting user to select encrypted email..."
# creating a file dialog for user to select encrypted email
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = "$HOME\Downloads"
$OpenFileDialog.filter = “Encrypted Emails (*.p7m)|*.p7m|All Files (*.*)|*.*”
$OpenFileDialog.ShowDialog() | Out-Null
$p7mFile = $OpenFileDialog.filename
if([string]::IsNullOrEmpty($p7mFile)){
    $errorMsg = "Error: no file selected. Please try again."
    Write-Error "$errorMsg. Script will terminate in 20 seconds."
    Start-Sleep -s 20
    exit
}
Write-Host "User selected $p7mFile for decryption"

Write-Host "Loading File..."
#loading file data as a byte array 
$bytes = [System.IO.File]::ReadAllBytes($p7mFile)  

Write-Host "Decoding File..."
# Decode bytes for further crypto processing, I think (?)
[void]$envelope.Decode($bytes)  

## this commented out section is to explicitly choose a certificate to decrypt the email. I don't think it's neccessary.
#$allCerts = Get-Childitem -Path Cert:\CurrentUser\My  
#$mycert = $allCerts[5] 
#[System.Security.Cryptography.X509Certificates.X509Certificate2Collection]$collection = [System.Security.Cryptography.X509Certificates.X509Certificate2Collection]::new()
#[void]$collection.Add($mycert) 
## Decrypt using collection of certs that includes your cert.
#$testFinal = $envelopObj.Decrypt($collection)

Write-Host "Decrypting File. User should see prompt for Cert Selection and/or PIN entry."
Write-Host "PIN entry may not appear if user has already decrypted several emails (cached by Windows/ActivClient)"
try{
    $envelope.Decrypt()
} catch{
    Write-Error "File decryption failed. Please try again. Script will terminate in 20 seconds."
    Start-Sleep -s 20
    exit
}

Write-Host "Decrypted Successfully. Decoding first layer of base64 encoding..."
# Now that the $envelope is decrypted, get the contents and convert to ASCII. the result should be a decrypted SMIME message.
$text = [System.Text.Encoding]::ASCII.GetString($envelope.ContentInfo.Content)

# This is my ghetto method of removing the first four lines of the SMIME message-- these lines should look like:
# Content-Type: application/x-pkcs7-mime; name=smime.p7m; smime-type=signed-data
# Content-Transfer-Encoding: base64
# Content-Disposition: attachment; filename=smime.p7m 
# [[blank line]]
# My method involves finding two line breaks (meaning the last blank line of above) and taking everything after that.
$metaRegex = [regex]::match($text,'(\r\n\r\n)(\r\n)*');
$textTruncated = $text.Substring($metaRegex.Index, $text.Length - $metaRegex.Index)

# this step decodes the base64 into nearly-legible text. At this point, you should be able to see your message but not any attachments, as the attachments are
# in another inception-like layer of base64 encoding. 
# echo $textTruncated

# check if text is encoded in the first place before trying to decode
if($textTruncated -imatch 'Content-Type: text'){ 
	$textDecoded = $textTruncated
}else{
	$textDecoded = [Text.Encoding]::Utf8.GetString([Convert]::FromBase64String($textTruncated))	
}


Write-Host "Prompting user to save raw email data..."
#creating a save file dialog for the raw email data
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.initialDirectory = [io.path]::GetDirectoryName($p7mFile)
$SaveFileDialog.filter = "Raw email text data (*.txt) | *.txt"
$SaveFileDialog.ShowDialog() |  Out-Null
$fileDecoded = $SaveFileDialog.filename

if([string]::IsNullOrEmpty($fileDecoded)){
    $errorMsg = "Error: no file selected. Please try again."
    Write-Error "$errorMsg. Script will terminate in 20 seconds."
    Start-Sleep -s 20
    exit
}
Write-Host "User selected $fileDecoded for raw email data output filename."

Write-Host "Saving raw email data to disk..."
# saving raw email data
Out-File -FilePath $fileDecoded -InputObject $textDecoded


#creating a folder based on the raw email data file name to put all the attachments and stuff
$outFolder = [io.path]::GetFileNameWithoutExtension($fileDecoded)

# this makes all the output files appear in a folder with the same name as the decoded text
$outPath = [io.path]::GetDirectoryName($fileDecoded) + '\' + $outFolder

Write-Host "Creating folder $outFolder to store extracted attachments..."

#creating the folder if it doesn't already exist
New-Item -ItemType Directory -Force -Path $outPath | Out-Null


Write-Host "Extracting attachments..."
# EXTRACTING ATTACHMENTS
# Attachments are stored as ASCII or base64 inside the raw email data. The attachments are separated by unique boundary statements that
# look like boundary="asfu98wf9weu..."

Write-Host "Finding boundary keywords..."
# To find these boundaries, I'm using a regex pattern that searches the location for where these boundaries are defined.
# Content-Type: multipart/.*;\r*\n*\s*boundary="(.*)"
$boundaryRegex = [regex]::matches( $textDecoded, "(----=_NextPart_[0-9A-Za-z_\.]*)")
$boundaries = @()
foreach ($match in $boundaryRegex){
    $boundaries += $match.Groups[1].Value
}
if($boundaries.length -ne 0){
	$boundaries  = $boundaries | select -Unique	
}

foreach ($boundary in $boundaries){
	Write-Host "`tFound Keyword " + $boundary
}
#reversing the boundary array to make the email text content end up appearing first.
[Array]::Reverse($boundaries)
$parts = @()

Write-Host "Searching for attachments..."
# After finding what the boundary definition text is, the next step is to go through the code and pull out blobs of text that represent each attachment
foreach($boundary in $boundaries){
    
    $boundaryEsc = [regex]::escape($boundary)
    
    # this regex is a mess. Here is an explanation of its parts.
    # (?s) : wilcards (*) can match across newlines
    # (?<= BOUNDARYSTUFF)(.*?) : tells regex to find some text (.*?) that is behind BOUNDARYSTUFF, but don't include BOUNDARYSTUFF in match
    # ... (?= BOUNDARYSTUFF) : Same as above but tells regex to find text that is in front of BOUNDARYSTUFF but don't include BOUNDARYSTUFF in match
    # BOUNDARYSTUFF : I do a search for two dashes "--" followed by the boundary text, followed by newlines
    $partsRegex = [regex]::matches($textDecoded,'(?s)' + '(?<=' + '--'+ $boundaryEsc + '\r\n' + ')' + '(.*?)' + '--' + '(?=' + $boundaryEsc + ')')

    foreach( $match in $partsRegex){
        #this if statement is a lazy way of ignoring multi-part matches. The code logic from above should still catch the subparts.
        if($match.Groups[1].Value -notmatch 'Content-Type: multipart/.*;'){
            $parts += $match.Groups[1].Value
            Write-Host "`tFound Attachment"
        }
        
    }
}

Write-Host "Formatting and Saving Attachments..."
# after getting each attachment (part), each part needs to be formatted, decoded (if base 64), and saved to a file.
$n = 0
foreach ($part in $parts){
    $contentRegex = [regex]::match($part,'\r\n\r\n');
    $contentMeta = $part.Substring(0,$contentRegex.Index);
    $data = $part.Substring($contentRegex.Index, $part.Length - $contentRegex.Index)

    $encoding = '';
    # not going to bother checking for encoding type (ascii vs something else), gonna assume ascii.
    if($contentMeta -match "Content-Type: text"){
        $encoding = 'text'
    }
    if($contentMeta -match "Content-Transfer-Encoding: base64"){
        $encoding = 'base64'
    }

    if($contentMeta -match 'name="(.*)"'){
        $outFilename = 'File ' + $n + " " + $Matches[1]
    }elseif ($encoding -eq 'text'){
        $outFilename = 'File ' + $n + '.txt'
    }else{
        $outFilename = 'File ' + $n + '.bin'
    }
    
    
    if($encoding -eq 'base64'){
        Write-Host "Saving File (decoding out of base64):"
        #Finding pairs of invalid base64 characters and removing them. Also removing newlines.
        $dataFixed = ($data -replace '\r\n') -replace '[^a-zA-Z0-9+/=].'
        $dataBinary = [Convert]::FromBase64String($dataFixed)
        
        Write-Host "`t $outFilename"
        [System.IO.File]::WriteAllBytes("$outPath\$outFilename", $dataBinary)
        
    }else{
        Write-Host "Saving file:"
        Write-Host "`t $outFilename..."
        Out-File -FilePath "$outPath\$outFilename" -InputObject $data
    }

    $n += 1;
}
Write-Host
Write-Host "Script Complete. All email files have been saved to $outPath."
Write-Host "The raw unencrypted email data is located here: $fileDecoded"
Write-Host
Write-Host "This window will close in 60 seconds."
Start-Sleep -s 60
exit
