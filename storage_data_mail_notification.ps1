#Asking for Vcenter IP
$VC = Read-Host "`nPlease enter Vcenter IP"

#Connecting to Vcenter 
connect-viserver $VC -credential (get-credential)

Start-Sleep -Seconds 2

#Adding text for understading 
write-host "`r`n"
write-host "#################### Existing hard disk details of VMs: "  -ForegroundColor DarkRed -BackgroundColor Yellow
write-host "`r`n"

#Creating a Variable and assigned a value(0) to it ($null =0 powershell special variable)
$disk=$null

#Creating a Variable and assigning an empty array to the variable (@()= way to create an empty array)\
$vmDiskInfo = @()

#foreach loop in PS that reads the lines from a file and [System.IO.File]::ReadLines() is framework method that reads all the lines in a file
#The loop iterates over each line of the array and assign it to the variable $line on each iteration
foreach($line in [System.IO.File]::ReadLines("$pwd\VMs.txt"))
{
$disk= Get-VM $line  | Get-Harddisk | Select Name,CapacityGB,StorageFormat 
$result = "Disk capacity for $line is $($disk -join ',')"
$vmDiskInfo +=[PSCustomObject]@{
    VMName = $line
    DiskCapacity = ($disk -join ',').Replace('@','')
}
write-host "`r`Copying result to excel..................................."-ForegroundColor Magenta -BackgroundColor black
}

# Below command will redirect the outpu to excel file
$file = $vmDiskInfo | Export-Excel -path $pwd\VM-DiskChange_1.xlsx -AutoSize -AutoFilter

#powershell will disconnect from the Vcenter
Disconnect-VIServer -Server $VC -Confirm:$False; 1> $null
write-host "#################### SCRIPT COMPLETED ################" -ForegroundColor DarkRed -BackgroundColor Green

# adding text 
write-host "`r`n"
write-host "Excel file has been created and we are sending it to given mail id" -ForegroundColor White -BackgroundColor black
write-host "`r`n"
Start-Sleep -Seconds 2
write-host "preparing the mail...................." -ForegroundColor white -BackgroundColor red
write-host "`r`n"
Start-Sleep -Seconds 2

#Password of our email
$passwd ="ENTER YOUR OUTLOOK PASSWORD HERE"
$passwd_enc = ConvertTo-SecureString $passwd -AsPlainText -Force  #secure you password

#your email account
$sendermail = "your mail id"

#email ID to whom you need to send
$recivermail = "o whom you need to send the mail type maid id"
#if you want to add another mail id use , and add another mail id


#Body of the mail 
$body = "This is the test mail for automation"
$attachment ="$pwd\VM-DiskChange_1.xlsx"

write-host "Sending mail to the users............................."
Start-Sleep -Seconds 2
#properties of the mail

$props = @{
    SMTPServer = "Enter you SMTP SERVER " #if you don't know type this in power shell: Resolve-DnsName -Type MX example.com | Sort-Object Priority | Select-Object -First 1 | Format-Table NameExchange -Auto

    From = $sendermail
    To = $recivermail
    Subject = "Testing email from powershell"
    port = 587 
    Credential = New-Object System.Management.Automation.PSCredential($sendermail,$passwd_enc)
}

send-MailMessage @props $body -Attachment $attachment
write-host "`r`n"
write-host "########## sent the mail to the user ########## "  -ForegroundColor red -BackgroundColor Green
