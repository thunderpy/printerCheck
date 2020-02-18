## Login to webpage of printer and get Printer error details.
## Writen for Canon imageRUNNER1435

<#

Displaying the toner level:

The toner level will be described by one of the following:
    <OK>

    <Low>
The message <Prepare the toner cartridge.> appears in the display.

    <None>
The message <Replace the toner cartridge.> appears in the display.

#>

$printers = @{
    
    'http://printerIP/login.html'  = "Printer Name"

    'http://printerIP/login.html' = "Printer Name"

}

# Mail Body Data
$MailBody=@()

# Mail Function
function Send_Mail([string]$email) {

    $message = new-object Net.Mail.MailMessage;

    $message.From = "web-admin@Test.com";
    
    $message.To.Add($email);

    $message.Subject = "Printer Alert!!!!";
    
    $message.Body = $MailBody;
    
    $smtp = new-object Net.Mail.SmtpClient("SMTP Server IP");
    
    $smtp.send($message);
    
    write-host "Mail Sent" ; 
    
 }


foreach ($printer in $printers.GetEnumerator()){

    $printerIP = $printer.Key
    $printerName = $printer.Value
    $extractIP = ($printerIP | Select-String -Pattern "\d{1,3}(\.\d{1,3}){3}" -AllMatches).Matches.Value

    $ping = (Test-Connection -ComputerName $extractIP -Count 1 -Quiet)

    if ($ping){

        Write-Host "Checking $printerIP...... Please Wait!!!"

        $ie = New-Object -ComObject InternetExplorer.Application
        $ie.visible=$false
        $ie.Navigate($printerIP)
        while( $ie.ReadyState -ne 4 ) { Start-Sleep 20 }
        Start-Sleep 5
        $ie.document.getElementById("i0019").value="Username"
        start-sleep 5
        $ie.document.getElementById("submitButton").Click()
        while($ie.Busy){ Start-Sleep 10 }
        $body = $ie.Document.body
        start-sleep 5
        #$ie.Quit()

        # Check printer error
        [string]$errorDiv = $body.Document.getElementById("deviceErrorInfoModule").innerText
    
        <#
        if ($errorDiv.Contains("No paper") -or $errorDiv.Contains("Paper is jammed")){

            Write-Host "Error in printer please check!!!!"

        }
        #>

        $noPaper=""
    
        if ($errorDiv.Contains("No paper")){

            $noPaper="No Paper in printer $printerName."
            $MailBody += "==> " + "$noPaper `n"  

        }

        $paperJam=""

        if ($errorDiv.Contains("Paper is jammed.")){

            $paperJam="Paper is jammed in $printerName."
            $MailBody += "==> " + "$paperJam `n"
        }

        # Check Tonner Status
        [string]$tonnerDiv = $body.Document.getElementById("tonerInfomationModule").innerText
        if ($tonnerDiv.Contains("OK")){

            continue

        }
        if($tonnerDiv.Contains("Low")){
    
            $MailBody += "==> " + "Prepare the toner cartridge for $printerName toner level is low!. `n"

        }
        if($tonnerDiv.Contains("None")){
    
            $MailBody += "==> " + "Replace the toner cartridge for $printerName. `n"
        }

        Start-Sleep 3

    }else {

        $MailBody += "==> " + "Not able to access $printerIP. `n"
        		
    }
      
}


Start-Sleep 1

Write-Host "Mail body content is "$MailBody

# Check $MailBody variable if not empty send mail.

if (![string]::IsNullOrEmpty($MailBody)){
    
    Send_Mail -email "myemail@email.com"
}

Start-Sleep 1
Clear-Variable MailBody


# Stop process if running.
$checkProcess = "iexplore", "dllhost", "ielowutil"

foreach($process in $checkProcess){
    
   if((get-process $process -ea SilentlyContinue) -eq $Null){
     
        "$process Not Running" 
    }

else{ 
    
    Stop-Process -Name $process -Force
    
 }

}
 
# Script END