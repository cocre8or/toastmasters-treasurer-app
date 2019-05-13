#Set up paths and directories
#Directory of Template files
$startUpFile = "D:\TM\ITApp\startup.csv"

#Set Globals
$clubName;
$companyName;
$logFile;
$directoryOfTemplateFiles;
$absolutePathForRoster;
$treasurerName;
$treasurerPhone;
$treasurerTitle;
$treasurerEmail;
$treasurerEmailSignature;
$VpOfMemberShipName;
$VpOfMemberShipPhone;
$nextTermStartDate;
$nextTermEndDate;

$clubAddress;
$clubUrl;


#Set Last meeting of the term date -  typically last Wednesday
$LastMeetingDateForTerm;
$today = get-date -format MM-dd-yyyy

#Employee types
$contractor = 'C'
$employee = 'P'

$startupInfo = Import-CSV $startUpFile

foreach($info in $startupInfo) {
       
	$logFile = $info.logFile
	$directoryOfTemplateFiles = $info.directoryOfTemplateFiles
    $absolutePathForRoster = $info.absolutePathForRoster
	
	#Set Treasurer information
	$treasurerName = $info.treasurerName
	$treasurerPhone = $info.treasurerPhone
	$treasurerTitle = $info.treasurerTitle
	$treasurerEmail = $info.treasurerEmail
	$treasurerEmailSignature = $info.treasurerEmailSignature
	
	#Set VP of Membership Information
	$VpOfMemberShipName = $info.vpOfMemberShipName
	$VpOfMemberShipPhone = $info.vpOfMemberShipPhone
	
	$clubName = $info.clubName
	$companyName = $info.companyName
	$clubAddress = $info.clubAddress
	$clubUrl = $info.clubUrl
	$nextTermStartDate = $info.nextTermStartDate
	$nextTermEndDate = $info.nextTermEndDate
	$LastMeetingDateForTerm = $info.LastMeetingDateForTerm
}


#Store the roster information
$roster = Import-CSV $absolutePathForRoster


<#
Person Name, Transaction Name, Date

#>
function LogInfo {
   [CmdletBinding()]
    param(
             [Parameter(Position=0)]
             [string] $person,
             [Parameter(Position=1)]
             [string] $transaction
    )
    process {
        
        $info = "Member: $($person) `t Transaction: $($transaction) `t Date: $($today)"

        if($logFile.Length -lt 1) {

            New-Item $logFile -ItemType file

        }else {
            
            Add-Content $logFile -Value $info
            Write-Host $info
        }
    }
}


function IsNewMemberRolesFulfilledWaived {
    [CmdletBinding()]
    param(
             [Parameter(Position=0)]
             [string] $rolesFilled,
             [Parameter(Position=1)]
             [string] $daysAsMember
    )
    process {
        if([int]$rolesFilled -lt 4 -and [int]$daysAsMember -gt 183){

           return 0
        }
        return 1
    }
}


function Send-Email {
    [CmdletBinding()]
    param(
         [Parameter(Position=0)]
         [string] $text,
         [Parameter(Position=1)]
         [string] $regex,
         [Parameter(Position=2)]
         [string] $validationMessage,
         [Parameter(Position=3)]
         [string] $readMessage
    )
    process {

        if ($text -notmatch $regex) {
            Write-Host $text " was invalid." $validationMessage
            $text = Read-Host $readMessage
            Send-Email -text $text -regex $regex -validationMessage $validationMessage -readMessage $readMessage
        }
    }
}

function Validate-Option {
    [CmdletBinding()]
    param(
         [Parameter(Position=0)]
         [string] $text
    )
    process {
        $range = "1-7"
        $regex = "^[$($range)]$"
        if ($text -notmatch $regex) {

            Write-Host $text " was invalid. Please enter a number from $($range).`r`n"
            Start-App
        }

        return 0;
    }
}

function Select-Option {
    [CmdletBinding()]
    param(
         [Parameter(Position=0)]
         [string] $option
    )
    
    process {

         if((Validate-Option -text $option) -eq 0){
      
              if($option -eq 1){

                #view roster
                $roster | Format-Table 'Count','Status','Name','Member Number','LAN ID', 'Cost Center','Emp. Type','Email','Member  Since','Roles Filled','AmtPaid'
              }
              if($option -eq 2){

                  #create batch invoice
                  Create-Batch-Invoice
              }
              if($option -eq 3){

                 #create batch reminders
                  Create-Batch-Reminder
              }
              if($option -eq 4){

                  #create reciept
                  Write-Host 'Who would you like to create a reciept email for?'
                  $name = Read-Host
                  Create-Reciept -person $name
              }
              if($option -eq 5){

                Write-Host 'Who would you like to create an invoice email for?'
                  #invoice
                 $name = Read-Host
                 Create-Invoice -person $name
              }
              if($option -eq 6){

                  Write-Host 'Who would you like to create a reminder email for?'
                   #reminder
                   $name = Read-Host
                   Create-Reminder -person $name
              }
              if($option -eq 7){

                #Exit
                exit(1)
              }
         }
         Start-App
     }
}

function Find-Member {
    [CmdletBinding()]
    param(
         [Parameter(Position=0)]
         [string] $person
    )
    process {

         #Find the persons
        $item = $roster | Where-Object {$_.name -Match "^$person.*" -and $_.status -eq 'A' } | Select-Object -First 1

        #If the person is found
        if($item) {
            
            return $item

        }else{

            Write-Host "$($person) is not found.`r`n"
            Start-App
        }
    }
    
}



#----------------------------------CREATE AN EMAIL------------------------------

function Create-Email {
    [CmdletBinding()]
    param(
         [Parameter(Position=0)]
         [string] $person,
         [Parameter(Position=1)]
         [string] $subject,
         [Parameter(Position=2)]
         [string] $body,
         [Parameter(Position=3)]
         [string] $attachment
    )
    process {

        #Find the persons
        $item = Find-Member -person $person

            
        #Get the person's email address
        $emailAddress = $item.Email

        #Create an outlook application object
        $outlook = New-Object -comObject Outlook.Application

        #Create an email item in via outlook and set the email properties
        $mail = $outlook.CreateItem(0)
        $mail.To = "$emailAddress"
        #$mail.Cc = "$treasurerEmail"
        $mail.Subject = "$subject"
        $mail.Body = $body + $treasurerEmailSignature

        #if a file attachment exists add it to the email
        if($attachment.length -gt 0) {

           $mail.Attachments.Add($attachment)
        }

        #Save the email
        $mail.save()

        #View the email
        $inspector = $mail.GetInspector
        $inspector.Display()
    }
}

#----------------------------------CREATE RECEIPT------------------------------
function Create-Reciept {
    [CmdletBinding()]
    param(
         [Parameter(Position=0)]
         [string] $person
    )
    process {

        #Find the persons
        $item = Find-Member -person $person

        #Set up email subject and body for the receipt
        $subject = "$($clubName) Dues Receipt"
        $body = "This email confirms that I received `$$($item.AmtPaid).00 for $($clubName) dues, which is `$5.00 per term.`r`n"

        if([int]$item.AmtPaid -eq 10) {

            $body += "The additional `$5.00 will go towards the following term."
        }
        

        #Create the email
        Create-Email -person $person -subject $subject -body $body -attachment ""

        #Log details
        LogInfo -person $item.Name -transaction "Received $($item.AmtPaid).00 toward dues."

    }
}

#----------------------------------CREATE INVOICE AN EMAIL------------------------------
function Create-Invoice {
    [CmdletBinding()]
    param(
         [Parameter(Position=0)]
         [string] $person
    )
    process {

        #Find the persons
        $item = Find-Member -person $person
            
        #Create a Word Doc object
        $wordApp = New-Object -ComObject Word.Application

        #Optionally hide or show the word document
        $wordApp.Visible = $false

        #Get the number of roles filled
        $rolesFulfilled = $item.'Roles Filled'
        if($rolesFulfilled -eq "") {
    
            $rolesFulfilled = 0
        }

        $daysAsMember = (New-TimeSpan -Start $item.'Member  Since' -End $today).Days
        $name = $item.Name
        $memberName = $name.Replace(" ","_")
        $subject = "$($clubName) Dues"
        $emailBody = "Attached is your dues invoice for the upcoming Toastmasters membership term ($($nextTermStartDate) to $($nextTermEndDate)).`r`nClub records show you have fulfilled $($rolesFulfilled) role(s) (in separate meetings) during the current membership term.`r`n"

        #Check if person has been a member for more than 2 months
        if([int]$daysAsMember -gt 60) {

            $emailBody = $emailBody + "Please let me know if you intend to renew your membership in $($clubName).`r`n"
            $emailBody = $emailBody + "Please deliver your dues payment to $($treasurerName).  I will email a receipt when I receive your payment.`r`n"
            
            #Check if roles filled and Full-time Employee
            if([int]$rolesFulfilled -lt 4 -and $item.'Emp. Type' -eq $employee) {

                 $emailBody = $emailBody + "Please pay this invoice ASAP.  If you finish fulfilling at least 4 roles by the end of this term, $($clubName) will refund the Toastmasters International portion of your dues payment."
            }
        }

        #Set up invoiceTemplates based on the certain criteria


        #Check if contractor because they have to pay their own TM fees
        if($item.'Emp. Type' -eq $contractor) {

           $invoiceTemplateFile = "Billing Statement_Contractor.docx"
        }
        elseif([int]$daysAsMember -gt 60) {

            $isWaived = IsNewMemberRolesFulfilledWaived -rolesFilled $rolesFulfilled -daysAsMember $daysAsMember

            if($isWaived -eq 0) {

                $invoiceTemplateFile = "Billing Statement_InfrequentParticipant.docx"

            } else {

                $invoiceTemplateFile = "Billing Statement_FrequentParticipant.docx"
            }

        } else {

            $invoiceTemplateFile = "Billing Statement_FrequentParticipantPaid.docx"

        }

        #Create the invoice document
        $invoiceDoc = $wordApp.Documents.Open($directoryOfTemplateFiles + "\\" + $invoiceTemplateFile)


        #POPULATE INVOICE FIELDS - These are text FormFields in the Word document templates
        $invoiceDoc.FormFields("Name").Result = "$name"
		$invoiceDoc.FormFields("ClubName").Result = "$clubName"
		$invoiceDoc.FormFields("ClubName_1").Result = "$clubName"
		$invoiceDoc.FormFields("ClubName_2").Result = "$clubName"
		$invoiceDoc.FormFields("ClubAddress").Result = "$clubAddress"
		$invoiceDoc.FormFields("ClubUrl").Result = "$clubUrl"
		$invoiceDoc.FormFields("CompanyName").Result = "$companyName"
		$invoiceDoc.FormFields("TreasurerTitle").Result = "$treasurerTitle"
        $invoiceDoc.FormFields("TreasurerName").Result = "$treasurerName"
        $invoiceDoc.FormFields("TreasurerPhone").Result = "$treasurerPhone"
        $invoiceDoc.FormFields("VpMembershipName").Result = "$VpOfMemberShipName"
        $invoiceDoc.FormFields("VpMembershipPhone").Result = "$VpOfMemberShipPhone"
        $invoiceDoc.FormFields("ProgTermRange").Result = "$($nextTermStartDate) to $($nextTermEndDate)"
        $invoiceDoc.FormFields("TmTermRange").Result = "$($nextTermStartDate) to $($nextTermEndDate)"
		
        

        #Store the file in the template directory as a PDF
        $invoiceFilename = $directoryOfTemplateFiles + "\Invoice_" + $memberName + ".pdf"
        $invoiceDoc.SaveAs([ref] $invoiceFilename, [ref] 17)

        #Close the document
        $invoiceDoc.Close()

        #Create the email and attach the invoice PDF
        Create-Email -person $person -subject $subject -body $emailBody -attachment $invoiceFilename

        #Close the word document
        $wordApp.Quit()

        #Log details
        LogInfo -person $item.Name -transaction "The $($invoiceTemplateFile) was sent."

        
     }
}

#----------------------------------CREATE REMINDER EMAIL------------------------------

function Create-Reminder {
    [CmdletBinding()]
    param(
         [Parameter(Position=0)]
         [string] $person
    )
    process {
       
       #Find the persons
       $item = Find-Member -person $person

       $rolesFulfilled = $item.'Roles Filled'

       $greeting = "This is a friendly mid-term reminder to participate in $($clubName) Toastmasters meetings.`r`n"
       $roleMsg = "Our attendance records show you have fulfilled $($rolesFulfilled) role(s) so far during the current membership term.`r`n"
       $expectationMsg = "$($companyName) pays your Toastmasters International dues with the expectation that you will fulfill at least 4 meeting roles during each 6 month term.`r`n"
       $lastTermMeeting = "The last meeting of the current term is $($LastMeetingDateForTerm).  Please plan accordingly."

       $subject = "$($clubName) Dues Reminder"
       $body = $greeting + $roleMsg + $expectationMsg + $lastTermMeeting

       #Create email no attachment needed
       Create-Email -person $person -subject $subject -body $body -attachment ""

       #Log details
       LogInfo -person $item.Name -transaction "The $($subject) was sent."

    }
}

#--------------------------------CREATE BATCH INVOICE----------
function Create-Batch-Invoice {
    [CmdletBinding()]
    param()

     process {

        foreach($person in $roster) {
            
            if($person.status -eq 'A') {
                
                Create-Invoice -person $person.name

            }
        }
    }
}

#--------------------------------CREATE BATCH REMINDER----------
function Create-Batch-Reminder {
    [CmdletBinding()]
    param()

     process {

        foreach($person in $roster) {
            
            if($person.status -eq 'A' -and [int]$person.'Roles Filled' -lt 4) {
                
                Create-Reminder -person $person.name

            }
        }
    }
}



function Start-App {
    <#
    .SYNOPSIS
    This is the starting point of the Treasurer application
    .DESCRIPTION
    The Start-App command gives you a list of options to choose from to handle Treasurer duties.
    #>

    Write-Host "$($clubName) Treasurer Application."
    Write-Host "Here are your options."
    Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    Write-Host "1. View Roster"
    Write-Host "2. Create batch invoices for Active members."
    Write-Host "3. Create batch reminder emails for Active members who filled less than 4 roles."
    Write-Host "4. Create a receipt for member dues email."
    Write-Host "5. Create invoice for a member."
    Write-Host "6. Create a reminder email."
    Write-Host "7. Exit"

    $option = Read-Host
    Select-Option $option

}

Start-App






