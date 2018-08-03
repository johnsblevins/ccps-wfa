$dir = 'C:\Users\jsble\OneDrive\ccps\Email PIA\Working1'

cd $dir

$files = get-childitem -Recurse -Filter "*eml"

$emails = @()

foreach ($file in $files)
{
    $recipientList = ""
    $calvertEmailList = ""
    $3rdPartyEmailList = ""
    $recipientCount = 0
    $calvertEmailCount = 0
    $3rdPartyEmailCount = 0

    $email = new-object System.Object
    $email | add-member -type NoteProperty -name Folder -value $file.DirectoryName
    $email | add-member -type NoteProperty -name File -value $file.Name
    $email | add-member -type NoteProperty -name FullPath -value $file.FullName
    $email | add-member -type NoteProperty -name Size -value $file.length

    $eml = gc $email.FullPath
    $eml | Select-String '^Subject:'
    $eml | Select-String '^From:'    
    $eml | Select-String '^To:' 
    $eml | Select-String '^CC:' 

    $subject = ($eml | Select-String '^Subject:') -replace "Subject: ",""
    $sentOn  = $msg.SentOn
    $recipients = $msg.Recipients
    $sender = ( $eml | Select-String '^From:') -replace "From: ", ""
    foreach($recipient in $recipients)
    {
        $recipientList += $recipient.addressentry.address + ";"
        $recipientCount++

        if($recipient.AddressEntry.Address -like "*calvertnet.k12.md.us")
        {
            $calvertEmailCount++
            $calvertEmailList += $recipient.addressentry.address + ";"
        }
        else {
            $3rdPartyEmailCount++
            $3rdPartyEmailList += $recipient.addressentry.address + ";"
        }
    }
    if ($sender  -like "*calvertnet.k12.md.us")
    {
        $calvertEmailCount++
        $calvertEmailList += $msg.SenderEmailAddress
    }
    else {
        $3rdPartyEmailCount++
        $3rdPartyEmailList += $msg.SenderEmailAddress
    }

    $email | add-member -type NoteProperty -name sender -value $sender
    $email | add-member -type NoteProperty -name subject -value $subject
    $email | add-member -type NoteProperty -name SentOn -value $sentOn
    $email | add-member -type NoteProperty -name recipients -value $recipientList
    $email | add-member -type NoteProperty -name recipientCount -value $recipientCount
    $email | add-member -type NoteProperty -name calvertEmailCount -value $calvertEmailCount
    $email | add-member -type NoteProperty -name calvertEmailList -value $calvertEmailList
    $email | add-member -type NoteProperty -name 3rdPartyEmailCount -value $3rdPartyEmailCount
    $email | add-member -type NoteProperty -name 3rdPartyEmailList -value $3rdPartyEmailList

    $emails += $email


}

