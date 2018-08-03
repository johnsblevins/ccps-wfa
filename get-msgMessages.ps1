$dir = 'C:\Users\jsble\OneDrive\ccps\Email PIA\Final - Copy' 

cd $dir

$files = get-childitem -Recurse -Filter "*msg"

$emails = @()
$emailAddresses = ""

foreach ($file in $files)
{
    $recipientList = ""
    $calvertEmailList = ""
    $otherEmailList = ""
    $recipientCount = 0
    $calvertEmailCount = 0
    $otherEmailCount = 0
    $dupeRecipientAddressCount = 0

    $email = new-object System.Object
    $email | add-member -type NoteProperty -name Folder -value $file.DirectoryName
    $email | add-member -type NoteProperty -name File -value $file.Name
    $email | add-member -type NoteProperty -name FullPath -value $file.FullName
    $email | add-member -type NoteProperty -name Size -value $file.length
    $outlook = New-Object -comobject outlook.application
    $msg = $outlook.CreateItemFromTemplate($file.FullName)
    $subject = $msg.Subject
    $sentOn  = $msg.SentOn
    $recipients = $msg.Recipients
    $sender = $msg.SenderEmailAddress.ToLower()
    if ( $emailAddresses -notmatch $sender )
    {
        $emailAddresses+=$sender + ";"
    }
    foreach($recipient in $recipients)
    {
        if ( $recipientList -match $recipient.addressentry.address )
        {
            $dupeRecipientAddressCount++
        }
        else
        {
            if ( $emailAddresses -notmatch $recipient.addressentry.address )
            {
                $emailAddresses+=$recipient.addressentry.address + ";"
            }

            $recipientList += $recipient.addressentry.address + ";"
            $recipientCount++

            if($recipient.AddressEntry.Address -like "*calvertnet.k12.md.us")
            {
                $calvertEmailCount++
                $calvertEmailList += $recipient.addressentry.address + ";"
            }
            else {
                $otherEmailCount++
                $otherEmailList += $recipient.addressentry.address + ";"
            }
        }
    }
    if ($sender  -like "*calvertnet.k12.md.us")
    {
        $calvertEmailCount++
        $calvertEmailList += $msg.SenderEmailAddress
    }
    else {
        $otherEmailCount++
        $otherEmailList += $msg.SenderEmailAddress
    }

    $email | add-member -type NoteProperty -name sender -value $sender
    $email | add-member -type NoteProperty -name subject -value $subject
    $email | add-member -type NoteProperty -name SentOn -value $sentOn
    $email | add-member -type NoteProperty -name recipients -value $recipientList
    $email | add-member -type NoteProperty -name recipientCount -value $recipientCount
    $email | add-member -type NoteProperty -name calvertEmailCount -value $calvertEmailCount
    $email | add-member -type NoteProperty -name calvertEmailList -value $calvertEmailList
    $email | add-member -type NoteProperty -name otherEmailCount -value $otherEmailCount
    $email | add-member -type NoteProperty -name otherEmailList -value $otherEmailList
    $email | add-member -type NoteProperty -name dupeRecipientAddressCount -value $dupeRecipientAddressCount

    $emails += $email

}

$groupedMessages = $emails | group-object size, sender, SentOn

$uniqueMessages = @()

foreach($groupedMessage in $groupedMessages)
{

    $uniqueMessage = new-object System.Object
    $uniqueMessage | add-member -type NoteProperty -name Size -value $groupedMessage.Group[0].Size
    $uniqueMessage | add-member -type NoteProperty -name Sender -value $groupedMessage.Group[0].Sender
    $uniqueMessage | add-member -type NoteProperty -name Subject -value $groupedMessage.Group[0].Subject
    $uniqueMessage | add-member -type NoteProperty -name SentOn -value $groupedMessage.Group[0].SentOn
    $uniqueMessage | add-member -type NoteProperty -name recipientCount -value $groupedMessage.Group[0].recipientCount
    $uniqueMessage | add-member -type NoteProperty -name calvertEmailCount -value $groupedMessage.Group[0].calvertEmailCount
    $uniqueMessage | add-member -type NoteProperty -name otherEmailCount -value $groupedMessage.Group[0].otherEmailCount
    $uniqueMessage | add-member -type NoteProperty -name otherEmailList -value $groupedMessage.Group[0].otherEmailList
    $uniqueMessage | add-member -type NoteProperty -name dupeRecipientAddressCount -value $groupedMessage.Group[0].dupeRecipientAddressCount
    $uniqueMessage | add-member -type NoteProperty -name Recipients -value $groupedMessage.Group[0].Recipients
    $uniqueMessage | add-member -type NoteProperty -name calvertEmailList -value $groupedMessage.Group[0].calvertEmailList
    
    $uniqueMessages += $uniqueMessage
}