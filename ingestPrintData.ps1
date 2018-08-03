$prints = import-csv '.\Communications PIA data breakdown 07302018.csv'

$emails = @()
$emailAddresses = ""

foreach($print in $prints)
{
    $recipientList = ""
    $calvertEmailList = ""
    $otherEmailList = ""
    $recipientCount = 0
    $calvertEmailCount = 0
    $otherEmailCount = 0
    $dupeRecipientAddressCount = 0

    $members = $print | gm -MemberType NoteProperty

    $email = new-object System.Object

    foreach($member in $members)
    {
        if ( $member.name -eq "Date")
        {
            $sentOn  = $print."$($member.name)"
        }
        elseif ($member.name -eq "Message")
        {
            $subject = $print."$($member.name)"
        }
        else {
            if ( $print."$($member.name)" -eq 0 )
            {
                $sender = $member.name
                if ( $emailAddresses -notmatch $sender )
                {
                    $emailAddresses+=$sender + ";"
                }
                if($member.name -like "*calvertnet.k12.md.us*")
                {
                    $calvertEmailCount++
                    $calvertEmailList += $member.name + ";"
                }
                else {
                    $otherEmailCount++
                    $otherEmailList += $member.name + ";"
                }
            }
            elseif ( $print."$($member.name)" -eq 1 )
            {
                $recipientList += $member.name + ";"
                if ( $emailAddresses -notmatch $member.name )
                {
                    $emailAddresses+=$member.name + ";"
                }                
                $recipientCount++

                if($member.name -like "*calvertnet.k12.md.us*")
                {
                    $calvertEmailCount++
                    $calvertEmailList += $member.name + ";"
                }
                else {
                    $otherEmailCount++
                    $otherEmailList += $member.name + ";"
                }
            }
            elseif ( $print."$($member.name)" -eq 2 )
            {
                $sender = $member.name
                if ( $emailAddresses -notmatch $sender )
                {
                    $emailAddresses+=$sender + ";"
                }
                
                $recipientList += $member.name + ";"
                if ( $emailAddresses -notmatch $member.name )
                {
                    $emailAddresses+=$member.name + ";"
                }                
                $recipientCount++

                if($member.name -like "*calvertnet.k12.md.us*")
                {
                    $calvertEmailCount++
                    $calvertEmailList += $member.name + ";"
                }
                else {
                    $otherEmailCount++
                    $otherEmailList += $member.name + ";"
                }

                

               
            }

            
        }
    }    
    

    $email | add-member -type NoteProperty -name Printed -value "Printed"
    $email | add-member -type NoteProperty -name size -value ""
    $email | add-member -type NoteProperty -name sender -value $sender
    $email | add-member -type NoteProperty -name subject -value $subject
    $email | add-member -type NoteProperty -name SentOn -value $sentOn    
    $email | add-member -type NoteProperty -name recipientCount -value $recipientCount
    $email | add-member -type NoteProperty -name calvertEmailCount -value $calvertEmailCount
    $email | add-member -type NoteProperty -name otherEmailCount -value $otherEmailCount
    $email | add-member -type NoteProperty -name dupeRecipientAddressCount -value ""
    $email | add-member -type NoteProperty -name recipients -value $recipientList
    $email | add-member -type NoteProperty -name calvertEmailList -value $calvertEmailList    
    $email | add-member -type NoteProperty -name otherEmailList -value $otherEmailList

    $emails += $email

}
