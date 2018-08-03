Add-Type -Path .\itextsharp.dll
$folder = "C:\temp\piaemailscan"
$files = Get-ChildItem -Path $folder -Recurse -Include "*.pdf"
$fileCount = $files.Count
$outputFolder = "c:\temp\export"
$i = 1

if ( ! ( Test-Path $outputFolder ) )
{
  mkdir $outputFolder
}

$emails = @()

foreach ($file in $files)
{
  # Parse PDF
  [string[]]$Text = $null
  $from = $null
  $sent = $null
  $to = $null
  $cc = $null
  $bcc = $null
  $subject = $null
  $attachments = $null


  #$file = get-item 'C:\temp\piaemailscan\PIA Email Scan 2018\email-1 (59).pdf'
  $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file.FullName 

  for ($page = 1; $page -le $reader.NumberOfPages; $page++)
  {
      $strategy = new-object  'iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy'            
      $currentText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page, $strategy);
      [string[]]$Text += [system.text.Encoding]::UTF8.GetString([System.Text.ASCIIEncoding]::Convert( [system.text.encoding]::default  , [system.text.encoding]::UTF8, [system.text.Encoding]::Default.GetBytes($currentText)));      
  }
  
  # Get Thread and Remove Prior Messages
  $msg = $text[0]

  $file.FullName

  $fromStart = $msg.IndexOf("From:")
  if ($fromStart -ne $msg.LastIndexOf("From:"))
  {
    $msg = $msg.Substring(0, $msg.IndexOf("From:",$fromStart+5)-1)
  }

  # Indexes
  $fromStart = $msg.IndexOf("From:")
  $sentStart = $msg.IndexOf("Sent:")
  $toStart = $msg.IndexOf("To:")
  $ccStart = $msg.IndexOf("Cc:")
  $bccStart = $msg.IndexOf("Bcc:")
  $subjectStart = $msg.IndexOf("Subject:")
  $attachmentsStart = $msg.IndexOf("Attachments:")

  # Calculate Remaining Indexes
  $fromEnd = $sentStart - 1
  $from = $msg.Substring($fromStart,$fromEnd - $fromStart)
  $from = $from.Replace("From:","").Replace("`n","").Trim()

    if ( $toStart -gt 0 )
    {
      $sentEnd = $toStart - 1
    }
    elseif ( $ccStart -gt 0 )
    {
      $sentEnd = $ccStart - 1
    }
    elseif ( $bccStart -gt 0)
    {
      $sentEnd = $bccStart - 1
    }
    elseif ( $subjectStart -gt 0 )
    {
      $sentEnd = $subjectStart - 1
    }
    elseif ( $attachmentsStart -gt 0 )
    {
      $sentEnd = $attachmentsStart - 1
    }
    else {
      $sentEnd = $msg.IndexOf("`n",$sentStart)
    }
    $sent = $msg.Substring($sentStart,$sentEnd - $sentStart)
    $sent = $sent.Replace("Sent:","").Replace("`n","").Trim()
    $sent = $sent.replace("2016"," 2016 ")
    $sent = $sent.replace("2017"," 2017 ")
    $sent = $sent.replace("2018"," 2018 ")
    $sent = $sent.replace("'","")
  

  if ( $toStart -gt 0 )
  {
    if ( $ccStart -gt 0 )
    {
      $toEnd = $ccStart - 1
    }
    elseif ( $bccStart -gt 0)
    {
      $toEnd = $bccStart - 1
    }
    elseif ( $subjectStart -gt 0 )
    {
      $toEnd = $subjectStart - 1
    }
    elseif ( $attachmentsStart -gt 0 )
    {
      $toEnd = $attachmentsStart - 1
    }
    else {
      $toEnd = $msg.IndexOf("`n",$toStart)
    }
    $to = $msg.Substring($toStart,$toEnd - $toStart)
    $to = $to.Replace("To:","").Replace("`n","").Trim()
  }

  if ( $ccStart -gt 0 )
  {
    if ( $bccStart -gt 0 )
    {
      $ccEnd = $bccStart - 1
    }    
    elseif ( $subjectStart -gt 0 )
    {
      $ccEnd = $subjectStart - 1
    }
    elseif ( $attachmentsStart -gt 0 )
    {
      $ccEnd = $attachmentsStart - 1
    }
    else {
      $ccEnd = $msg.IndexOf("`n",$ccStart)
    }
    $cc = $msg.Substring($ccStart,$ccEnd - $ccStart)
    $cc = $cc.Replace("Cc:","").Replace("`n","").Trim()
  }
  
  if ( $bccStart -gt 0 )
  {
    if ( $subjectStart -gt 0 )
    {
      $bccEnd = $subjectStart - 1
    }
    elseif ( $attachmentsStart -gt 0 )
    {
      $bccEnd = $attachmentsStart - 1
    }
    else {
      $bccEnd = $msg.IndexOf("`n",$bccStart)
    }
    $bcc = $msg.Substring($bccStart,$bccEnd - $bccStart)
    $bcc = $bcc.Replace("Bcc:","").Replace("`n","").Trim()
  }

  if ( $subjectStart -gt 0 )
  {
    if ( $attachmentsStart -gt 0 )
    {
      $subjectEnd = $attachmentsStart - 1
    }
    else {
      $subjectEnd = $msg.IndexOf("`n",$subjectStart)
    }
    $subject = $msg.Substring($subjectStart,$subjectEnd - $subjectStart)
    $subject = $subject.Replace("Subject:","").Replace("`n","").Trim()
  }

  if ( $attachmentsStart -gt 0 )
  {
    $attachmentsEnd = $msg.IndexOf("`n",$attachmentsStart)
    if ( $attachmentsEnd -lt 0 )
    {
      $attachmentsEnd = $msg.Length
    }
    $attachments = $msg.Substring($attachmentsStart, $attachmentsEnd - $attachmentsStart)
    $attachments = $attachments.Replace("Attachments:","").Replace("`n","").Trim()
  }

  if($sent)
  {
    $sentDate = ([datetime]$sent).Date
  
  $filename = $i.ToString() + "-" + (get-date $sentDate -Format "yyyy-MM-dd") + "_" + $from.Substring(0,10) + ".pdf"
  
  }
  else {
    $filename = $i.ToString() + "-" + $file.Name
  }
  
  copy $file.FullName $outputFolder\$filename

  $email = new-object System.Object
  $email | add-member -type NoteProperty -name File -value ("file:\\$outputfolder\$filename")
  $email | add-member -type NoteProperty -name From -value $from
  $email | add-member -type NoteProperty -name Sent -value $sent
  $email | add-member -type NoteProperty -name To -value $to
  $email | add-member -type NoteProperty -name Cc -value $cc
  $email | add-member -type NoteProperty -name Bcc -value $bcc
  $email | add-member -type NoteProperty -name Subject -value $subject
  $email | add-member -type NoteProperty -name Attachments -value $attachments

  $emails+= $email

  $Reader.Close();
  
  $i++
}
$emails | export-csv $outputFolder\emails.csv -NoTypeInformation

