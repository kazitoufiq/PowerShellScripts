#To access outlook Inbox and get email attachment downloaded for specific attachmnet file name pattern


$olFolderInbox =6
$outlook = new-object -com outlook.application; 
$ns = $outlook.GetNameSpace("MAPI"); 
$inbox = $ns.GetDefaultFolder($olFolderInbox) 
$messages = $inbox.items 
write-host $messages.count 
$messcount = $messages.count 

$countprocessed = 0 

foreach($message in $messages){ 

$msubject = $message.subject 
 
 
$filepath = "J:\Asset Management\Asset Performance and Investment\11_Analytics\Lifts & Escalators - PDF file collection" 

$message.attachments|foreach { 
    Write-Host $_.filename 
    $attr = $_.filename 
    
    $a = $_.filename 
    If ($a.Contains("Lifts & Escalators")) { 
    $_.saveasfile((Join-Path $filepath $a)) 
                             } 
  } 


$countprocessed = $countprocessed + 1 
} 
