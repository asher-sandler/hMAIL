$now=$(Get-Date)
$logfile = ($now.GetDateTimeFormats()[24]).Replace(":","-")
$logFile="c:\admindir\scripts\log\hmail_clear_ECO_MSG.log"
Start-Transcript $logFile

$domainName ="mail.kaluga.sds.cbr.ru"
$accountName="eco@mail.kaluga.sds.cbr.ru"
$folderName="INBOX"
;$oldDate=$(Get-Date).AddMonths(-2)
$oldDate=$(Get-Date).AddDays(-35)
$oldMonth = $oldDate.Month
$oldYear  = $oldDate.Year







$user = "eco@mail.kaluga.sds.cbr.ru"
$pwd  = "!QAZ1qaz"

$hmail = New-Object -ComObject hMailServer.Application

$hmail.Authenticate($user,$pwd)

$oDomain=$hmail.Domains.ItemByName($domainName)

$oAccount = $oDomain.Accounts.ItemByAddress($accountName)

$folders=$oAccount.IMAPFolders
$iFoldCount = $folders.Count
<#
for ($i=0;$i -lt $iFoldCount;$i++ )
{
    $folder=$folders.Item($i)
    
    #$folder.Name
    #$folder.Messages.Count
    $sf = $folder.SubFolders.Count
    if ($sf -gt 0)
    {
        $folder.Name
        $subfolder=$folder.SubFolders
    }

}

#>

# $oMessages=$folders.Item(3).Messages
$inbox=$folders.ItemByName($folderName)
#$outbox=$folders.Item(3)
#$deletedFolder = $folders.ItemByName("Удаленные")
$oMessages=$inbox.Messages
$itemsToDelete = 0

Write-Output $("Deleting item older than "+$oldDate.ToLongDateString())

for($j=1;$j -lt 2000;$j++)
{
    $iCount = $oMessages.Count

    #$iCount

    $msgID=-1
    #$j
    $wasDeleted = $false
    for($iMessage=0;$iMessage -lt $iCount;$iMessage++)
    {
      $oMessage=$oMessages.Item($iMessage)
            $mDate = $oMessage.InternalDate
               
            #if (($mDate.Month -le $oldMonth) -and (($mDate.Year -le $oldYear)))
	if ($mDate -le $oldDate)
            {
                $itemsToDelete++
                $msgID=$oMessage.ID

                $oMessage.Date
                $oMessage.Subject
                
                #$msgID
                #read-host
                
                $oMessages.DeleteByDBID($msgID)
                $wasDeleted = $True
                break;
                
            }
      
     }
     if (!$wasDeleted)       
     {
            break;
     }
}

#$oMessage | gm


Write-Output $("Count of deleted items was:"+$itemsToDelete)

Stop-Transcript


