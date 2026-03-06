$mailreport = Generate-ReportHeader "mailreport.png" "$l_mail_header"

$cells=@("$l_mail_sendcount","$l_mail_reccount","$l_mail_volsend","$l_mail_volrec")
$mailreport += Generate-HTMLTable "$l_mail_header2 $ReportInterval $l_mail_days" $cells

$mailexclude = ($excludelist | where {$_.setting -match "mailreport"}).value
if ($mailexclude)
	{
		[array]$mailexclude = $mailexclude.split(",")
	}

if ($emsversion -match "2010")
	{
		$exchangeInstallPath = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup -ea 0).MsiInstallPath
		$transportservers    = Get-TransportServer
	}
else
	{
		$exchangeInstallPath = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -ea 0).MsiInstallPath
		$transportservers    = Get-TransportService
	}

$trackingLogJob = {
	param($srvname, $Start, $End, $exchangeInstallPath)
	$VerbosePreference = 'SilentlyContinue'
	. ($exchangeInstallPath + "bin\RemoteExchange.ps1") *>&1 | Out-Null
	Connect-ExchangeServer -auto *>&1 | Out-Null
	$send = Get-MessageTrackingLog -Server $srvname -Start $Start -End $End `
		-EventId Send -ResultSize Unlimited -ea 0 |
		where {$_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.source -match "SMTP"} |
		select sender,Recipients,timestamp,totalbytes,clienthostname,@{N='EventType';E={'Send'}}
	$receive = Get-MessageTrackingLog -Server $srvname -Start $Start -End $End `
		-EventId Receive -ResultSize Unlimited -ea 0 |
		where {$_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.source -match "SMTP"} |
		select sender,Recipients,timestamp,totalbytes,serverhostname,@{N='EventType';E={'Receive'}}
	@($send) + @($receive)
}

$jobServerMap = @{}
$serverJobs   = @()
foreach ($server in $transportservers)
	{
		$job = Start-Job -ScriptBlock $trackingLogJob -ArgumentList $server.Name, $Start, $End, $exchangeInstallPath
		$jobServerMap[$job.Id] = $server.Name
		$serverJobs += $job
	}

$null = $serverJobs | Wait-Job -Timeout 1800

$allMails      = [System.Collections.Generic.List[PSObject]]::new()
$failedServers = [System.Collections.Generic.List[string]]::new()

foreach ($job in $serverJobs)
	{
		if ($job.State -eq 'Completed')
			{
				foreach ($r in @(Receive-Job -Job $job)) { if ($r) { $allMails.Add($r) } }
			}
		else
			{
				$failedServers.Add($jobServerMap[$job.Id])
			}
		Remove-Job -Job $job -Force
	}

foreach ($srvname in $failedServers)
	{
		$fallbackSend = Get-MessageTrackingLog -Server $srvname -Start $Start -End $End `
			-EventId Send -ResultSize Unlimited -ea 0 |
			where {$_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.source -match "SMTP"} |
			select sender,Recipients,timestamp,totalbytes,clienthostname,@{N='EventType';E={'Send'}}
		$fallbackReceive = Get-MessageTrackingLog -Server $srvname -Start $Start -End $End `
			-EventId Receive -ResultSize Unlimited -ea 0 |
			where {$_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.source -match "SMTP"} |
			select sender,Recipients,timestamp,totalbytes,serverhostname,@{N='EventType';E={'Receive'}}
		foreach ($r in @($fallbackSend) + @($fallbackReceive)) { if ($r) { $allMails.Add($r) } }
	}

$SendMails     = @($allMails | where { $_.EventType -eq 'Send' })
$ReceivedMails = @($allMails | where { $_.EventType -eq 'Receive' })

if ($mailexclude)
	{
		foreach ($entry in $mailexclude) {$SendMails = $SendMails | where {$_.sender -notmatch $entry -and $_.recipients -notmatch $entry}}
		foreach ($entry in $mailexclude) {$ReceivedMails = $ReceivedMails | where {$_.sender -notmatch $entry -and $_.recipients -notmatch $entry}}
	}

#Total

$totalsendmail = $sendmails | measure-object Totalbytes -sum
$totalreceivedmail = $receivedmails  | measure-object Totalbytes -sum

$totalsendvol = $totalsendmail.sum
$totalreceivedvol = $totalreceivedmail.sum
$totalsendvol = $totalsendvol / 1024 /1024
$totalreceivedvol = $totalreceivedvol / 1024 /1024
$totalsendvol = [System.Math]::Round($totalsendvol , 2)
$totalreceivedvol  = [System.Math]::Round($totalreceivedvol , 2)

$totalsendcount = $totalsendmail.count
$totalreceivedcount = $totalreceivedmail.count

$totalmail = @{$l_mail_send=$totalsendcount}
$totalmail +=@{$l_mail_received=$totalreceivedcount}

new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_count" $totalmail "$tmpdir\totalmailcount.png"

$totalmail = @{$l_mail_send=$totalsendvol}
$totalmail +=@{$l_mail_received=$totalreceivedvol}

new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_size" $totalmail "$tmpdir\totalmailvol.png"

$cells=@("$totalsendcount","$totalreceivedcount","$totalsendvol","$totalreceivedvol")
$mailreport += New-HTMLTableLine $cells
$mailreport += End-HTMLTable
$mailreport += Include-HTMLInlinePictures "$tmpdir\totalmail*.png"

#Je Server
if ($transportservers.count -gt 1)
{
$cells=@("$l_mail_servername","$l_mail_overallcount","$l_mail_overallvolume","$l_mail_sendcount","$l_mail_reccount","$l_mail_volsend","$l_mail_volrec")
$mailreport += Generate-HTMLTable "$l_mail_header2 $ReportInterval $l_mail_days $l_mail_perserver" $cells

$perserverstats  = [System.Collections.Generic.List[PSObject]]::new()
foreach ($transportserver in $transportservers)
	{
		$tpsname = $transportserver.name
		$tpssend = $sendmails | where {$_.Clienthostname -match "$tpsname"} | measure-object Totalbytes -sum
		$tpsreceive = $ReceivedMails | where {$_.serverhostname -match "$tpsname"} | measure-object Totalbytes -sum
		$tpssendcount = $tpssend.count
		$tpsreceivecount = $tpsreceive.count
		
		$tpssendvol = $tpssend.sum
		$tpssendvol = $tpssendvol / 1024 / 1024
		$tpssendvol = [System.Math]::Round($tpssendvol , 2)
		$tpsreceivevol = $tpsreceive.sum
		$tpsreceivevol = $tpsreceivevol / 1024 /1024
		$tpsreceivevol = [System.Math]::Round($tpsreceivevol , 2)
		

		$tpstotalvol = $tpsreceivevol + $tpssendvol
		$tpstotalcount = $tpsreceivecount + $tpssendcount
		
		$cells=@("$tpsname","$tpstotalcount","$tpstotalvol","$tpssendcount","$tpsreceivecount","$tpssendvol","$tpsreceivevol")
		$mailreport += New-HTMLTableLine $cells
		
		$perserverstats.Add((new-object PSObject -property @{Name="$tpsname";TotalCount=$tpstotalcount;SendCount=$tpssendcount;ReceiveCount=$tpsreceivecount;ToltalVolume=$tpstotalvol;SendVolume=$tpssendvol;Receivevolume=$tpsreceivevol}))
	}
$mailreport += End-HTMLTable

foreach ($tpserver in $perserverstats)
	{
		$tpsname = $tpserver.name
		$tpstotalvol = $tpserver.ToltalVolume
		$tpstotalcount = $tpserver.TotalCount		
		$tpssendvol = $tpserver.SendVolume
		$tpsreceivedvol = $tpserver.Receivevolume
		$tpssendcount = $tpserver.SendCount
		$tpsreceivedcount = $tpserver.ReceiveCount
		
		$tpsvoldata += [ordered]@{$tpsname=$tpstotalvol}
		$tpscountdata += [ordered]@{$tpsname=$tpstotalcount}

		$tpsrscountdata += [ordered]@{"$tpsname $l_mail_send"=$tpssendcount}
		#$tpsrscountdata += @{"$tpsname $l_mail_received"=$tpsreceivedcount}
		
		$tpsrsvoldata += [ordered]@{"$tpsname $l_mail_send"=$tpssendvol}
		#$tpsrsvoldata += @{"$tpsname $l_mail_received"=$tpsreceivedvol}
	
	}
	
foreach ($tpserver in $perserverstats)
	{
		$tpsname = $tpserver.name
		$tpsreceivedvol = $tpserver.Receivevolume
		$tpsreceivedcount = $tpserver.ReceiveCount
		


		#$tpsrscountdata += @{"$tpsname $l_mail_send"=$tpssendcount}
		$tpsrscountdata += [ordered]@{"$tpsname $l_mail_received"=$tpsreceivedcount}
		
		#$tpsrsvoldata += @{"$tpsname $l_mail_send"=$tpssendvol}
		$tpsrsvoldata += [ordered]@{"$tpsname $l_mail_received"=$tpsreceivedvol}
	}
		
new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_size $l_mail_overall" $tpsvoldata "$tmpdir\pertpsvol.png"
new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_count $l_mail_overall" $tpscountdata "$tmpdir\pertpscount.png"
new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_size" $tpsrsvoldata "$tmpdir\pertpsvolrs.png"
new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_coun" $tpsrscountdata "$tmpdir\pertpscountrs.png"

$mailreport += Include-HTMLInlinePictures "$tmpdir\pertps*.png"
}
#days - pre-group by date key once (O(n)) instead of filtering per day (O(n * days))
$sendByDay    = @{}
$receiveByDay = @{}

foreach ($mail in $SendMails)
	{
		$key = $mail.timestamp.ToString("dd.MM.yy")
		if (-not $sendByDay.ContainsKey($key)) { $sendByDay[$key] = @{Count=0;Bytes=[long]0} }
		$sendByDay[$key].Count++
		$sendByDay[$key].Bytes += [long]$mail.totalbytes
	}

foreach ($mail in $ReceivedMails)
	{
		$key = $mail.timestamp.ToString("dd.MM.yy")
		if (-not $receiveByDay.ContainsKey($key)) { $receiveByDay[$key] = @{Count=0;Bytes=[long]0} }
		$receiveByDay[$key].Count++
		$receiveByDay[$key].Bytes += [long]$mail.totalbytes
	}

$cells=@("$l_mail_date","$l_mail_sendcount","$l_mail_reccount","$l_mail_volsend","$l_mail_volrec")
$mailreport += Generate-HTMLTable "$l_Mail_overviewperday" $cells

$daycounter = 1
do
 {
 $daystart = (Get-Date -Hour 00 -Minute 00 -Second 00).AddDays(-$daycounter)
 $day = $daystart | get-date -Format "dd.MM.yy"

 $daytotalsendcount     = if ($sendByDay.ContainsKey($day))    { $sendByDay[$day].Count }    else { 0 }
 $daytotalsendvol       = if ($sendByDay.ContainsKey($day))    { [System.Math]::Round($sendByDay[$day].Bytes / 1024 / 1024, 2) } else { 0 }
 $daytotalreceivedcount = if ($receiveByDay.ContainsKey($day)) { $receiveByDay[$day].Count } else { 0 }
 $daytotalreceivedvol   = if ($receiveByDay.ContainsKey($day)) { [System.Math]::Round($receiveByDay[$day].Bytes / 1024 / 1024, 2) } else { 0 }

 $daystotalmailvol   += [ordered]@{$day=$daytotalreceivedvol}
 $daystotalmailcount += [ordered]@{$day=$daytotalreceivedcount}

 $cells=@("$day","$daytotalsendcount","$daytotalreceivedcount","$daytotalsendvol","$daytotalreceivedvol")
 $mailreport += New-HTMLTableLine $cells

 $daycounter++
 }
 while ($daycounter -le $reportinterval)

 new-cylinderchart 500 400 "$l_mail_daycount" Mails "$l_mail_count" $daystotalmailcount "$tmpdir\dailymailcount.png"
 new-cylinderchart 500 400 "$l_mail_daysize" Mails "$l_mail_size" $daystotalmailvol "$tmpdir\dailymailvol.png"

$mailreport += End-HTMLTable
$mailreport += Include-HTMLInlinePictures "$tmpdir\dailymail*.png"

$sendstat = $SendMails | select sender,totalbytes
$receivedstat = $receivedMails | select sender,totalbytes

$sendmails     = @($sendmails     | ForEach-Object { [string]$_.sender })
$ReceivedMails = @($ReceivedMails | ForEach-Object { $_.Recipients | ForEach-Object { [string]$_ } })

$topsenders = $sendmails | Group-Object -noelement | Sort-Object Count -descending | Select-Object -first $DisplayTop
$toprecipients = $ReceivedMails | Group-Object -noelement | Sort-Object Count -descending | Select-Object -first $DisplayTop

$cells=@("$l_mail_sender","$l_mail_count")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_sender ($l_mail_count)" $cells
foreach ($topsender in $topsenders)
{
 $tsname = $topsender.name
 $tscount = $topsender.count
 
 $cells=@("$tsname","$tscount")
 $mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable

$cells=@("$l_mail_recipient","$l_mail_count")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_recipient ($l_mail_count)" $cells
foreach ($toprecipient in $toprecipients)
{
 $trname = $toprecipient.name
 $trcount = $toprecipient.count
 
 $cells=@("$trname","$trcount")
 $mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable

#--------------
#Sender
$cells=@("$l_mail_sender","$l_mail_sizemb")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_sender ($l_mail_size)" $cells

$senderVolHash = @{}
foreach ($mail in $sendstat)
	{
		$key = [string]$mail.sender
		if ($key) { $senderVolHash[$key] = ([long]$senderVolHash[$key]) + ([long]$mail.totalbytes) }
	}
$toptensendersvol = $senderVolHash.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First $DisplayTop |
	Select-Object @{N='Name';E={$_.Key}},@{N='Volume';E={$_.Value}}

foreach ($topsender in $toptensendersvol)
{
 $tsname = $topsender.name
 $tsvolume= $topsender.volume
 $tsvolume = $tsvolume / 1024 /1024
 $tsvolume = [System.Math]::Round($tsvolume , 2)
 $cells=@("$tsname","$tsvolume")
 $mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable

#Recipient
$cells=@("$l_mail_recipient","$l_mail_sizemb")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_recipient ($l_mail_size)" $cells

$recipientVolHash = @{}
foreach ($mail in $receivedstat)
	{
		$key = [string]$mail.sender
		if ($key) { $recipientVolHash[$key] = ([long]$recipientVolHash[$key]) + ([long]$mail.totalbytes) }
	}
$toptenrecipientsvol = $recipientVolHash.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First $DisplayTop |
	Select-Object @{N='Name';E={$_.Key}},@{N='Volume';E={$_.Value}}

foreach ($toprecipient in $toptenrecipientsvol)
{
 $trname = $toprecipient.name
 $trvolume = $toprecipient.Volume
 $trvolume = $trvolume / 1024 /1024
 $trvolume = [System.Math]::Round($trvolume , 2)
 $cells=@("$trname","$trvolume")
 $mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable

#---------------------------

#Durchschnitt
try
{
$usercount = (Get-Mailbox -ResultSize Unlimited).Count

$dsend = $totalsendcount / $usercount
 $dsend = [System.Math]::Round($dsend , 2)
$dreceived = $totalreceivedcount / $usercount
 $dreceived= [System.Math]::Round($dreceived , 2)
$dsendvol = $totalsendvol / $usercount
 $dsendvol= [System.Math]::Round($dsendvol , 2)
$dreceivedvol = $totalreceivedvol / $usercount
 $dreceivedvol = [System.Math]::Round($dreceivedvol  , 2)
$dmailsizesend = $totalsendvol / $totalsendcount
 $dmailsizesend= [System.Math]::Round($dmailsizesend , 2)
$dmailsizereceived = $totalreceivedvol / $totalreceivedcount
 $dmailsizereceived= [System.Math]::Round($dmailsizereceived , 2)

$cells=@("$l_mail_average","$l_mail_value")
$mailreport += Generate-HTMLTable "$l_mail_averagevalue" $cells

 $cells=@("$l_mail_avmbxsendcount","$dsend")
 $mailreport += New-HTMLTableLine $cells
 
 $cells=@("$l_mail_avmbxreccount","$dreceived")
 $mailreport += New-HTMLTableLine $cells
 
 $cells=@("$l_mail_avmbxsendsize","$dsendvol MB")
 $mailreport += New-HTMLTableLine $cells
 
 $cells=@("$l_mail_avmbxrecsize","$dreceivedvol MB")
 $mailreport += New-HTMLTableLine $cells

 $cells=@("$l_mail_avmailsendsize","$dmailsizesend MB")
 $mailreport += New-HTMLTableLine $cells
 
 $cells=@("$l_mail_avmailrecsize","$dmailsizereceived MB")
 $mailreport += New-HTMLTableLine $cells
 
$mailreport += End-HTMLTable
}
catch
{
}
 
$mailreport | set-content "$tmpdir\mailreport.html"
$mailreport | add-content "$tmpdir\report.html"