$rightsreport = Generate-ReportHeader "rightsreport.png" "$l_perm_header"

$selfSid = [System.Security.Principal.SecurityIdentifier]"S-1-5-10"

$cells=@("$l_perm_mbx","$l_perm_database","$l_perm_user","$l_perm_permission")
$rightsreport += Generate-HTMLTable "$l_perm_header2" $cells

$allmbx = get-mailbox -resultsize unlimited 
foreach ($mailbox in $allmbx)
	{
		$mbxname = $mailbox.displayname
		$mbxdb = $mailbox.Database
		$rights = Get-MailboxPermission $mailbox | where {$_.IsInherited -match "False" -and $_.user -notmatch "Selbst" -and $_.user -notmatch "Self" -and $_.Deny -match "False"}
		if ($rights)
			{
				foreach ($right in $rights)
					{
						$username = $right.user.RawIdentity
						$accessright = "$l_perm_fuccaccess"
						
						$cells=@("$mbxname","$mbxdb","$username","$accessright")
						$rightsreport += New-HTMLTableLine $cells
					}
			}
		$sendob = $mailbox.GrantSendOnBehalfTo
		if ($sendob)
			{
				foreach ($right in $sendob)
					{
						$username = $right.name
						$accessright = "$l_perm_sendonbehalf"
						
						$cells=@("$mbxname","$mbxdb","$username","$accessright")
						$rightsreport += New-HTMLTableLine $cells
					}
			}
		$sendas = Get-ADPermission $mailbox.DistinguishedName | where {($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and ($_.User.SecurityIdentifier -ne $selfSid)}
		if ($sendas)
			{
				foreach ($right in $sendas)
					{
						$username = $right.user.RawIdentity
						$accessright = "$l_perm_sendas"
						
						$cells=@("$mbxname","$mbxdb","$username","$accessright")
						$rightsreport += New-HTMLTableLine $cells
					}
			}
	
	}

$rightsreport += End-HTMLTable

$rightsreport | set-content "$tmpdir\rightsreport.html"
$rightsreport | add-content "$tmpdir\report.html"
