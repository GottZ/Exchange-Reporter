# Tage die eine Verteilerliste nicht benutzt wurde
#---------------------------------------------------
$unuseddays = 14
#---------------------------------------------------

$dgreport = Generate-ReportHeader "dgreport.png" "$l_dg_header"

$cells=@("$l_dg_name","$l_dg_email","$l_dg_member")
$dgreport += Generate-HTMLTable "$l_dg_t1header $unuseddays $l_dg_t1header2" $cells

$end = get-date
$dgstart = $end.addDays(-$unuseddays)

# Version-Konsolidierung: Exchange 2010 verwendet Get-TransportServer, 2013+ Get-TransportService
if ($emsversion -match "2010")
	{ $trackingSvc = Get-TransportServer }
else
	{ $trackingSvc = Get-TransportService }

# Tracking Log einmalig laden (nur Expand-Events für DL-Nutzungscheck)
$trackingLogs = $trackingSvc | Get-MessageTrackingLog -EventId Expand -ResultSize Unlimited -start $dgstart -end $end

# Alle DGs einmalig laden inkl. DisplayName
# Verhindert O(n) get-distributiongroup-Einzelaufrufe in der Ausgabe-Schleife
$allDGs = Get-DistributionGroup -ResultSize Unlimited | Select-Object PrimarySMTPAddress, DisplayName

# Genutzte SMTP-Adressen in ein case-insensitives HashSet aufbauen (O(1)-Lookup)
# Ersetzt Sort-Object + Group-Object + Compare-Object (O(n log n) + Merge-Join)
$usedAddresses = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
if ($trackingLogs)
	{
		foreach ($entry in $trackingLogs)
			{
				if ($entry.RelatedRecipientAddress) { $null = $usedAddresses.Add([string]$entry.RelatedRecipientAddress) }
			}
	}

# Ungenutzte DLs: DGs deren SMTP-Adresse im Beobachtungszeitraum nicht im Tracking Log erscheint
$unuseddls = $allDGs | where { -not $usedAddresses.Contains([string]$_.PrimarySMTPAddress) }

if ($unuseddls)
	{
		foreach ($unuseddl in $unuseddls)
			{
				[string]$smtpaddress = $unuseddl.primarysmtpaddress
				$dgname = $unuseddl.displayname
				# -ResultSize 1: nur prüfen ob mindestens ein Mitglied vorhanden ist
				$members = Get-DistributionGroupMember $smtpaddress -ResultSize 1 -ea 0 -wa 0
				if ($members)
					{
						$hasmembers = "$l_dg_memberyes"
					}
				else
					{
						$hasmembers = "$l_dg_memberno"
					}
				$cells=@("$dgname","$smtpaddress","$hasmembers")
				$dgreport += New-HTMLTableLine $cells
			}
	}
else
	{
		$cells=@("$l_dg_nounuseddg")
		$dgreport += New-HTMLTableLine $cells
	}

$dgreport += End-HTMLTable

$dgreport| set-content "$tmpdir\dgreport.html"
$dgreport| add-content "$tmpdir\report.html"
