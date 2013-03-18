# Use the following area to define the colours of the report
$Colour1 = "CC0000" # Main Title - currently red
$Colour2 = "7BA7C7" # Secondary Title - currently blue
$enableExport = "true"


function Send-SMTPmail($to, $from, $subject, $smtpserver, $body) {
    $mailer = new-object Net.Mail.SMTPclient($smtpserver)
	$msg = new-object Net.Mail.MailMessage($from,$to,$subject,$body)
	$msg.IsBodyHTML = $true
	$mailer.send($msg)
}


Function Get-CustomHTML ($Header){
$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($Header)</title>
		<META http-equiv=Content-Type content='text/html; charset=windows-1252'>

		<style type="text/css">

		TABLE 		{
						TABLE-LAYOUT: fixed; 
						FONT-SIZE: 100%; 
						WIDTH: 100%
					}
		*
					{
						margin:0
					}

		.dspcont 	{
	
						BORDER-RIGHT: #bbbbbb 1px solid;
						BORDER-TOP: #bbbbbb 1px solid;
						PADDING-LEFT: 0px;
						FONT-SIZE: 8pt;
						MARGIN-BOTTOM: -1px;
						PADDING-BOTTOM: 5px;
						MARGIN-LEFT: 0px;
						BORDER-LEFT: #bbbbbb 1px solid;
						WIDTH: 95%;
						COLOR: #000000;
						MARGIN-RIGHT: 0px;
						PADDING-TOP: 4px;
						BORDER-BOTTOM: #bbbbbb 1px solid;
						FONT-FAMILY: Tahoma;
						POSITION: relative;
						BACKGROUND-COLOR: #f9f9f9
					}
					
		.filler 	{
						BORDER-RIGHT: medium none; 
						BORDER-TOP: medium none; 
						DISPLAY: block; 
						BACKGROUND: none transparent scroll repeat 0% 0%; 
						MARGIN-BOTTOM: -1px; 
						FONT: 100%/8px Tahoma; 
						MARGIN-LEFT: 43px; 
						BORDER-LEFT: medium none; 
						COLOR: #ffffff; 
						
						MARGIN-RIGHT: 0px; 
						PADDING-TOP: 4px; 
						BORDER-BOTTOM: medium none; 
						POSITION: relative
					}

		.pageholder	{
						margin: 0px auto;
					}
					
		.dsp
					{
						BORDER-RIGHT: #bbbbbb 1px solid;
						PADDING-RIGHT: 0px;
						BORDER-TOP: #bbbbbb 1px solid;
						DISPLAY: block;
						PADDING-LEFT: 0px;
						FONT-WEIGHT: bold;
						FONT-SIZE: 8pt;
						MARGIN-BOTTOM: -1px;
						MARGIN-LEFT: 0px;
						BORDER-LEFT: #bbbbbb 1px solid;
						COLOR: #000000;
						MARGIN-RIGHT: 0px;
						PADDING-TOP: 4px;
						BORDER-BOTTOM: #bbbbbb 1px solid;
						FONT-FAMILY: Tahoma;
						POSITION: relative;
						HEIGHT: 2.25em;
						WIDTH: 95%;
						TEXT-INDENT: 10px;
					}

		.dsphead0	{
						BACKGROUND-COLOR: #$($Colour1);
					}
					
		.dsphead1	{
						
						BACKGROUND-COLOR: #$($Colour2);
					}
					
	.dspcomments 	{
						BACKGROUND-COLOR:#FFFFE1;
						COLOR: #000000;
						FONT-STYLE: ITALIC;
						FONT-WEIGHT: normal;
						FONT-SIZE: 8pt;
					}

	td 				{
						VERTICAL-ALIGN: TOP; 
						FONT-FAMILY: Tahoma
					}
					
	th 				{
						VERTICAL-ALIGN: TOP; 
						COLOR: #$($Colour1); 
						TEXT-ALIGN: left
					}
					
	BODY 			{
						margin-left: 4pt;
						margin-right: 4pt;
						margin-top: 6pt;
					} 
	.MainTitle		{
						font-family:Arial, Helvetica, sans-serif;
						font-size:20px;
						font-weight:bolder;
					}
	.SubTitle		{
						font-family:Arial, Helvetica, sans-serif;
						font-size:14px;
						font-weight:bold;
					}
	.Created		{
						font-family:Arial, Helvetica, sans-serif;
						font-size:10px;
						font-weight:normal;
						margin-top: 20px;
						margin-bottom:5px;
					}
	.links			{	font:Arial, Helvetica, sans-serif;
						font-size:10px;
						FONT-STYLE: ITALIC;
					}
					
		</style>
	</head>
	<body>
<div class="MainTitle">$($Header)</div>
        <hr size="8" color="#$($Colour1)">
        <div class="SubTitle">vCheck for Advantage Team generated on $($ENV:Computername)</div>
	    <br/>
		<div class="Created">Report created on $(Get-Date)</div>
"@
Return $Report
}


Function Get-CustomHeader0 ($Title){
$Report = @"
		<div class="pageholder">		

		<h1 class="dsp dsphead0">$($Title)</h1>
	
    	<div class="filler"></div>
"@
Return $Report
}


Function Get-CustomHeader ($Title, $cmnt){
$Report = @"
	    <h2 class="dsp dsphead1">$($Title)</h2>
"@
If (1) {
	$Report += @"
			<div class="dsp dspcomments">$($cmnt)</div>
"@
}
$Report += @"
        <div class="dspcont">
"@
Return $Report
}


Function Get-CustomHeaderClose{

	$Report = @"
		</DIV>
		<div class="filler"></div>
"@
Return $Report
}


Function Get-CustomHeader0Close{
	$Report = @"
</DIV>
"@
Return $Report
}



Function Get-CustomHTMLClose{
	$Report = @"
</div>

</body>
</html>
"@
Return $Report
}


Function Get-HTMLTable {
	param([array]$Content)
	$HTMLTable = $Content | ConvertTo-Html
	$HTMLTable = $HTMLTable -replace '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">', ""
	$HTMLTable = $HTMLTable -replace '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN"  "http://www.w3.org/TR/html4/strict.dtd">', ""
	$HTMLTable = $HTMLTable -replace '<html xmlns="http://www.w3.org/1999/xhtml">', ""
	$HTMLTable = $HTMLTable -replace '<html>', ""
	$HTMLTable = $HTMLTable -replace '<head>', ""
	$HTMLTable = $HTMLTable -replace '<title>HTML TABLE</title>', ""
	$HTMLTable = $HTMLTable -replace '</head><body>', ""
	$HTMLTable = $HTMLTable -replace '</body></html>', ""
	$HTMLTable = $HTMLTable -replace '&lt;', "<"
	$HTMLTable = $HTMLTable -replace '&gt;', ">"
	Return $HTMLTable
}


Function Get-HTMLDetail ($Heading, $Detail){
$Report = @"
<TABLE>
	<tr>
	<th width='50%'><b>$Heading</b></font></th>
	<td width='50%'>$($Detail)</td>
	</tr>
</TABLE>
"@
Return $Report
}





#import powershell community extensions for increased functionality
Import-Module Pscx 

#import vmware cli
Add-PSSnapin "Vmware.VimAutomation.Core" -ErrorAction SilentlyContinue


#vm vsphere server
$vsphereServer = "your vsphere server ip"
$vsphereUser = "administrator"
$vspherePass = 'Password'
$Hostlist = get-content .\vmHostName.txt
$AdvDSvCenter= Get-Content .\ADV-DS-vCenter.txt
$AdvDSLM =  Get-Content .\ADV-DS-LM.txt
$Date = get-date 
$dateFormat = ( get-date ).ToString(‘yyyy-MM-dd’) 
$FileCSV = ".\AllVMInfo.csv"

#connect to the virtual server the VM's are on
$VIServer = Connect-VIServer -Server $vsphereServer -Protocol https -User $vsphereUser -Password $vspherePass

#Delte the old date before collecting the new data.
Get-ChildItem -Path ".\PMS data folder\" | Remove-Item -Force


# Collect all vm OS info.
Write-Host "Starting to collect all VM OS info..."
$vmOSReport = @()
foreach ($hostserver in $Hostlist){ `
Get-VMhost -Name $hostserver |get-vm | Sort Name -Descending | % `
{   

$vm = Get-View $_.ID 
$vms = "" | Select-Object vmname, hostname, osinfo, status,owner
$vms.vmname = $vm.Name 
$vms.hostname = $_.Host 
$vms.osinfo = ($vm.Config.GuestFullName).replace(",","") 
$vms.status = $vm.summary.runtime.powerState 
$vms.owner = ($_ | Get-Annotation -CustomAttribute Owner).value

$vmOSReport += $vms 
}

}

#Output 
if ($enableExport -eq "true") { 
$vmOSReport | Export-Csv $FileCSV -NoTypeInformation -Confirm:$false 
Write-Host "Finished to collect all VM OS info..."
}



#Collect generic host and vm info data.
Write-Host "Starting to collect generic host and vm info data..."
$DeptShow = @()
$vCenterFreeSpaceGB = 0
$vCenterTotalCapacityGB = 0
$LMFreeSpaceGB = 0
$LMTotalCapacityGB = 0
$ADVHosts = 0
$ADVVMs = 0

foreach ($ds1 in $AdvDSvCenter){
    $DSInfo = Get-Datastore -Name $ds1
    $vCenterFreeSpaceGB += [int]($DSInfo.FreeSpaceMB/1000)
    $vCenterTotalCapacityGB += [int]($DSInfo.CapacityMB/1000)
    Write-Host "Collecting vCenter Datastore..."

}

foreach ($ds2 in $AdvDSLM){
    $DSInfo = Get-Datastore -Name $ds2
    $LMFreeSpaceGB += [int]($DSInfo.FreeSpaceMB/1000)
    $LMTotalCapacityGB += [int]($DSInfo.CapacityMB/1000)
    Write-Host "Collecting LabManager Datastore..."

}

foreach ($Dept in (Get-View -ViewType ClusterComputeResource)){
    if ($Dept.Name -like "EGS_ADV*" -or $Dept.Name -like "EGS_Lagmgr*"){
    
		$ADVName = "ADV"
		$ADVHosts = $ADVHosts + ( Get-Cluster $Dept.Name|Get-VMHost |Measure-Object).Count
						
		$ADVVMs = $ADVVMs + (Get-Cluster $Dept.Name|Get-VM |Measure-Object).Count 
             
         Write-Host "Collecting Department info..."     
		#$ADVHDDSize +=  Get-VMHHDSize ($Dept.Name)
		#$ADVHDDSize = [Math]::Round($ADVHDDSize)
		#Write-Host $Dept.Name $ADVName "Number of Hosts:"$ADVHosts "Number of VMs:"$ADVVMs "Total Size of VMs HDD:"$ADVHDDSize        
    }

}


	### ADV
	$DeptShow= ""|Select deptname ,numberofhosts ,numberofvms, vcenterstorageused, vcenterstoragetotal, lmstorageused, lmstoragecapacity
	$DeptShow.deptname = $ADVName
	$DeptShow.numberofhosts = $ADVHosts
	$DeptShow.numberofvms = $ADVVMs
	$DeptShow.vcenterstorageused = $vCenterTotalCapacityGB - $vCenterFreeSpaceGB
    $DeptShow.vcenterstoragetotal = $vCenterTotalCapacityGB
    $DeptShow.lmstorageused = $LMTotalCapacityGB - $LMFreeSpaceGB
    $DeptShow.lmstoragecapacity = $LMTotalCapacityGB
	#$DeptInfo +=  $DeptShow1




# Collect Host CPU and Memory Usage performance data.
Write-Host "Collecting Host CUP/Mem usage info..."
$HostPerfInfo = @()
foreach ($ESXHost in $Hostlist ){

					$row=""|select hostname ,cpuuse,memuse
					$row.hostname=$ESXHost
					$statsCPU=Get-Stat -Entity $ESXHost -start (get-date).AddDays(-7) -stat cpu.usage.average

					$statsMEM=Get-Stat -Entity $ESXHost -start (get-date).AddDays(-7) -Stat mem.usage.average
				 	
					$statsCPUgrouped=$statsCPU|group-object -Property MetricId
					$statsMEMgrouped=$statsMEM|group-object -Property MetricId
									
					$row.cpuuse=[Math]::Round(($statsCPUgrouped|where {$_.Name -eq "cpu.usage.average"}|%{$_.Group|Measure-Object -Property Value -Average}).Average,2)
					if ($row.cpuuse -eq 0) {
						$statsCPU=Get-Stat -Entity $ESXHost -Realtime -stat cpu.usage.average
						$statsCPUgrouped=$statsCPU|group-object -Property MetricId
						$row.cpuuse=[Math]::Round(($statsCPUgrouped|where {$_.Name -eq "cpu.usage.average"}|%{$_.Group|Measure-Object -Property Value -Average}).Average,2)
						}
					$row.memuse=[Math]::Round(($statsMEMgrouped|where {$_.Name -eq "mem.usage.average"}|%{$_.Group|Measure-Object -Property Value -Average}).Average,2)
					if ($row.memuse -eq 0) {
						$statsMEM=Get-Stat -Entity $ESXHost -Realtime -stat mem.usage.average
						$statsMEMgrouped=$statsMEM|group-object -Property MetricId
						$row.memuse=[Math]::Round(($statsMEMgrouped|where {$_.Name -eq "mem.usage.average"}|%{$_.Group|Measure-Object -Property Value -Average}).Average,2)
						}					

					$row.cpuuse =$($row.cpuuse) #+"%"
					$row.memuse =$($row.memuse) #+"%"
					
					#Write-Host "CPU : $($row.CPURate) "
					#Write-Host "Memory : $($row.MemoryRate) "
                    $HostPerfInfo += $row


}




#Find out the powered off VM more than 3 months.
Write-Host "Collecting long time powered off VM info..."
$PowerOffVMs = @()


			foreach ($managedHost in $Hostlist ){
				$managedHostName=(Get-VMHost $managedHost|Get-View).Name
				if ((Get-VMHost $managedHost|Get-VM|where {$_.PowerState -eq "poweredOff"}|Measure-Object).count -gt 0){
					$VMs = Get-VMHost $managedHost|Get-VM|where {$_.PowerState -eq "poweredOff"}
					ForEach ($VM in $VMs){
					$VMView = $VM | Get-View
					$VMXPath = $VMView.Config.Files.VMPathName
					$DataStoreLength =$VMXPath.IndexOf(']') - $VMXPath.IndexOf('[')-1
					$VMFolderPath=$VMXPath.IndexOf('.vmx') - $VMXPath.IndexOf('] ')
	
					$VMXPathLength= $VMXPath.length
	
					$DataStoreName=$VMXPath.substring(1,$DataStoreLength  )

					$VMFolder =$VMXPath.substring($DataStoreLength+3 ,$VMXPathLength- $DataStoreLength-3 )
		
					#Write-Host "VM Name :" $VM.name
					#Write-Host "Host Name:" $managedHost
					#Write-Host "DataStore Name :" $DataStoreName
					#Write-Host "Full Path:" $VMFolder
		
					$VMFolder=$VMFolder.substring(0,$VMFolder.IndexOf('/'))
					$VMXFolderPath=$VMXPath.substring($VMXPath.IndexOf('/')+1,$VMXPath.length -$VMXPath.IndexOf('/')-1  )
			    	#write-host "Folder Name:" $VMFolder
					#Write-Host "VMX Path:" $VMXFolderPath
					#Write-Host "------------------------------------------------------------------------"
	
					$dsv=Get-VMHost $managedHostName|Get-Datastore $DataStoreName|Get-View

					$fqFlags = New-Object VMware.Vim.FileQueryFlags
					$fqFlags.FileSize = $true
					$fqFlags.FileType = $true
					$fqFlags.Modification = $true
					$searchSpec = New-Object VMware.Vim.HostDatastoreBrowserSearchSpec
					$searchSpec.details = $fqFlags
					$searchSpec.matchPattern = $VMXFolderPath
	
					#Write-Host "Host Name:" $managedHostName
					$dsBrowser = Get-View (Get-VMHost $managedHostName|Get-Datastore $DataStoreName|Get-View).browser
	
					$rootPath = "["+$dsv.summary.Name+"]"
		
					if ($dsBrowser.Client.Version -like "Vim4*") {
						$searchSpec = [VMware.Vim.VIConvert]::ToVim41($searchSpec)
						$searchSpec.details.fileOwnerSpecified = $true
						$dsBrowserMoRef = [VMware.Vim.VIConvert]::ToVim41($dsBrowser.MoRef)
						$searchTaskMoRef = $dsBrowser.Client.VimService.SearchDatastoreSubFolders_Task($dsBrowserMoRef, $rootPath, $searchSpec)
						$searchResult = [VMware.Vim.VIConvert]::ToVim($dsBrowser.WaitForTask([VMware.Vim.VIConvert]::ToVim($searchTaskMoRef)))
						} 
					else {
						trap {
							if ($_.Exception.MethodFault -is [VMware.Vim.FileNotFound]) {
								continue
								} 
							elseif ($_.Exception.MethodFault -is [VMware.Vim.NoPermission]) {
								Write-Warning $_
								continue
								}
							else {
								Write-Warning $_
								continue
								}
							}
						$searchResult = $dsBrowser.SearchDatastoreSubFolders($rootPath, $searchSpec)
						}
		
	
					foreach ($fPath in $searchResult){
						$folderPath = $fPath.folderPath
						if ($fPath.File){
							$fPath.File | ForEach-Object {
						
								$fileSize = $_.FileSize/1024/1024
								$lastModified = $_.Modification
								
								$Diffyear =(get-date).year - (get-date $lastModified).year  
								$Diffmonth=(get-date).month - (get-date $lastModified).month
								$DiffDay =(get-date).Day - (get-date $lastModified).Day
								$DiffTime = $Diffyear * 365 + $Diffmonth * 30 + $DiffDay
								if ($DiffTime -gt 90){
									$row=""|select vmname ,hostname ,status,datastore,lastmodified,owner
									$row.vmname=$VM.Name
									$row.hostname=$managedHost
									$row.status="poweredOff"
									$row.datastore=$DataStoreName
									#$row."Folder Name"=$VMFolder
									$row.lastmodified=$lastModified
                                    $row.owner=($VM | Get-Annotation -CustomAttribute Owner).value
									#Write-Host $row -foregroundcolor Darkblue -backgroundcolor white $lastModified
													
									$PowerOffVMs += $row
									
									#Write-Host $PowerOffVMs -foregroundcolor DarkGreen -backgroundcolor white $lastModified
									}
								}
							}	
						}		
		
					}
				}
			}
            
#$PowerOffVMs | Export-Csv "C:\Wayllen\Build-BVT\Utility-Scripts\PoweredOffVMs.csv" -NoTypeInformation


# Writing Report.
# The Report Title:
$MyReport = Get-CustomHTML "Advantage Virtual Resource vCheck"

# For Generic host and VM info report section:

If (($DeptShow | Measure-Object).count -gt 0) {
    
    $MyReport += Get-CustomHeader "Dept. Resource Info" "The following gives brief capacity information ADV Dept."
	$MyReport += Get-HTMLTable $DeptShow
					
    $MyReport += Get-CustomHeaderClose
    
    if ($enableExport -eq "true"){
    $DeptShow | Export-Csv -Path "C:\Wayllen\Build-BVT\Utility-Scripts\PMS data folder\DeptInfo.csv"  -NoTypeInformation
    }

}

# For Monitor host CPU/MeM info report section:
If (($HostPerfInfo | Measure-Object).count -gt 0) {
	$MyReport += Get-CustomHeader "Host Performance Info" "The following gives brief Host performance information for all Advantage team Host based on average CPU/Mem usage "
	$MyReport += Get-HTMLTable $HostPerfInfo
	$MyReport += Get-CustomHeaderClose
    
    if ($enableExport -eq "true"){
    $HostPerfInfo | Export-Csv -Path "C:\Wayllen\Build-BVT\Utility-Scripts\PMS data folder\HostInfo.csv"  -NoTypeInformation
    }
}

# For finding the powered off vm more than 3 months report section:
$MyReport += Get-CustomHeader0 "Sleeping VM List------------------"
If (($PowerOffVMs | Measure-Object).count -gt 0) {
				$MyReport += Get-CustomHeader "Checking VM Info in PowerOff State more than 3 Months"
				$MyReport += Get-HTMLTable $PowerOffVMs
				$MyReport += Get-CustomHeaderClose
                
                if ($enableExport -eq "true"){
                    $PowerOffVMs | Export-Csv -Path "C:\Wayllen\Build-BVT\Utility-Scripts\PMS data folder\VMClearanceCheck.csv" -NoTypeInformation
                 }
}
$MyReport += Get-CustomHeader0Close


# Close the report:
$MyReport += Get-CustomHTMLClose


# Export to HTML file.
#$MyReport | Out-File "C:\Wayllen\Build-BVT\Utility-Scripts\PMS data folder\AdvCheck.html" 


$EmailTo = "your email address"
$EmailFrom = "sample@love.com"
$SMTPSRV = "SMTP server ip"

send-SMTPmail $EmailTo $EmailFrom "Advantage Team vCheck Report-$(get-date -f G)" $SMTPSRV $MyReport

$VIServer | Disconnect-VIServer -Confirm:$false


#Put the CSV files to remote Linux FTF server
ftp -v -d -s: .\FTPCommands.ftp


#Invoke the PMS web service POST method to get all resource date updated.
$request = [System.Net.WebRequest]::Create("http://IP:port/admin/resources/vms/update")
$request.Method = "POST"
$request.ContentType = "application/x-www-form-urlencoded"
$request.Timeout = 8000
$response = $request.GetResponse()
write-host $response.statuscode
$response.Close()
