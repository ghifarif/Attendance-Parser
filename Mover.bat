$messageTable = @{
    0 = "No files were copied. No failure was met. No files were mismatched. The files already exist in the destination directory; so the copy operation was skipped."
    1 = "All files were copied successfully."
    2 = "There are some additional files in the destination directory that aren't present in the source directory. No files were copied."
	3 = "Some files were copied. Additional files were present. No failure was met."
	4 = "Some Mismatched files or directories were detected. Examine the output log. Housekeeping might be required."
	5 = "Some files were copied. Some files were mismatched. No failure was met."
	6 = "Additional files and mismatched files exist. No files were copied and no failures were met. Which means that the files already exist in the destination directory."
	7 = "Files were copied, a file mismatch was present, and additional files were present."
	8 = "Some files or directories could not be copied."
	16 = "Several files didn't copy."
}

function ValidateHost{
	Param($hostname)
	if(!(Test-Connection -BufferSize 32 -Count 1 -ComputerName $hostname -Quiet)) {return $FALSE}

	return $TRUE
}

function RunSync{
	Param($source,$dest)
	$date       = Get-Date -UFormat "%Y%m%d-%s"
	$Logfile    = "D:\LOG\$branch-$date.txt"
	$what       = @("/MOV","/FFT","/Z","/E")
	$options    = @("/R:3","/W:10","/NP","/NDL","/LOG:$logfile")
	$cmdArgs    = @("$source","$dest",$what,$options)
	robocopy @cmdArgs
	
	return $lastexitcode
}

function WriteLog{
	Param($branch,$hostname,$PingResult,$RespondCode,$source,$dest)
	$log_date 			= Get-Date -UFormat "%Y%m%d-%R:%S"
	$logfile_status 	= "D:\LOG\HRMS-StatusLog.log"
	if($PingResult -eq $FALSE){
		$log_date + "-[ERROR]The remote computer " + $hostname +" for " + $branch + " is Offline." | out-file -Filepath $logfile_status -append
		exit}
	
	$message = $messageTable[$RespondCode]
	if($RespondCode -gt 7){
		$log_date + "-[FAILED]Moving HRMS data from " + $source + " to " +  $dest + " finished with respond code : " + $lastexitcode + $message | out-file -Filepath $logfile_status -append
		exit}
	$log_date + "-[SUCCESS]Moving HRMS data from " + $source + " to " +  $dest + " finished with respond code : " + $lastexitcode + $message | out-file -Filepath $logfile_status -append
	exit
}
##########################################################################################################################################################
$branch	     = "JAKARTA"
$hostname    = "JkcPcLgl001"
$source      = "\\$hostname\c$\Users\Admin\Desktop\HRMS - 01 JAKARTA"
$dest        = "D:\HRMS\HRMS - 01 JAKARTA"

$PingResult = ValidateHost -hostname $hostname
if($PingResult -eq $FALSE){WriteLog -branch $branch -hostname $hostname -PingResult $PingResult}

$RespondCode = RunSync -source $source -dest $dest
WriteLog -branch $branch -hostname $hostname -RespondCode $RespondCode -
