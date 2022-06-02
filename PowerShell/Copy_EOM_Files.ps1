# First find the most recent FY folder, then most recent fortnight folder
# Assumes sensible naming convention (eg yyyy mm dd) that sorts both alphabetically & numerically
# gci | where {is folder} | sort name z->a | select most recent 
$allFYfolders = "Z:\20*"
$FYfolder = Get-ChildItem $allFYfolders | ? { $_.PSIsContainer } | Sort-Object Name -Descending | Select-Object -f 1
$allMonthFolders = $FYfolder.FullName
# Get most recent fortnight folder (FPR, )
$currentFN = Get-ChildItem $allMonthFolders | ? { $_.PSIsContainer } | Sort-Object Name -Descending | Select-Object -f 1
#To find fortnight for most recent end of month, take most recent month number and subtract 1
$prevMonthNum = $currentFN.name.substring(5,2) - 1
#put the leading 0 back on (if single digit) and then find most recent folder with that number
$prevMonth = "{0:d2}" -f $prevMonthNum
$prevMonthFN = Get-ChildItem -Path ($allMonthFolders + "\*" + $prevMonth + "*") | ? { $_.PSIsContainer } | Sort-Object Name -Descending | Select-Object -f 1

# Prompt user to decide which fortnight folder to use (current or latest EOM)
$title    = "Hol' up"
$question = "Is the most recent fortnight the last in the reporting month?"
# Set up popup options
$YesNoButtons =  4
$QuestionIcon = 32
$Options = $YesNoButtons + $QuestionIcon
# Zero means = no timeout
$Timeout = 0 
# Creat WSH Shell object and launch popup
$WshShell = New-Object -ComObject WScript.Shell
$Result = $WshShell.Popup($question, $Timeout, $title, $Options) 
# Cleanup
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null
# Handle input
switch ($Result)
{
    6 {"You selected Yes. Using most recent fortnight..."
    $useThisFN = $currentFN}
    7 {"You selected No. Using fortnight for most recent end of month..."
    $useThisFN = $prevMonthFN}
}

# Then copy files from selected fortnight's folder. Note use of * for 
$PBIabsentFolder = "Z:\Absent"
Copy-Item -Path ($useThisFN.FullName + "\Absent - bla -*") -Destination $PBIAbsentFolder
$PBIeeoFolder = "Z:\EEO"
Copy-Item -Path ($useThisFN.FullName + "\XXX bla_bla_EEO_*") -Destination $PBIeeoFolder
$PBIestabsFolder = "Z:\Estabs"
Copy-Item -Path ($useThisFN.FullName + "\XXX bla_bla_Establishment_*") -Destination $PBIfestabsFolder
