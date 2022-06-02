# Find folder for SQL output files
$PBIfolder = "Z:\Output"
# Find the most recent (current) working directory
# gci | where {is folder} | sort date created | select most recently created
$FYfolderPath = "C:\Working_directories"
# Group 1 (eg Year)
$FYfolder = Get-ChildItem $FYfolderPath | ? { $_.PSIsContainer } | sort CreationTime -desc | select -f 1
$FNfolderPath = $FYfolder.FullName
# Group 2 (eg Week)
$FNfolder = Get-ChildItem $FNfolderPath | ? { $_.PSIsContainer } | sort CreationTime -desc | select -f 1
Copy-Item -Path ($FNfolder.FullName + "\*") -Exclude *.ps1 -Destination $PBIfolder
# Get-ChildItem returns the desired files, which are piped ("|") into the Rename-Item function
# Rename-Item then reduces the length of the name by 11 characters (date, underscore and file ext).
Get-ChildItem -Path ($PBIfolder + "\" + "Qry_Export_T???_[0-9]*") | Rename-Item -newname { $_.name.substring(0,$_.name.length - 11) + $_.Extension } 
