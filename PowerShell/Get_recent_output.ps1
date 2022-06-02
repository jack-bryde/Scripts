# Find most recent cubes (automated SQL output) and move to current working directory (assumed to be empty)
$FYfolderPath = "Z:\Cubes"
# gci | where {is folder} | sort date created | select most recently created
$FYfolder = Get-ChildItem $FYfolderPath | ? { $_.PSIsContainer } | sort CreationTime -desc | select -f 1
# Then copy files
$PBIfolder = "C:\Current\Cubes"
Copy-Item -Path ($FYfolder.FullName + "\*.txt") -Destination $PBIfolder
