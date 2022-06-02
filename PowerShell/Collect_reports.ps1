# Collect all files from a folder, including subfolders, and copy elsewhere.
# Originally used to collate all reports of a category (eg Absent) from multiple reporting periods

# First find the most recent year's folder, then iterate over each fortnight
# gci | where {is folder} | sort name z->a | select most recent 
$allFYfolders = "T:\bla\20*"
$FYfolder = Get-ChildItem $allFYfolders | ? { $_.PSIsContainer } | Sort-Object Name -Descending | Select-Object -f 1
$allMonthFolders = $FYfolder.FullName
write-output $allMonthFolders

#Iterate through each folder
$folders = Get-ChildItem -Path $allMonthFolders -Recurse | Where-Object { $_.PSIsContainer }
foreach ($fortnight in $folders) {
    # Then copy files (repeat for each file/category to copy)
    $PBIabsentFolder = "D:\Absent"
    Copy-Item -Path ($fortnight.FullName + "\Absent - blabla -*") -Destination $PBIAbsentFolder
}
