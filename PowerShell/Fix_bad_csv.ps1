# A large .csv file that contained " as a field value, breaking upon import by Excel
# (works one file at a time, put in loop for folder)
$ADT = Get-ChildItem "T:\file_path\file.csv"
$content = Get-Content $ADT.FullName -Raw
# New name removed the datestamp in each file
$newName = $ADT.Name.substring($ADT.Name.Length - 15, 15).Replace("_", " ")
# Replace the miscreant chars with a word and save modified file with new name to show it has been adjusted
$content.Replace('"',"QuoteMark") | Set-Content -Path $newName
Set-Content $newName
