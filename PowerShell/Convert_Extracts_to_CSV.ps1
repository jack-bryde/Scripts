# Iterate xlsx extracts from SOMEWHERE, remove some columns, convert to csv, save to other folder.
# Overall this doesn't appear to be faster than R, as I had hoped.

# Remove old files, so the cycle of life and death can start anew
$currentCSV = "C:\Current CSV"
Get-ChildItem -Path $currentCSV | foreach { $_.Delete()}

# Then, for each .xlsx, convert it to a csv, and save to previous folder
# (note, Excel object (see MS ComObject Application documentation) seems to allow VBA methods on workbooks & worksheets)
$currentXL = "C:\AIR Data - COVID\Current"
Get-ChildItem $currentXL | ForEach-Object {
    $XL = New-Object -ComObject Excel.Application
    $XL.Visible = $false
    $XL.DisplayAlerts = $false
    $wb = $XL.Workbooks.Open($_.FullName)
    # Iterate through each worksheet. Works even with only one.
    foreach ($ws in $wb.Worksheets) {
        #$n = $_.Name + "_" + $ws.Name ### (For multiple worksheets.)

        # Delete unnecessary columns (known from manual inspection). Order Right to left.
        $ws.Range("AV:AV").Delete()
        $ws.Range("AP:AS").Delete()
        $ws.Range("AA:AN").Delete()
        $ws.Range("Y:Y").Delete()
        $ws.Range("L:U").Delete()
        $ws.Range("J:J").Delete()
        $ws.Range("A:H").Delete()

        #save output
        $ws.SaveAs($currentCSV + "\" + $_.BaseName + ".csv", 6)
    }
    $XL.Quit()
}
