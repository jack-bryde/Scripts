## Export Power BI files from SQL Server 
# Script modified from blog:
# https://social.technet.microsoft.com/wiki/contents/articles/890.export-sql-server-blob-data-with-powershell.aspx

## Notes:
# Due to expected size of Content VarBinary(MAX), data must be streamed, not buffered
# For each PBI report on RS, there are 3 files
#  1. CatalogItem: the working PBIX file.
#  2. DataModel: the compressed Vertipaq model. Since 2022 contains the power query which cannot be separated.
#  3. RerportDefinition: seems to be same as CatalogItem without the DataModel.
# Due to compression via proprietary software, DataModel cannot be edited.
# Power query M script can be extracted from either CatalogItem or ReportDefinition:
#  - PBIX compressed M with DataModel. Must convert to PBIT (template) file.
#  - Cannot be done via automation currently.
#  - Possibly done via CMD interface of Tabular Editor.
#  - Once converted, rename file extension to ZIP and open DataModelSchema (json text file).
#  - M script is located under the "partions" object, "expression" nested array with an M step at each index.
#  - Possibly via Tabular Editor CMD interface, a DataModelSchema may be edited before saving back to PBIX.

## TODO Consider:
# - Tabular Editor CMD interface: https://docs.tabulareditor.com/te2/Command-line-Options.html
# - Re-publish to RS via SQL
# - Get/modify automatic refresh data on Server
# - Code in R

# Configuration data
$Server = "server"; # SQL Server Instance.
$Database = "ReportServer";
$Dest = "C:\BLOBOut"; # Path to export to.
$bufferSize = 8192; # Stream buffer size in bytes.

# Query (may return multiple reports)
$Sql = "
SELECT 
    c.PATH
    , ec.ContentType
    , ec.Content
FROM ReportServer.dbo.[Catalog] c 
    INNER JOIN ReportServer.dbo.CatalogItemExtendedContent ec ON ec.ItemId = c.ItemID
WHERE 
    c.[Path] LIKE  '/path/%'
";
 
# Create ADO.NET SQL Connection object
$con = New-Object Data.SqlClient.SqlConnection;
$con.ConnectionString = "Data Source=$Server;" +
                        "Integrated Security=True;" +
                        "Initial Catalog=$Database";
$con.Open();
 
# Create Command and Reader objects
$cmd = New-Object Data.SqlClient.SqlCommand $Sql, $con;
$rd = $cmd.ExecuteReader();

# Instantiate a byte array for the stream
$out = [array]::CreateInstance('Byte', $bufferSize)

# Loop through each report ($rd.Read() returns next row from table)
While ($rd.Read())
{
    # GetString() converts bytes to string
    Write-Output ("Exporting: {0}" -f $rd.GetString(0) + "\" + $rd.GetString(1));

    # File stream object (construct with output path & roles)
    $file_path = ($Dest + $rd.GetString(0) + "\" + $rd.GetString(1))
    # Ensure folder exists
    if(!(Test-Path -Path ($Dest + $rd.GetString(0)))){
        New-Item -Path ($Dest + $rd.GetString(0)) -ItemType "directory"
    }
    $fs = New-Object System.IO.FileStream $file_path, Create, Write;
    # New BinaryWriter (construct with file stream object)
    $bw = New-Object System.IO.BinaryWriter $fs;

    # Iterate the file by chunks (bufferSize above)
    $start = 0;
    # Read first (amount of) byte stream (note content in third column)
    $received = $rd.GetBytes(2, $start, $out, 0, $bufferSize - 1);
    While ($received -gt 0)
    {
       $bw.Write($out, 0, $received);
       # Flush - ensure bytes in stream are written to file
       $bw.Flush();
       $start += $received;
       # Read next byte stream
       $received = $rd.GetBytes(2, $start, $out, 0, $bufferSize - 1);
    }

    $bw.Close();
    $fs.Close();
}

# Close server connection; Free all objects
$fs.Dispose();
$rd.Close();
$cmd.Dispose();
$con.Close();

Write-Output ("Finished");
