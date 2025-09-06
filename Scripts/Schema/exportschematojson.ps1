# Export SQL Server schema to JSON with correct foreign key detection

# --- CONFIGURATION ---
$server = "MYPC\SQLEXPRESS"
$database = "GigGearCoversDb"
$outputPath = "C:\Users\ksmar\visualStudioProjects\GGC\Scripts\Schema\DailySchema.json"

# --- CONNECTION STRING ---
$connectionString = "Server=$server;Database=$database;Integrated Security=True;Encrypt=False;TrustServerCertificate=True"

# --- SQL QUERY ---
$query = @"
SELECT
    t.name AS TableName,
    c.name AS ColumnName,
    ty.name AS DataType,
    c.max_length AS MaxLength,
    c.is_nullable AS IsNullable,
    c.is_computed AS IsComputed,
    cc.definition AS ComputedDefinition,
    fk.name AS ForeignKeyName,
    rt.name AS ReferencedTable,
    rc.name AS ReferencedColumn
FROM sys.tables t
JOIN sys.columns c ON t.object_id = c.object_id
JOIN sys.types ty ON c.user_type_id = ty.user_type_id
LEFT JOIN sys.computed_columns cc ON c.object_id = cc.object_id AND c.column_id = cc.column_id
LEFT JOIN sys.foreign_key_columns fkc ON c.object_id = fkc.parent_object_id AND c.column_id = fkc.parent_column_id
LEFT JOIN sys.foreign_keys fk ON fkc.constraint_object_id = fk.object_id
LEFT JOIN sys.tables rt ON fkc.referenced_object_id = rt.object_id
LEFT JOIN sys.columns rc ON fkc.referenced_object_id = rc.object_id AND fkc.referenced_column_id = rc.column_id
ORDER BY t.name, c.column_id
"@

# --- RUN QUERY ---
$connection = New-Object System.Data.SqlClient.SqlConnection $connectionString
$connection.Open()
$command = $connection.CreateCommand()
$command.CommandText = $query
$reader = $command.ExecuteReader()
$table = New-Object System.Data.DataTable
$table.Load($reader)
$connection.Close()

# --- LOGGING ---
if ($table.Rows.Count -eq 0) {
    Add-Content "C:\Schema\debuglog.txt" "No results returned at $(Get-Date)"
    Write-Host "No results returned from SQL query."
    exit
} else {
    Add-Content "C:\Schema\debuglog.txt" "$($table.Rows.Count) rows returned at $(Get-Date)"
    Write-Host "$($table.Rows.Count) rows returned from SQL query."
}

# --- GROUP AND FORMAT ---
$grouped = $table | Group-Object TableName | ForEach-Object {
    $tableName = $_.Name
    $columns = $_.Group | Select-Object ColumnName, DataType, MaxLength, IsNullable, IsComputed, ComputedDefinition
    # Only include true foreign keys (all three fields must be non-null and non-empty)
    $foreignKeys = $_.Group | Where-Object {
    ($_.ReferencedTable -ne $null -and $_.ReferencedTable -ne "") -and
    ($_.ReferencedColumn -ne $null -and $_.ReferencedColumn -ne "")
} | Select-Object ColumnName, ReferencedTable, ReferencedColumn

    [PSCustomObject]@{
        name = $tableName
        columns = $columns
        foreignKeys = $foreignKeys
    }
}

# --- WRAP IN DATABASE OBJECT ---
$schema = [PSCustomObject]@{
    database = $database
    generatedOn = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    tables = $grouped
}

# --- EXPORT TO JSON ---
$schema | ConvertTo-Json -Depth 5 | Out-File $outputPath -Encoding UTF8
Write-Host "✅ Schema exported to $outputPath"