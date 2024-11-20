# Install the necessary module for working with SQL Server and Analysis Services
# Uncomment to install if not already done
# Install-Module -Name SqlServer
# Get-Module -ListAvailable -Name SqlServer

#-----------------------------------------------------------
param (
    [string]$servername  # Allow passing the server name as a parameter
)

# Load Analysis Services assemblies required for connecting to Power BI Analysis Services
# Ensure the paths match your SQL Server SDK installation
Add-Type -Path "C:\Program Files\Microsoft SQL Server\150\Setup Bootstrap\Update Cache\KB5046859\GDR\x64\MICROSOFT.ANALYSISSERVICES.CORE.DLL"
Add-Type -Path "C:\Program Files\Microsoft SQL Server\150\Setup Bootstrap\Update Cache\KB5046859\GDR\x64\MICROSOFT.ANALYSISSERVICES.TABULAR.DLL"

# Connect to the Analysis Services instance
# Use the provided server name to create a connection string
$server = New-Object Microsoft.AnalysisServices.Tabular.Server
$connectionString = "DataSource=$serverName"

try {
    $server.Connect($connectionString)
    Write-Host "Connected to Power BI Analysis Services instance on $serverName" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to the Analysis Services instance: $_"
    return
}

# Analyze the model to identify measures with errors
# Retrieve the model from the first database and inspect its measures
$database = $server.Databases[0]  # Assuming only one database
$model = $database.Model

# List to store measures with errors
$measuresWithErrors = @()

foreach ($table in $model.Tables) {
    foreach ($measure in $table.Measures) {
        if ($measure.ErrorMessage -ne '') {
            # Add the measure with its error details to the list
            $errorDetails = @{
                "Measure Name"  = $measure.Name
                "Table Name"    = $table.Name
                "Folder Name"   = $measure.DisplayFolder
                "Error Message" = $measure.ErrorMessage
            }
            $measuresWithErrors += New-Object PSObject -Property $errorDetails
        }
    }
}

# Disconnect from the Analysis Services instance
$server.Disconnect()
Write-Host "Disconnected from Power BI Analysis Services instance." -ForegroundColor Cyan

# Display results of measures with errors
# Show measures with errors in a table format, or confirm no errors were found
if ($measuresWithErrors.Count -gt 0) {
    Write-Host "`nMeasures with Errors:" -ForegroundColor Black -BackgroundColor Yellow
    $measuresWithErrors | Format-Table "Measure Name", "Table Name", "Folder Name", "Error Message" -AutoSize
} else {
    Write-Host "No objects with errors found!" -ForegroundColor Black -BackgroundColor Yellow
}

# Provide instructions for the user to close the window
Write-Host "Press Enter to close this window..." -ForegroundColor Green
Read-Host
