# Define the directories to check
$directories = @(
    "C:\_Paycom Frontline Data Transfer\Import_to_Frontline\_logs",
    "C:\_Paycom Frontline Data Transfer\Import_to_Frontline\Employees",
    "C:\_Paycom Frontline Data Transfer\Import_to_Frontline\Substitutes",
    "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\Employees\Completed_imports",
    "C:\_Paycom Frontline Data Transfer\Export_from_Paycom\Substitutes\Completed_imports"
)

# Set the cutoff date
$cutoffDate = (Get-Date).AddDays(-14)

foreach ($dir in $directories) {
    if (Test-Path $dir) {
        Write-Host "Processing: $dir" -ForegroundColor Cyan

        Get-ChildItem -Path $dir -File -Recurse |
            Where-Object { $_.LastWriteTime -lt $cutoffDate } |
            ForEach-Object {
                Write-Host "  Deleting: $($_.FullName)" -ForegroundColor Yellow
                Remove-Item -Path $_.FullName -Force
            }
    } else {
        Write-Host "Directory not found: $dir" -ForegroundColor Red
    }
}

Write-Host "Done." -ForegroundColor Green