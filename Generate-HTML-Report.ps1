# Generate HTML Report from CSV Output
# This script converts the CSV output into a formatted HTML report

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\M365ObjectCounts_Report.html"
)

# Import CSV data
$data = Import-Csv -Path $CsvPath

# Generate HTML
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Object Counts Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
        }
        .header p {
            font-size: 1.1em;
            opacity: 0.9;
        }
        .content {
            padding: 30px;
        }
        .section {
            margin-bottom: 40px;
        }
        .section-title {
            font-size: 1.8em;
            color: #0078d4;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        thead {
            background: #0078d4;
            color: white;
        }
        th {
            padding: 15px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.9em;
            letter-spacing: 0.5px;
        }
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .count {
            font-weight: bold;
            color: #0078d4;
            font-size: 1.1em;
        }
        .manual-check {
            color: #ff6b6b;
            font-style: italic;
        }
        .footer {
            text-align: center;
            padding: 20px;
            background: #f5f5f5;
            color: #666;
            font-size: 0.9em;
        }
        .summary-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        .card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            text-align: center;
        }
        .card-title {
            font-size: 0.9em;
            opacity: 0.9;
            margin-bottom: 10px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .card-value {
            font-size: 2.5em;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Microsoft 365 Object Counts Report</h1>
            <p>Generated on $(Get-Date -Format "MMMM dd, yyyy 'at' hh:mm tt")</p>
        </div>
        
        <div class="content">
            <div class="summary-cards">
"@

# Add summary cards for key metrics
$teamsCount = ($data | Where-Object { $_.ObjectType -eq "Teams sites" }).Count
$guestsCount = ($data | Where-Object { $_.ObjectType -eq "Guests" }).Count
$totalSites = ($data | Where-Object { $_.Category -eq "SharePoint Online" } | Measure-Object -Property Count -Sum).Sum

if ($teamsCount) {
    $html += @"
                <div class="card">
                    <div class="card-title">Teams Sites</div>
                    <div class="card-value">$teamsCount</div>
                </div>
"@
}

if ($guestsCount) {
    $html += @"
                <div class="card">
                    <div class="card-title">Guest Users</div>
                    <div class="card-value">$guestsCount</div>
                </div>
"@
}

$html += @"
            </div>
"@

# Group data by category
$categories = $data | Group-Object -Property Category

foreach ($category in $categories) {
    $html += @"
            <div class="section">
                <h2 class="section-title">$($category.Name)</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Object Type</th>
                            <th>Count / Quantity</th>
                        </tr>
                    </thead>
                    <tbody>
"@
    
    foreach ($item in $category.Group) {
        $countClass = if ($item.Count -eq "N/A") { "manual-check" } else { "count" }
        $html += @"
                        <tr>
                            <td>$($item.ObjectType)</td>
                            <td class="$countClass">$($item.Count)</td>
                        </tr>
"@
    }
    
    $html += @"
                    </tbody>
                </table>
            </div>
"@
}

$html += @"
        </div>
        
        <div class="footer">
            <p>Generated by Microsoft 365 Object Counts Extraction Script</p>
            <p>Source CSV: $CsvPath</p>
        </div>
    </div>
</body>
</html>
"@

# Save HTML file
$html | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host "✅ HTML report generated successfully!" -ForegroundColor Green
Write-Host "📄 Output file: $OutputPath" -ForegroundColor Cyan
Write-Host "`nOpening report in default browser..." -ForegroundColor Yellow

# Open in default browser
Start-Process $OutputPath

Write-Host "✅ Done!" -ForegroundColor Green
