Import-Csv .\AllUsers_Tasks_v5.csv | Where-Object { $_.Origin -eq 'DOCSCAN' } | Format-List *
