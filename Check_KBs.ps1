Import-Module ActiveDirectory

$ou = "Enter Your Favorite OU"  
$favKBs = @("KB5050187","KB5046547","KB5048654","KB5049983")  
$outputFile = "Your_Path\InstalledUpdates.csv"
$computers = Get-ADComputer -Filter {Enabled -eq $true} -SearchBase $ou | Select-Object -ExpandProperty Name
$resultsArray = @()

$scriptBlock = {
    param($favKBs)
    $Session = New-Object -ComObject "Microsoft.Update.Session"
    $Searcher = $Session.CreateUpdateSearcher()
    $historyCount = $Searcher.GetTotalHistoryCount()
    $installedKBs = @{}
    if ($historyCount -gt 0) {
        $updates = $Searcher.QueryHistory(0, $historyCount) | 
            Select-Object Title, Date, 
            @{name="KB"; expression={
                if ($_.Title -match "KB\d+") {
                    $matches[0]
                } else {
                    "N/A"
                }
            }}

        foreach ($update in $updates) {
            $installedKBs[$update.KB] = $true
        }
    }

    return $installedKBs
}

foreach ($computer in $computers) {
    try {
        $installedKBs = Invoke-Command -ComputerName $computer -ScriptBlock $scriptBlock -ArgumentList $favKBs

        foreach ($kb in $favKBs) {
            $resultsArray += [PSCustomObject]@{
                ComputerName = $computer
                KB            = $kb
                Status        = if ($installedKBs.ContainsKey($kb)) { "Installed" } else { "Not Installed" }
            }
        }
    } catch {
        $resultsArray += [PSCustomObject]@{
            ComputerName = $computer
            KB            = "N/A"
            Status        = "Failed to connect: $_"
        }
    }
}

$resultsArray | Export-Csv -Path $outputFile -NoTypeInformation

Write-Host "Results exported to $outputFile"
