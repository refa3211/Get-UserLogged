# $headers = ("PC", "USERNAME", "SESSIONNAME", "ID", "STATE", "IDLE TIME", "Date", "TIME")
$OU = "OU=Users,OU=Company_1OU,DC=Company_1,DC=internal"

$pclist = Get-ADComputer -Filter {Enabled -eq 'True'} -SearchBase $OU


foreach ($pc in $pclist) {
    $pc = $pc.Name
    if (Test-Connection -ComputerName $pc -Count 1 -ErrorAction "Quiet") {
        Write-Host $pc
        try {
            $last_user = (quser /server:$pc | Select-Object -Skip 1) -replace ' {1,}', ',' 
            $last_user_string = $last_user.ForEach({ "$pc,$_" }) -join "`n" | Out-File .\csv-1.csv -Append  -Encoding 'utf8' 
            Write-Host "$pc,$last_user_string"
        } 
        catch {
            Write-Warning "Failed to retrieve user information from $pc"
        }
    } 
    else {
        Write-Warning "Unable to communicate with $pc"
    }
}

