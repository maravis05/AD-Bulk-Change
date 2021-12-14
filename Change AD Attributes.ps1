

Add-Type -AssemblyName System.Windows.Forms

$ADFields = "name","userprincipalname","department","title","manager"
$User = ""
$AllUsers = @()
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog


Clear-Host
Write-Host "-------- AD ATTRIBUTE BULK CHANGE --------`n"

While ($User -ne "done") {
    
    $User = Read-Host "Enter an account to be changed. (e.g. 'mmarquis')`nWhen finished, type 'done'`n"
    
    if ($User -eq "done") {Break}
    
    $UserObj = Get-ADUser $User -Properties $ADFields | Select-Object $ADFields -ErrorAction SilentlyContinue
    
    if ($Null -eq $UserObj) { 
        Write-Host $User" Not Found in AD."
        Continue }

    $ManagerName = Get-ADUser $UserObj.manager | Select-Object name
    $UserObj.manager = $ManagerName.name
    Write-Output $UserObj
    $AllUsers += $UserObj
    
    }

if ($AllUsers) {

    Read-Host "Press Enter to Save CSV Template"
    
    $SaveFileDialog.initialDirectory = "C:\temp"
    $SaveFileDialog.filter = "CSV (*.csv)| *.csv"
    $SaveFileDialog.FileName = "ADUsers.csv"
    [void]$SaveFileDialog.ShowDialog()
    $AllUsers | Export-Csv -Path $SaveFileDialog.FileName -NoTypeInformation
    $A = Read-Host "Press Enter After Making Changes to CSV or type 'quit' to quit."
    If ($A -eq "quit") {Break}

    }

$OpenFileDialog.InitialDirectory = $SaveFileDialog.FileName
$OpenFileDialog.filter = "CSV (*.csv)| *.csv"
[void]$OpenFileDialog.ShowDialog()

$CSVPath = $OpenFileDialog.FileName | Resolve-Path
$CSVFile = Import-Csv $CSVPath

$Header = (Get-Content -path $CSVPath -TotalCount 1).split(',')


$CSVFile | ForEach-Object {

    $UPN = $PSItem.userprincipalname
    $Name = $PSItem.name

    try { 

        $CurrentUser = Get-ADUser -filter 'userprincipalname -eq $UPN' -Properties $Header | Select-Object $Header
        
        if ($CurrentUser -eq $null) { $CurrentUser = Get-ADUser -filter 'name -eq $Name' -Properties $Header | Select-Object $Header }

        if ($CurrentUser -eq $null) {
            Write-Output ($Name + " " + $UPN + " not found in AD.`n")
            return }
        }

    catch { 
        Write-Host $_
        Break }
    
    $NewManagerName = $PSItem.manager
    
    try { $NewManagerObj = Get-ADUser -Filter 'name -eq $NewManagerName' | Select-Object distinguishedname }
        catch { write-host $NewManagerName" manager not found in AD." }
    
    $PSItem.manager = $NewManagerObj.distinguishedname
        
    $NewAttributes = @{}
    $NewAttributes.clear()
    
    foreach ($Attribute in $Header) {

        if ($Attribute -eq "name") {continue}
        
        if ( $PSItem.$Attribute -ne $CurrentUser.$Attribute -and [string]::IsNullOrWhiteSpace($PSItem.$Attribute) -ne $True ) { 
            $NewAttributes[$Attribute] = $PSItem.$Attribute 
            }

        else {continue}

        }
                
    if($NewAttributes.Count -ne 0) {

        Write-Host "Changing:"
        Write-Output $PSItem.name
        Write-Output $NewAttributes
        Get-ADUser -filter 'userprincipalname -eq $UPN' | Set-ADUser @NewAttributes

        }

}

Read-Host "Done. Press Enter to Quit."

