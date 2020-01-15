#Builds an asset list based on several CSV's of random asset data, by using AD lookups to create a 'normalized' set of data pulled from Active Directory from wildcard searches based on the CSV.

$O365CSV = Import-Csv 'C:\AssetList\O365License.csv'
$SccmCSV = Import-csv 'C:\AssetList\AllSystems.csv'
$MobCSV = Import-csv 'C:\AssetList\PhoneList.csv'

#Create a hasf of AD users for fast lookup
$ADLookup = @()
$ADLookup =  Get-ADUser -Filter * -Properties ProxyAddresses, Manager, Department

$O365Table = @()
$SccmTable = @()
$MobTable = @()


#Search for user based on the email field being one of the email addresses present in the proxy address attribute.
#Pulls device info from device column of CSV
$O365CSV | ForEach-Object {
    
    $UPN = $_.UserPrincipalName
    $Lookup = Try {$ADLookup |? {$_.UserPrincipalName -eq $UPN} -ErrorAction Stop} Catch {}
    
    $Obj = [PSCustomObject]@{
        DisplayName = $_."UserPrincipalName"
        Employee = $Lookup.Name
        Device = $_.Device
    }
    $O365Table += $Obj
}


$SccmCSV | ForEach-Object {
    #Search for user based on the SAM account name field being one of the email addresses present in the proxy address attribute.
    #Pulls device info from device column of CSV    
    
    $SAM = $_.SamName
    $Lookup = Try { $ADLookup |? {$_.SamAccountName -eq $SAM} } Catch {}
    
    If ($Lookup -eq $Null){
        $Lookup = Try { $ADLookup |? {$_.Name -eq $Name} } Catch {}
    }

    $Manager = Try {(Get-ADUser (($Lookup).Manager)).Name} Catch {}

    $Obj = [PSCustomObject]@{
        DisplayName = ($_."SamName")
        Employee = $Lookup.Name
        Device = ($_.Manufacturer + " " + $_.Model)
        Serial = $_."SerialNumber"
        Manager = $Manager
        Department = $Lookup.Department
    }
    $SCCMTable += $Obj
}

$MobCSV | ForEach-Object {
      
    $Proxy = $_.Email
    $Proxy = ("*" + $Proxy + "*")
    $Lookup = Try { $ADLookup |? {$_.ProxyAddresses -like $Proxy} -EA SilentlyContinue} Catch {}
    
        
    $Phone = $_.Device
    $Split = $Phone.Split(" ")
    
    $Obj = [PSCustomObject]@{
        DisplayName = $_.Email  
        Employee = $Lookup.Name
        Device = $_.Device
    }
    $MoTable += $Obj
}


# Create Master Excel Table

$table = New-Object system.Data.DataTable “$tabName”

$col1 = New-Object system.Data.DataColumn Employee,([string])
$col2 = New-Object system.Data.DataColumn DisplayName,([string])
#$col3 = New-Object system.Data.DataColumn ComputerName,([string])
$col3 = New-Object system.Data.DataColumn Device,([string])
#$col4 = New-Object system.Data.DataColumn Manufacturer,([string])
#$col5 = New-Object system.Data.DataColumn Model,([string])
$col4 = New-Object system.Data.DataColumn SerialNumber,([string])
$col5 = New-Object system.Data.DataColumn Manager,([string])
$col6 = New-Object system.Data.DataColumn Department,([string])
#$col4 = New-Object system.Data.DataColumn Mobile,([string])
#$col5 = New-Object system.Data.DataColumn License,([string])

#Add required column numbers

$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)
#$table.columns.add($col7)
#$table.columns.add($col8)
#$table.columns.add($col9)
#$table.columns.add($col4)
#$table.columns.add($col5)

#Now add data from each hash created from CSV


$O365Table | ForEach-Object {

$row = $table.NewRow()
$row.DisplayName = $_.DisplayName
$row.Employee = $_.Employee
$row.Device = $_.Device
#$row.License = $_.License

$table.Rows.Add($row)
}


$SccmTable | ForEach-Object {

$row = $table.NewRow()
$row.DisplayName = $_.DisplayName
$row.Employee = $_.Employee
#$row.ComputerName = $_.ComputerName
$row.Device = $_.Device
#$row.Manufacturer = $_.Manufacturer
#$row.Model = $_.Model
$row.SerialNumber = $_.Serial
$row.Manager = $_.Manager
$row.Department = $_.Department
$table.Rows.Add($row)
}


$MobTable | ForEach-Object {

$row = $table.NewRow()
$row.DisplayName = $_.DisplayName
$row.Employee = $_.Employee
$row.Device = $_.Device
#$row.Device = $_.Mobile

$table.Rows.Add($row)
}


#Export table to CSV.

$Table | ConvertTo-Csv -NoTypeInformation | Out-File C:\AssetList\AssetDatabase.csv
