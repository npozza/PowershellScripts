$ous = Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName

$table = New-Object System.Data.DataTable
$table.Columns.Add("Name")
$table.Columns.Add("Distinguished Name")
$table.Columns.Add("Description")
$table.Columns.Add("Creation Date")
$table.Columns.Add("Linked GPOs")
$table.Columns.Add("Object Class")
$table.Columns.Add("Child Item Count")

Foreach ($ou in $ous){

$ouinfo = Get-ADOrganizationalUnit -Filter 'DistinguishedName -like $ou.distinguishedname' -Properties Name, DistinguishedName, Description, created, LinkedGroupPolicyObjects, objectClass

$count = Get-ADObject -Filter * -SearchBase $ou | where {($_.DistinguishedName -ne $ou.DistinguishedName)} | Measure-Object | select -ExpandProperty Count

$r = $table.NewRow()
$r.Name = $ouinfo.Name
$r.'Distinguished Name' = $ouinfo.DistinguishedName
$r.Description = $ouinfo.Description
$r.'Creation Date' = $ouinfo.Created
$r.'Linked GPOs' = $ouinfo.LinkedGroupPolicyObjects
$r.'Object Class' = $ouinfo.ObjectClass
$r.'Child Item Count' = $count
$table.Rows.Add($r)

}

$table | Export-Excel -Path C:\temp\OUs.xlsx -IncludePivotTable -PivotTableName "Empty OUs" -PivotRows "Distinguished Name" -PivotData @{'Child Item Count'=''} -IncludePivotChart -ChartType PieExploded3D

$table | Export-csv -Path C:\temp\OUs.csv -Encoding UTF8 -Delimiter ";" -Append