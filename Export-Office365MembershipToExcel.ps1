#Get current folder and set up Excel document
$filePath =  (Split-Path $MyInvocation.MyCommand.Path)
$today = Get-Date -Format "yyyy-MM-dd"
$fileName = $filePath + "\Office 365 Group Membership as " + $today + ".xlsx"

#Run the membersgip script and import members in a variable
Write-Host "Running Get-Office365GroupMembership.ps1 script..."
$groups = Invoke-Expression ((Split-Path $MyInvocation.InvocationName) + "\Get-Office365GroupMembership.ps1")

# Open the Excel document and possition on the ADUSers worksheet
Write-Host "Creating file " $fileName "..."
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$Workbook = $Excel.Workbooks.Add();
$Worksheet = $Workbook.Worksheets.Add()
$Worksheet.Name = "DistributionGroup"

# Set variables for the worksheet cells, and for navigation
$cells = $Worksheet.Cells
$row = 1
$col = 1

# Add the headers to the worksheet
Write-Host "Writing the header"
$headers = "GroupType", "GroupName",  "GroupEmail", "GroupAccess", "MemberName", "MemberEmail", "MemberType"

$headers | foreach {
    $cells.item($row, $col) = $_
    $col++
}

# Add the results from the DataTable object to the worksheet
foreach ($user in $groups) {
    $row++
    if($user.GroupType -ne ""){
        Write-Host "  ..writing data on row " $row ": " $user.GroupName " - " $user.MemberName
        $col = 1
        $cells.item($row, $col) = $user.GroupType
        $col++
        $cells.item($row, $col) = $user.GroupName
        $col++
        $cells.item($row, $col) = $user.GroupEmail
        $col++
        $cells.item($row, $col) = $user.GroupAccess
        $col++
        $cells.item($row, $col) = $user.MemberName
        $col++
        $cells.item($row, $col) = $user.MemberEmail
        $col++
        $cells.item($row, $col) = $user.MemberType
        $col++
    }
}


# Set the width of the columns automatically
Write-Host "Applying columns autofit..."
$Worksheet.Columns.Item("A:O").EntireColumn.AutoFit() | Out-Null
Add-Type -AssemblyName "Microsoft.Office.Interop.Excel"

# Apply formating
$ListObject = $Excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $Excel.ActiveCell.CurrentRegion, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
$ListObject.Name = "GroupsTable"
$ListObject.TableStyle = "TableStyleMedium6"

#Save the workbook
Write-Host "Saving file " $fileName
$Worksheet.SaveAs($fileName)

#Close the workbook and exit Excel
Write-Host "Closing the workbook and quitting Excel..."
$Workbook.Close($true)
$Excel.Quit()


Write-Host "Script completed!"

