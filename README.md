Excel to SharePoint List Sync using PowerShell (For Lookup Columns) 🚀

This PowerShell script automates the process of importing data from a local Excel file 📊 into a SharePoint Online List 🌐. It reads data using the Import-Excel module, resolves lookup columns (such as ExpenseCategory, PaymentMethod, Status, and Department), and adds or updates list items in SharePoint via PnP PowerShell.

Features ✨

✅ Reads data from an Excel file (.xlsx) stored locally.
✅ Maps lookup columns (e.g., ExpenseCategory, PaymentMethod, Status, Department) to SharePoint List IDs dynamically.
✅ Supports updating existing records and adding new ones.
✅ Uses PnP PowerShell for seamless SharePoint integration.
✅ Handles date formats and missing data errors.

Requirements ⚙️

📌 PnP PowerShell module (Install-Module PnP.PowerShell -Force -Scope CurrentUser).
📌 ImportExcel PowerShell module (Install-Module ImportExcel).
📌 SharePoint Online access with necessary permissions.

Usage 🫠️

Update the Excel file path, SharePoint List name, and lookup field names in the script.
Run the script in PowerShell after logging in (Connect-PnPOnline).
Data from Excel will be synced to SharePoint automatically.

Installation 📝

Install required PowerShell modules:
Install-Module PnP.PowerShell -Force -Scope CurrentUser
Install-Module ImportExcel
Connect to SharePoint Online:
Connect-PnPOnline -Url "https://futurrizoninterns.sharepoint.com/sites/Company'sFinancial" -UseWebLogin
Run the script:
.\ExcelToSharePoint.ps1

Example Code 📝

# Load Excel File and Read Data
$ExcelFilePath = "C:\Users\91915\Downloads\Data_1.xlsx"
$ExcelData = Import-Excel -Path $ExcelFilePath

# SharePoint list name
$ListName = "MainList"

foreach ($Row in $ExcelData) {
  $dateValue = $null
  $approvalDateValue = $null

  if ([string]::IsNullOrWhiteSpace($Row.Date) -eq $false) {
    $dateValue = [datetime]::ParseExact($Row.Date, 'dd/MM/yyyy', $null)
  }
  if ([string]::IsNullOrWhiteSpace($Row.'Approval Date') -eq $false) {
    $approvalDateValue = [datetime]::ParseExact($Row.'Approval Date', 'dd/MM/yyyy', $null)
  }

  $expenseCategoryId = (Get-PnPListItem -List "ExpenseCategoryList" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$Row.'Expense Category'</Value></Eq></Where></Query></View>").FieldValues["ID"]
  $paymentId = (Get-PnPListItem -List "PaymentMethodList" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$Row.'Payment Method'</Value></Eq></Where></Query></View>").FieldValues["ID"]
  $StatusId = (Get-PnPListItem -List "Status" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$Row.Status</Value></Eq></Where></Query></View>").FieldValues["ID"]
  $departmentId = (Get-PnPListItem -List "DepartmentList" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$Row.Department</Value></Eq></Where></Query></View>").FieldValues["ID"]

  Add-PnPListItem -List $listName -Values @{
    "Title"                 = $Row.'Expense ID'
    "Date"                  = $dateValue
    "ExpenseCategory"       = $expenseCategoryId
    "Amount"                = $Row.'Amount ($)'
    "BudgetAllocated"       = [decimal]$Row.'Budget Allocated ($)'
    "BudgetUtilization"     = $Row.'Budget Utilization(%)'
    "PaymentMethod"         = $paymentId
    "Vendor_x002f_Supplier" = $Row.'Vendor/Supplier'
    "Status"                = $StatusId
    "ApprovalDate"          = $approvalDateValue
    "ApproverName"          = $Row.'Approver Name'
    "Department"            = $departmentId
    "EmployeeName"          = $Row.'Employee Name'
    "EmployeeID"            = $Row.'Employee ID'		
  }
}

# Disconnect from SharePoint
Disconnect-PnPOnline

GitHub Tags🔍

PowerShell SharePoint Automation
Excel to SharePoint List Upload
PnP PowerShell Import Data
Sync Excel with SharePoint
SharePoint List Bulk Upload
