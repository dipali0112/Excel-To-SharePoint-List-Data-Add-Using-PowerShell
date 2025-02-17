# Excel to SharePoint List Sync using PowerShell (For Lookup Columns) 🚀

## Overview
This PowerShell script automates the process of importing data from a local Excel file (.xlsx) into a SharePoint Online List. It utilizes the `Import-Excel` module to read data and `PnP PowerShell` for SharePoint integration. The script dynamically resolves lookup columns (such as ExpenseCategory, PaymentMethod, Status, and Department) and efficiently updates or adds list items.

## Features ✨
- ✅ Reads data from an Excel file (.xlsx) stored locally.
- ✅ Maps lookup columns (e.g., ExpenseCategory, PaymentMethod, Status, Department) to SharePoint List IDs dynamically.
- ✅ Supports updating existing records and adding new ones.
- ✅ Uses PnP PowerShell for seamless SharePoint integration.
- ✅ Handles date formats and missing data errors.

## Requirements ⚙️
To run this script, you need:
- 📌 **PnP PowerShell module** (Install using `Install-Module PnP.PowerShell -Force -Scope CurrentUser`)
- 📌 **ImportExcel PowerShell module** (Install using `Install-Module ImportExcel`)
- 📌 **SharePoint Online access** with necessary permissions to read and update the list.

## Installation 📝
### Step 1: Install Required PowerShell Modules
```powershell
Install-Module PnP.PowerShell -Force -Scope CurrentUser
Install-Module ImportExcel
```

### Step 2: Connect to SharePoint Online
```powershell
Connect-PnPOnline -Url "https://yoursharepointsite.sharepoint.com/sites/yourSite" -UseWebLogin
```

### Step 3: Run the Script
Save the script as `ExcelToSharePoint.ps1` and execute it using:
```powershell
.\ExcelToSharePoint.ps1
```

## Usage 
1. Update the **Excel file path**, **SharePoint List name**, and **lookup field names** in the script.
2. Run the script in PowerShell **after logging into SharePoint Online**.
3. Data from Excel will be automatically synchronized to the SharePoint List.

## Example Code 📝
```powershell
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
    
    $expenseCategoryId = (Get-PnPListItem -List "ExpenseCategoryList" -Query "$Row.'Expense Category'").FieldValues["ID"]
    $paymentId = (Get-PnPListItem -List "PaymentMethodList" -Query "$Row.'Payment Method'").FieldValues["ID"]
    $StatusId = (Get-PnPListItem -List "Status" -Query "$Row.Status").FieldValues["ID"]
    $departmentId = (Get-PnPListItem -List "DepartmentList" -Query "$Row.Department").FieldValues["ID"]
    
    Add-PnPListItem -List $ListName -Values @{
        "Title" = $Row.'Expense ID'
        "Date" = $dateValue
        "ExpenseCategory" = $expenseCategoryId
        "Amount" = $Row.'Amount ($)'
        "BudgetAllocated" = [decimal]$Row.'Budget Allocated ($)'
        "BudgetUtilization" = $Row.'Budget Utilization(%)'
        "PaymentMethod" = $paymentId
        "Vendor_x002f_Supplier" = $Row.'Vendor/Supplier'
        "Status" = $StatusId
        "ApprovalDate" = $approvalDateValue
        "ApproverName" = $Row.'Approver Name'
        "Department" = $departmentId
        "EmployeeName" = $Row.'Employee Name'
        "EmployeeID" = $Row.'Employee ID'
    }
}

# Disconnect from SharePoint
Disconnect-PnPOnline
```

## GitHub Tags 🔍
- PowerShell
- SharePoint Automation
- Excel to SharePoint List
- PnP PowerShell
- Import Data
- SharePoint List Bulk Upload

