function Install-TervisPaylocity {
    param (
        [Parameter(Mandatory)]$PathToPaylocityDataExport,
        [Parameter(Mandatory)]$PaylocityDepartmentsWithNiceNamesJsonPath
    )
    Set-PathToPaylocityDataExport -PathToPaylocityDataExport $PathToPaylocityDataExport
    Set-PaylocityDepartmentsWithNiceNamesJsonPath -PaylocityDepartmentsWithNiceNamesJsonPath $PaylocityDepartmentsWithNiceNamesJsonPath
}

function Get-PathToPaylocityDataExport {   
    $Env:PathToPaylocityDataExport
}

function Set-PathToPaylocityDataExport {
    param (
        [Parameter(Mandatory)][String]$PathToPaylocityDataExport
    )
    [Environment]::SetEnvironmentVariable("PathToPaylocityDataExport", $PathToPaylocityDataExport, "User")
}

function Get-PaylocityEmployees {
    param(
        [ValidateSet("A","T")]$Status
    )
    
    if (-not $Script:PaylocityEmployees) {

        $PathToPaylocityDataExport = Get-PathToPaylocityDataExport

        $MostRecentPaylocityDataExport = Get-ChildItem -File $PathToPaylocityDataExport | sort -Property CreationTime -Descending | select -First 1
        [xml]$Content = Get-Content $MostRecentPaylocityDataExport.FullName
        $Details = $Content.Report.CustomReportTable.Detail_Collection.Detail

        $PaylocityEmployees = ForEach ($Detail in $Details) {
            [pscustomobject][ordered]@{
                Organization = $Detail.col10 | ConvertTo-TitleCase
                State = $Detail.col9
                Status = $Detail.col8
                DepartmentName = $Detail.col7
                DepartmentCode = $Detail.col6
                JobTitle = $Detail.col5 | ConvertTo-TitleCase
                ManagerEmployeeID = $Detail.col4
                ManagerName = $Detail.col3 | ConvertTo-TitleCase
                Surname = $Detail.col2 | ConvertTo-TitleCase
                GivenName = $Detail.col1 | ConvertTo-TitleCase
                EmployeeID = $Detail.col0
                TerminationDate = if ($Detail.col11) {Get-Date $Detail.col11}
            } |
            Add-Member -MemberType ScriptProperty -Name DepartmentNiceName -PassThru -Value {
                Get-DepartmentNiceName -PaylocityDepartmentName $this.DepartmentName
            } |
            Add-Member -Name DepartmentRoleSAMAccountName -MemberType ScriptProperty -PassThru -Force -Value {
                "Role_Paylocity$($this.DepartmentCode)"
            } |
            Add-Member -Name DepartmentRoleName -MemberType ScriptProperty -PassThru -Force -Value {
                "Role_Paylocity$($this.DepartmentName)"
            }
        }
    
        $Script:PaylocityEmployees = $PaylocityEmployees 
    }
    
    $Script:PaylocityEmployees | 
    Where { -not $Status -or $_.Status -eq $Status }
}

function Get-PaylocityEmployeesEmployeeIDHashValue {
    param (
        $EmployeeID
    )
    
    if (-not $Script:PaylocityEmployeesEmployeeIDHash) {
        $PaylocityEmployeesEmployeeIDHash = @{}

        Get-PaylocityEmployees |
        ForEach-Object -Process {
            $PaylocityEmployeesEmployeeIDHash.Add($_.EmployeeId , $_)
        }

        $Script:PaylocityEmployeesEmployeeIDHash = $PaylocityEmployeesEmployeeIDHash
    }

    $Script:PaylocityEmployeesEmployeeIDHash[$EmployeeID]
}


function Get-PaylocityEmployee {
    param (
        $EmployeeID        
    )
    Get-PaylocityEmployeesEmployeeIDHashValue -EmployeeID $EmployeeID    
}

function Get-PaylocityDepartmentsWithNiceNamesJsonPath {
    Import-Clixml -Path $env:USERPROFILE\PaylocityDepartmentsWithNiceNamesJsonPath.xml
}

function Set-PaylocityDepartmentsWithNiceNamesJsonPath {
    param (
        $PaylocityDepartmentsWithNiceNamesJsonPath
    )
    $PaylocityDepartmentsWithNiceNamesJsonPath | Export-Clixml -Path $env:USERPROFILE\PaylocityDepartmentsWithNiceNamesJsonPath.xml
}

function Get-PaylocityDepartment {
    $PaylocityRecords = Get-PaylocityEmployees
    $(
        $PaylocityRecords | 
        group departmentname, departmentcode | 
        select -ExpandProperty name
    ) | % {
        [pscustomobject][ordered]@{
            Name = $($_ -split ", ")[0]
            Code = $($_ -split ", ")[1] 
        }|
        Add-Member -MemberType ScriptProperty -Name NiceName -PassThru -Value {
            Get-DepartmentNiceName -PaylocityDepartmentName $this.Name
        } |
        Add-Member -Name RoleSAMAccountName -MemberType ScriptProperty -PassThru -Force -Value {
            "Role_Paylocity$($this.Code)"
        } |
        Add-Member -Name RoleName -MemberType ScriptProperty -PassThru -Force -Value {
            "Role_Paylocity$($this.Name)"
        }
    }
}

function Get-PaylocityDepartmentNamesAndCodesAsPowerShellPSCustomObjectText {
    $PaylocityDepartments = Get-PaylocityDepartment
    $PaylocityDepartments | 
    sort departmentname | % {
@"
[pscustomobject][ordered]@{
    Name = "$($_.DepartmentName)"
    Code = "$($_.DepartmentCode)"
    NiceName = ""
},
"@
    }
}

function Get-DepartmentNiceName {
    param(
        $PaylocityDepartmentName
    )
    
    if (-not $Script:PaylocityDepartmentsWithNiceNames) {
        $Script:PaylocityDepartmentsWithNiceNames = Get-Content -Path $(Get-PaylocityDepartmentsWithNiceNamesJsonPath) | 
        ConvertFrom-Json
    }

    $Script:PaylocityDepartmentsWithNiceNames | 
    where DepartmentName -eq $PaylocityDepartmentName | 
    select -ExpandProperty DepartmentNiceName
}

function Get-PaylocityEmployeesGroupedByDepartment {
    $PaylocityRecords = Get-PaylocityEmployees
    $PaylocityRecords| group departmentname | sort count -Descending
}

function Get-TopLevelManager {
    param (
        $Employee,
        $EmployeesSubSet
    )
    if ($Employee.ManagerEmployeeID -notin $EmployeesSubSet.EmployeeID) {
        return $Employee.ManagerEmployeeID
    } else {
        Get-TopLevelManager -Employee ($EmployeesSubSet | where EmployeeId -eq $Employee.ManagerEmployeeID) -EmployeesSubSet $EmployeesSubSet
    }
}

filter Add-PaylocityReportDetailsCustomMembers {
    $_ | Add-Member -MemberType ScriptProperty -Name "Organization" -Value {$This.col10 | ConvertTo-TitleCase}
    $_ | Add-Member -MemberType ScriptProperty -Name "State" -Value {$This.col9}
    $_ | Add-Member -MemberType ScriptProperty -Name "Status" -Value {$This.col8}
    $_ | Add-Member -MemberType ScriptProperty -Name "DepartmentName" -Value {$This.col7}
    $_ | Add-Member -MemberType ScriptProperty -Name "DepartmentCode" -Value {$This.col6}
    $_ | Add-Member -MemberType ScriptProperty -Name "JobTitle" -Value {$This.col5 | ConvertTo-TitleCase}
    $_ | Add-Member -MemberType ScriptProperty -Name "ManagerEmployeeID" -Value {$This.col4}
    $_ | Add-Member -MemberType ScriptProperty -Name "ManagerName" -Value {$This.col3 | ConvertTo-TitleCase}
    $_ | Add-Member -MemberType ScriptProperty -Name "Surname" -Value {$This.col2 | ConvertTo-TitleCase}
    $_ | Add-Member -MemberType ScriptProperty -Name "GivenName" -Value {$This.col1 | ConvertTo-TitleCase}
    $_ | Add-Member -MemberType ScriptProperty -Name "EmployeeID" -Value {$This.col0}
}

function Get-AllActiveEmployeesWithTheirTervisEmailAddress {
    $ActiveEmployees = Get-PaylocityEmployees -Status A    
    $ADusersOfActivePaylocityUsers = Get-PaylocityADUser -Status A

    Import-TervisOffice365ExchangePSSession
    $Mailboxes = Get-O365Mailbox
    
    $ADUsersWithMailboxes = $ADusersOfActivePaylocityUsers |
    where UserPrincipalName -In $Mailboxes.UserPrincipalName

    $ActiveEmployees | Add-Member -MemberType ScriptProperty -Name EmailAddress -Force -Value {
        $ADUsersWithMailboxes | 
        Where-Object EmployeeID -eq $This.EmployeeID |
        Select-Object -ExpandProperty UserPrincipalName
    }

    $ActiveEmployees | 
    Select-Object -Property SurName, GivenName, EmailAddress, DepartmentName |
    Sort-Object -Property Surname |
    Export-Csv -Path $Home\ActiveEmployeesAndTheirWorkEmails.csv -NoTypeInformation
}