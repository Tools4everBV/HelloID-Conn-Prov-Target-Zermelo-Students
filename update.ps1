#################################################
# HelloID-Conn-Prov-Target-Zermelo-Update
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

#region functions
function Get-ZermeloAccount {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Code,

        [Parameter(Mandatory)]
        [string]
        $Type,

        [Parameter()]
        [string]
        $Fields
    )

    $splatParams = @{
        Method = 'GET'
    }

    switch ($Type) {
        'users' {
            if ($Fields){
                $fields = "$fields,code"
                $splatParams['Endpoint'] = "users/$($Code)?fields=$($Fields.Trim("'"))"
            } else {
                $splatParams['Endpoint'] = "users/$Code"
            }
            (Invoke-ZermeloRestMethod @splatParams).response.data
        }

        'students' {
            $splatParams['Endpoint'] = "students/$Code"
            (Invoke-ZermeloRestMethod @splatParams).response.data
        }
    }
}

function Get-DepartmentToAssign {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $SchoolName,

        [Parameter()]
        [string]
        $DepartmentName,

        [Parameter()]
        [string]
        $SchoolYear
    )

    try {
        $splatParams = @{
            Method   = 'GET'
            Endpoint = 'departmentsofbranches'
        }
        $responseDepartments = (Invoke-ZermeloRestMethod @splatParams).response.data

        if ($null -ne $responseDepartments) {
            $lookup = $responseDepartments | Group-Object -AsHashTable -Property 'code'
            $departments = $lookup[$DepartmentName]
            $departmentToAssign = $departments | Where-Object { $_.schoolInSchoolYearName -match "$schoolNameToMatch $SchoolYear" }
            Write-Output $departmentToAssign
        }
    } catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Get-CurrentSchoolYear {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [DateTime]
        $ContractStartDate
    )

    $currentDate = Get-Date
    $year = $currentDate.Year

    # Determine the start and end dates of the current school year
    if ($currentDate.Month -lt 8) {
        $startYear = $year - 1
    } else {
        $startYear = $year
    }

    $schoolYearStartDate = (Get-Date -Year $startYear)

    Write-Output $schoolYearStartDate
}


function ConvertTo-HashTableToObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $HashTableString
    )

    $trimmedString = $HashTableString.TrimStart('@{').TrimEnd('}')
    $keyValuePairs = $trimmedString -split ';'
    $hashTable = @{}
    foreach ($pair in $keyValuePairs) {
        $key, $value = $pair -split '=', 2
        $hashTable[$key.Trim()] = $value.Trim()
    }

    Write-Output $hashTable
}

function Get-NestedPropertyValue {
    param (
        [object]
        $Object,

        [string]
        $PropertyPath
    )

    $properties = $PropertyPath -split '\.'
    foreach ($property in $properties) {
        if ($null -eq $Object -or -not $Object.PSObject.Properties[$property]) {
            return $null
        }

        try {
            $Object = $Object | Select-Object -ExpandProperty $property -ErrorAction Stop
        } catch {
            return $null
        }
    }
    Write-Output $Object
}

function Resolve-ZermeloError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorRecord
    )
    process {
        $errorObject = [PSCustomObject]@{
            ScriptLineNumber = $ErrorRecord.InvocationInfo.ScriptLineNumber
            Line             = $ErrorRecord.InvocationInfo.Line
            ErrorDetails     = $ErrorRecord.Exception.Message
            FriendlyMessage  = $ErrorRecord.Exception.Message
        }

        try {
            if ($ErrorRecord.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') {
                $rawErrorObject = ($ErrorRecord.ErrorDetails.Message | ConvertFrom-Json).response
                $errorObject.ErrorDetails = "Code: [$($rawErrorObject.status)], Message: [$($rawErrorObject.message)], Details: [$($rawErrorObject.details)], EventId: [$($rawErrorObject.eventId)]"
                $errorObject.FriendlyMessage = $rawErrorObject.message
            } elseif ($ErrorRecord.Exception.GetType().FullName -eq 'System.Net.WebException') {
                if ($ErrorRecord.Exception.InnerException.Message) {
                    $errorObject.FriendlyMessage = $($ErrorRecord.Exception.InnerException.Message)
                } else {
                    $streamReaderResponse = [System.IO.StreamReader]::new($ErrorRecord.Exception.Response.GetResponseStream()).ReadToEnd()
                    if (-not[string]::IsNullOrEmpty($streamReaderResponse)) {
                        $rawErrorObject = ($streamReaderResponse | ConvertFrom-Json).response
                        $errorObject.ErrorDetails = "Code: [$($rawErrorObject.status)], Message: [$($rawErrorObject.message)], Details: [$($rawErrorObject.details)], EventId: [$($rawErrorObject.eventId)]"
                        $errorObject.FriendlyMessage = $rawErrorObject.message
                    }
                }
            } elseif ($ErrorRecord.Exception.GetType().FullName -eq 'System.Net.Http.HttpRequestException') {
                $errorObject.FriendlyMessage = $($ErrorRecord.Exception.Message)
            } else {
                $errorObject.FriendlyMessage = $($ErrorRecord.Exception.Message)
            }
        } catch {
            $errorObject.FriendlyMessage = "Received an unexpected response, error: $($ErrorRecord.Exception.Message)"
        }

        Write-Output $errorObject
    }
}

function Invoke-ZermeloRestMethod {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Method,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Endpoint,

        [object]
        $Body,

        [string]
        $ContentType = 'application/json'
    )

    process {
        $baseUrl = "$($actionContext.Configuration.BaseUrl)/api/v3"
        try {
            $headers = [System.Collections.Generic.Dictionary[[String], [String]]]::new()
            $headers.Add('Authorization', "Bearer $($actionContext.Configuration.Token)")

            $splatParams = @{
                Uri         = "$baseUrl/$Endpoint"
                Headers     = $Headers
                Method      = $Method
                ContentType = $ContentType
            }

            if ($Body) {
                Write-Information 'Adding body to request'
                $splatParams['Body'] = $Body
            }
            Invoke-RestMethod @splatParams -Verbose:$false
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}
#endregion


try {
    # Verify if [aRef] has a value
    if ([string]::IsNullOrEmpty($($actionContext.References.Account))) {
        throw 'The account reference could not be found'
    }

    # Exclude departmentOfBranch fields from the actionContext.Data to retrieve only the fields managed by HelloID
    # We also create a new object 'actionContextDataFiltered' for a solid compare
    $excludedFields = 'schoolName', 'classRoom', 'participationWeight', 'startDate', 'isStudent', 'code'
    $actionContextDataFiltered = [PSCustomObject]@{}
    foreach ($property in $actionContext.Data.PSObject.Properties) {
        if ($property.Name -notin $excludedFields) {
            $actionContextDataFiltered | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
        }
    }

    Write-Information "Verifying if a Zermelo account for [$($personContext.Person.DisplayName)] exists"
    try {
        $filteredFields = $actionContextDataFiltered.PSObject.Properties.Name
        $correlatedAccount = Get-ZermeloAccount -Code $actionContext.References.Account -Type 'users' -Fields ($filteredFields -join ',')
        $outputContext.PreviousData = $correlatedAccount
    } catch {
        throw
    }

    if (-not [string]::IsNullOrEmpty($actionContext.Data.schoolName) -and
    -not [string]::IsNullOrEmpty($actionContext.Data.classRoom) -and
    $actionContext.Data.startDate -ne [DateTime]::MinValue) {
        Write-Information 'Determine school year based on the startDate specified in [actionContext.Data.startDate]'
        $currentSchoolYear = Get-CurrentSchoolYear -ContractStartDate $($actionContext.Data.startDate)
        $schoolYearToMatch = "$($currentSchoolYear.Year)" + '-' + "$($currentSchoolYear.AddYears(1).Year)"

        Write-Information 'Determine which departmentOfBranch will need to be assigned'
        try {
            $splatGetDepartmentToAssign = @{
                SchoolName        = $actionContext.Data.schoolName
                DepartmentName    = $actionContext.Data.classRoom
                SchoolYear        = $schoolYearToMatch
            }
            $departmentToAssign = Get-DepartmentToAssign @splatGetDepartmentToAssign
            $dryRunMessageDepartmentOfBranchToAssign = "SchoolName: [$($actionContext.Data.schoolName) $($actionContext.Data.startDate)] for classRoom: [$($actionContext.Data.classRoom)] will be assigned"
        } catch {
            throw
        }
    }

    # Define the empty array of actions that will be processed during enforcement
    $actions = @()

    # Check for changes within the personDifferences object
    Write-Information 'Verify if the user account must be updated'
    if ($null -ne $correlatedAccount) {
        $splatCompareProperties = @{
            ReferenceObject  = @($correlatedAccount.PSObject.Properties)
            DifferenceObject = @($actionContextDataFiltered.PSObject.Properties)
        }
        $userPropertiesChanged = (Compare-Object @splatCompareProperties -PassThru).Where({$_.SideIndicator -eq '=>'})
        if ($userPropertiesChanged -and ($null -ne $correlatedAccount)) {
            $actions += 'Update-Account'
            $dryRunMessage = "Account property(s) required to update: [$($userPropertiesChanged.name -join ",")]"

            $updateObject = [PSCustomObject]@{}
            foreach ($property in $userPropertiesChanged) {
                $updateObject | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
            }
        } elseif (-not($userPropertiesChanged)) {
            $actions += 'NoChangesToUser'
            $dryRunMessage = 'No changes will be made to the account during enforcement'
        }
    } else {
        $actions += 'UserNotFound'
        $dryRunMessage = "Zermelo account for: [$($personContext.Person.DisplayName)] not found. Possibly deleted"
    }

    # A change to either the school or classroom will always result in an assignment of a new 'departmentOfBranch'
    # A 'departmentOfBranch' always includes both the school, year and classroom information
    Write-Information 'Verify if the school or classroom  must be updated'
    $departmentUpdated = $false

    if ($null -eq $departmentToAssign) {
        $actions += 'DepartmentOfBranchNotFound'
    } else {
        # Check if the school must be updated
        $schoolValue = Get-NestedPropertyValue -Object $personContext.PersonDifferences.PrimaryContract -PropertyPath $actionContext.Configuration.SchoolNameField
        if (-not [string]::IsNullOrEmpty($schoolValue)) {
            $pdHash = ConvertTo-HashTableToObject -HashTableString $schoolValue
            if (($pdHash.Change -eq 'updated') -and ($actionContext.Data.schoolName -match $pdHash.New)) {
                $actions += 'Update-DepartmentOfBranch'
                $departmentUpdated = $true
            }
        }

        # Check if the classroom must be updated
        $classroomValue = Get-NestedPropertyValue -Object $personContext.PersonDifferences.PrimaryContract -PropertyPath $actionContext.Configuration.ClassroomField
        if (-not [string]::IsNullOrEmpty($classroomValue)) {
            $pdHash = ConvertTo-HashTableToObject -HashTableString $classroomValue
            if (($pdHash.Change -eq 'updated') -and ($actionContext.Data.classRoom -match $pdHash.New)) {
                if (-not $departmentUpdated) {
                    $actions += 'Update-DepartmentOfBranch'
                    $departmentUpdated = $true
                }
            }
        }

        # If no department or school update was made
        if (-not $departmentUpdated) {
            $actions += 'NoChangesToDepartmentOfBranch'
        }
    }

    # Add a message and the result of each of the validations showing what will happen during enforcement
    if ($actionContext.DryRun -eq $true) {
        Write-Information "[DryRun] $dryRunMessage"
        Write-Information "[DryRun] $dryRunMessageDepartmentOfBranchToAssign"
    }

    # Process actions
    if (-not($actionContext.DryRun -eq $true)) {
        foreach ($action in $actions) {
            switch ($action) {
                'Update-Account' {
                    Write-Information "Updating Zermelo account with accountReference: [$($actionContext.References.Account)]"
                    Write-Information "Account property(s) required to update: $($propertiesChanged.Name -join ', ')"
                    $updateObject | Add-Member -MemberType NoteProperty -Name 'code' -Value $actionContext.References.Account
                    $splatUpdateUserParams = @{
                        Endpoint    = "users/$($actionContext.References.Account)"
                        Method      = 'PUT'
                        Body        =  $updateObject | ConvertTo-Json
                        ContentType = 'application/json'
                    }
                    $null = Invoke-ZermeloRestMethod @splatUpdateUserParams

                    $outputContext.success = $true
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Message = "Update account was successful, Account property(s) updated: [$($propertiesChanged.name -join ',')]"
                        IsError = $false
                    })
                    break
                }

                'Update-DepartmentOfBranch' {
                    Write-Information "Updating departmentOfBranch for Zermelo account with accountReference: [$($actionContext.References.Account)]"
                    Write-Information "New department: [$($departmentToAssign.schoolInSchoolYearName)] with id: [$($departmentToAssign.id)] will be assigned"
                    $splatStudentInDepartmentParams = @{
                        Endpoint    = 'studentsindepartments'
                        Method      = 'POST'
                        Body        = @{
                            departmentOfBranch  = $departmentToAssign.id
                            student             = $correlatedAccount.code
                            participationWeight = $actionContext.Data.participationWeight
                        } | ConvertTo-Json
                        ContentType = 'application/json'
                    }
                    $null = Invoke-ZermeloRestMethod @splatStudentInDepartmentParams

                    $outputContext.success = $true
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Message = "Update department was successful. Department updated to: [$($departmentToAssign.schoolInSchoolYearName)] with id: [$($departmentToAssign.id)]"
                        IsError = $false
                    })
                    break
                }

                'NoChangesToUser' {
                    Write-Information "No changes to Zermelo account with accountReference: [$($actionContext.References.Account)]"

                    $outputContext.success = $true
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Message = 'No changes will be made to the account during enforcement'
                        IsError = $false
                    })
                    break
                }

                'NoChangesToDepartmentOfBranch' {
                    Write-Information 'No changes will be made to the department or school during enforcement'

                    $outputContext.success = $true
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Message = 'No changes will be made to the department or school during enforcement'
                        IsError = $false
                    })
                    break
                }

                'UserNotFound' {
                    Write-Information "Zermelo account for: [$($personContext.Person.DisplayName)] not found. Possibly deleted"

                    $outputContext.success = $false
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Message = "Zermelo account for: [$($personContext.Person.DisplayName)] not found. Possibly deleted"
                        IsError = $true
                    })
                    break
                }

                'DepartmentOfBranchNotFound' {
                    Write-Information "A departmentOfBranch for school: [$($actionContext.Data.schoolName)] year: [$($actionContext.Data.startDate)] and classroom [$($actionContext.Data.classRoom)] could not be found. Possibly deleted"

                    $outputContext.success = $false
                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Message = "A departmentOfBranch for school: [$($actionContext.Data.schoolName)] year: [$($actionContext.Data.startDate)] and classroom [$($actionContext.Data.classRoom)] could not be found. Possibly deleted"
                        IsError = $true
                    })
                    break
                }
            }
        }
    }
} catch {
    $outputContext.Success = $false
    $errorObject = Resolve-ZermeloError -ErrorRecord $_
    $auditMessage = "Could not update Zermelo account. Error: $($errorObject.FriendlyMessage)"
    Write-Warning "Error at Line '$($_.InvocationInfo.ScriptLineNumber)': $($_.InvocationInfo.Line). Error: $($errorObject.ErrorDetails)"
    $outputContext.AuditLogs.Add([PSCustomObject]@{
        Message = $auditMessage
        IsError = $true
    })
}
