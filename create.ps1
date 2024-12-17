#################################################
# HelloID-Conn-Prov-Target-Zermelo-Create
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
                if ($ErrorRecord.Exception.InnerException.Message){
                    $errorObject.FriendlyMessage = $($ErrorRecord.Exception.InnerException.Message)
                } else {
                    $streamReaderResponse = [System.IO.StreamReader]::new($ErrorRecord.Exception.Response.GetResponseStream()).ReadToEnd()
                    if (-not[string]::IsNullOrEmpty($streamReaderResponse)){
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
            $headers = [System.Collections.Generic.Dictionary[[String],[String]]]::new()
            $headers.Add('Authorization', "Bearer $($actionContext.Configuration.Token)")

            $splatParams = @{
                Uri         = "$baseUrl/$Endpoint"
                Headers     = $Headers
                Method      = $Method
                ContentType = $ContentType
            }

            if ($Body){
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
    # By default, we assume that both the user and student account are not present
    $isUserAccountCreated = $false
    $IsStudentAccountCreated = $false
    $outputContext.AccountReference = 'Currently not available'

    # Define the empty array of actions that will be processed during enforcement
    $actions = @()

    if ([string]::IsNullOrEmpty($($actionContext.Data.code))) {
        throw 'Mandatory attribute [code] is empty. Please make sure it is correctly mapped'
    }
    if ([string]::IsNullOrEmpty($actionContext.Data.startDate)) {
        throw 'The mandatory property [startDate] used to look up the department is empty. Please verify your script mapping.'
    }

    # Validate correlation configuration
    if ($actionContext.CorrelationConfiguration.Enabled) {
        $correlationField = $actionContext.CorrelationConfiguration.AccountField
        $correlationValue = $actionContext.CorrelationConfiguration.AccountFieldValue

        if ([string]::IsNullOrEmpty($($correlationField))) {
            throw 'Correlation is enabled but not configured correctly'
        }
        if ([string]::IsNullOrEmpty($($correlationValue))) {
            throw 'Correlation is enabled but [PersonFieldValue] is empty. Please make sure it is correctly mapped'
        }
    }

    # Exclude departmentOfBranch fields from the actionContext.Data to retrieve only the fields managed by HelloID
    # We also create a new object 'actionContextDataFiltered' for a solid compare
    $excludedFields = 'schoolName', 'classRoom', 'participationWeight', 'startDate'
    $actionContextDataFiltered = [PSCustomObject]@{}
    foreach ($property in $actionContext.Data.PSObject.Properties) {
        if ($property.Name -notin $excludedFields) {
            $actionContextDataFiltered | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
        }
    }

    # Validate the user account
    try {
        $filteredFields = $actionContextDataFiltered.PSObject.Properties.Name
        $responseUser = Get-ZermeloAccount -Code $actionContext.Data.code -Type 'users' -Fields ($filteredFields -join ',')
        if ($null -ne $responseUser) {
            $isUserAccountCreated = $true
        }
    } catch {
        if ($_.Exception.Response.StatusCode -eq 'NotFound') {
            $isUserAccountCreated = $false
        } else {
            throw
        }
    }

    # Validate the student account
    try {
        $responseStudent = Get-ZermeloAccount -Code $actionContext.Data.code -Type 'students'
        if ($null -ne $responseStudent) {
            $isStudentAccountCreated = $true
        }
    } catch {
        if ($_.Exception.Response.StatusCode -eq 'NotFound') {
            $IsStudentAccountCreated = $false
        } else {
            throw
        }
    }

    # Validate if we need to update the department
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
            if ($null -ne $departmentToAssign){
                $actions += 'Update-DepartmentOfBranch'
                $dryRunMessageDepartmentOfBranchToAssign = "SchoolName: [$($actionContext.Data.schoolName) $($actionContext.Data.startDate)] for classRoom: [$($actionContext.Data.classRoom)] will be assigned"
            } else {
                $dryRunMessageDepartmentOfBranchToAssign = "A classroom with schoolName: [$($actionContext.Data.schoolName) $($actionContext.Data.startDate)] for classRoom: [$($actionContext.Data.classRoom)] cannot be found"
            }
        } catch {
            throw
        }
    }

    # If both the user and student account don't exist, create the user account (with the isStudent set to true) and correlate
    # Note that 'isStudent = $true' will automatically create the student account
    if (-not($isUserAccountCreated) -and (-not($isStudentAccountCreated))){
        $actions += 'Create-Correlate'
    }

    # If we have a user account but no student account, update the user account (with isStudent) and correlate
    # Note that 'isStudent = $true' will automatically create the student account
    if ($isUserAccountCreated -eq $true -and -not $isStudentAccountCreated){
        $actions += 'Create-StudentAccount-Correlate-User'
    }

    # If we have a student account but no user account, create the user account (with isStudent) and correlate
    # Note that 'isStudent = $true' will automatically create the student account
    if ($isStudentAccountCreated -eq $true -and -not $isUserAccountCreated){
        $actions += 'Create-Correlate'
    }

    # If we have both a user and student account, match the userCode. If a match is found, correlate
    if ($isUserAccountCreated -and $isStudentAccountCreated){
        if ($responseUser.code -eq $responseStudent.userCode) {
            $actions += 'Correlate'
            $outputContext.AccountReference = $responseUser.code
        }
    }

    # Add a message and the result of each of the validations showing what will happen during enforcement
    if ($actionContext.DryRun -eq $true) {
        Write-Information "[DryRun] $action Zermelo account for: [$($personContext.Person.DisplayName)], will be executed during enforcement" -Verbose
        if ($null -ne $dryRunMessageDepartmentOfBranchToAssign){
            Write-Information "[DryRun] $dryRunMessageDepartmentOfBranchToAssign"
        }
    }

    # Separate 'Update-DepartmentOfBranch' from other actions to make sure we always loop through the actions in a specific order
    $orderedActions = @($actions | Where-Object { $_ -notin @('Update-DepartmentOfBranch') })
    if ($actions -contains 'Update-DepartmentOfBranch') {
        $orderedActions += 'Update-DepartmentOfBranch'
    }

    # Process
    if (-not($actionContext.DryRun -eq $true)) {
        foreach ($action in $orderedActions){
            switch ($action) {
                'Create-Correlate' {
                    Write-Information 'Creating and correlating Zermelo account'
                    $splatCreateUserParams = @{
                        Endpoint    = 'users'
                        Method      = 'POST'
                        Body        = $actionContextDataFiltered | ConvertTo-Json
                        ContentType = 'application/json'
                    }
                    $responseCreateUserAccount = Invoke-ZermeloRestMethod @splatCreateUserParams

                    $outputContext.AccountCorrelated = $true
                    $outputContext.Data = $responseCreateUserAccount.response.data
                    $outputContext.AccountReference = $responseCreateUserAccount.response.data.code

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Action  = 'CreateAccount'
                        Message = "Create-Correlate account was successful. AccountReference is: [$($outputContext.AccountReference)]"
                        IsError = $false
                    })
                    break
                }

                'Create-StudentAccount-Correlate-User'{
                    Write-Information 'Creating Zermelo student account and correlating user account'
                    $splatCreateStudentParams = @{
                        Endpoint    = "users/$($responseUser.code)"
                        Method      = 'PUT'
                        Body        = @{
                            isStudent = $actionContext.Data.isStudent
                        } | ConvertTo-Json
                        ContentType = 'application/json'
                    }
                    $null = Invoke-ZermeloRestMethod @splatCreateStudentParams

                    $outputContext.AccountCorrelated = $true
                    $outputContext.Data = $responseUser
                    $outputContext.AccountReference = $responseUser.code

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Action  = 'CreateAccount'
                        Message = "Create-StudentAccount-Correlate-User account was successful. AccountReference is: [$($outputContext.AccountReference)]"
                        IsError = $false
                    })
                    break
                }

                'Correlate' {
                    Write-Information 'Correlating Zermelo user account'
                    $outputContext.AccountCorrelated = $true
                    $outputContext.Data = $responseUser
                    $outputContext.AccountReference = $responseUser.code

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Action  = 'CreateAccount'
                        Message = "Correlate account was successful. AccountReference is: [$($outputContext.AccountReference)]"
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
                            student             = $outputContext.Data.code
                            participationWeight = $actionContext.Data.participationWeight
                        } | ConvertTo-Json
                        ContentType = 'application/json'
                    }
                    try {
                        $null = Invoke-ZermeloRestMethod @splatStudentInDepartmentParams
                        $auditLogMessage = 'DepartmentOfBranch created'
                    } catch {
                        if ($_.Exception.StatusCode -eq 409){
                            $auditLogMessage = 'DepartmentOfBranch already exists'
                        } else {
                            throw
                        }
                    }

                    $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Message = "Update-DepartmentOfBranch was successful with message: [$auditLogMessage]. Department set to: [$($departmentToAssign.schoolInSchoolYearName)] with id: [$($departmentToAssign.id)]"
                        IsError = $false
                    })
                    break
                }
            }
        }
        $outputContext.success = $true
    }
} catch {
    $outputContext.success = $false
    $errorObject = Resolve-ZermeloError -ErrorRecord $_
    $auditMessage = "Could not $action Zermelo account. Error: $($errorObject.FriendlyMessage)"
    Write-Warning "Error at Line '$($_.InvocationInfo.ScriptLineNumber)': $($_.InvocationInfo.Line). Error: $($errorObject.ErrorDetails)"
    $outputContext.AuditLogs.Add([PSCustomObject]@{
        Message = $auditMessage
        IsError = $true
    })
}
