#requires -Version 7.0
 
[CmdletBinding()]
param(
    [ValidateRange(1024, 65535)]
    [int]$Port = 8765,
 
    [switch]$Manual,
 
    [SecureString]$ManualPassword,
 
    [ValidateRange(0, 3600)]
    [int]$ManualWaitSeconds = 0
)
 
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
 
$backendRoot = Split-Path -Parent $PSScriptRoot
$originalLocation = (Get-Location).Path
$databaseName = "authentication_smoke_$([guid]::NewGuid().ToString('N')).db"
$databasePath = Join-Path $backendRoot $databaseName
$baseUri = "http://127.0.0.1:$Port"
$smokePassword = $null
$logPrefix = "schedule-maker-auth-smoke-$([guid]::NewGuid().ToString('N'))"
$stdoutPath = Join-Path ([System.IO.Path]::GetTempPath()) "$logPrefix.stdout.log"
$stderrPath = Join-Path ([System.IO.Path]::GetTempPath()) "$logPrefix.stderr.log"
$serverProcess = $null
$passedChecks = 0
$exitCode = 0
 
$originalEnvironment = @{
    DATABASE_URL = [Environment]::GetEnvironmentVariable(
        "DATABASE_URL",
        "Process"
    )
    SECRET_KEY = [Environment]::GetEnvironmentVariable(
        "SECRET_KEY",
        "Process"
    )
    LOG_LEVEL = [Environment]::GetEnvironmentVariable(
        "LOG_LEVEL",
        "Process"
    )
    AUTH_SMOKE_PASSWORD = [Environment]::GetEnvironmentVariable(
        "AUTH_SMOKE_PASSWORD",
        "Process"
    )
}
 
function Write-Pass {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )
 
    $script:passedChecks++
    Write-Host "[PASS] $Name" -ForegroundColor Green
}
 
function Invoke-ApiRequest {
    param(
        [Parameter(Mandatory)]
        [ValidateSet("GET", "POST", "PUT", "PATCH", "DELETE")]
        [string]$Method,
 
        [Parameter(Mandatory)]
        [string]$Path,
 
        [hashtable]$Headers = @{},
 
        [AllowNull()]
        [object]$Body,
 
        [string]$ContentType = "application/json"
    )
 
    $requestParameters = @{
        Method = $Method
        Uri = "$script:baseUri$Path"
        Headers = $Headers
        SkipHttpErrorCheck = $true
        TimeoutSec = 10
    }
 
    if ($PSBoundParameters.ContainsKey("Body")) {
        $requestParameters.ContentType = $ContentType
        $requestParameters.Body = if ($ContentType -eq "application/json") {
            $Body | ConvertTo-Json -Depth 10 -Compress
        }
        else {
            $Body
        }
    }
 
    Invoke-WebRequest @requestParameters
}
 
function Assert-Response {
    param(
        [Parameter(Mandatory)]
        $Response,
 
        [Parameter(Mandatory)]
        [int]$ExpectedStatus,
 
        [Parameter(Mandatory)]
        [string]$Name,
 
        [AllowNull()]
        [string]$ExpectedDetail,
 
        [switch]$ExpectBearerHeader,
 
        [switch]$ExpectNoBearerHeader
    )
 
    if ([int]$Response.StatusCode -ne $ExpectedStatus) {
        throw (
            "$Name returned $($Response.StatusCode); expected " +
            "$ExpectedStatus. Body: $($Response.Content)"
        )
    }
 
    if ($PSBoundParameters.ContainsKey("ExpectedDetail")) {
        $responseBody = $Response.Content | ConvertFrom-Json
        if ([string]$responseBody.detail -ne $ExpectedDetail) {
            throw (
                "$Name returned detail '$($responseBody.detail)'; expected " +
                "'$ExpectedDetail'."
            )
        }
    }
 
    $authenticateHeader = [string]$Response.Headers["WWW-Authenticate"]
 
    if ($ExpectBearerHeader -and $authenticateHeader -ne "Bearer") {
        throw "$Name did not return 'WWW-Authenticate: Bearer'."
    }
 
    if (
        $ExpectNoBearerHeader -and
        -not [string]::IsNullOrEmpty($authenticateHeader)
    ) {
        throw "$Name unexpectedly returned WWW-Authenticate."
    }
 
    Write-Pass $Name
}
 
function Assert-JsonProperty {
    param(
        [Parameter(Mandatory)]
        $Response,
 
        [Parameter(Mandatory)]
        [string]$PropertyName,
 
        [Parameter(Mandatory)]
        [string]$ExpectedValue,
 
        [Parameter(Mandatory)]
        [string]$Name
    )
 
    $body = $Response.Content | ConvertFrom-Json
    $property = $body.PSObject.Properties[$PropertyName]
 
    if ($null -eq $property) {
        throw "$Name did not contain property '$PropertyName'."
    }
 
    if ([string]$property.Value -ne $ExpectedValue) {
        throw (
            "$Name returned '$($property.Value)' for '$PropertyName'; " +
            "expected '$ExpectedValue'."
        )
    }
 
    Write-Pass $Name
}
 
function Assert-ListIds {
    param(
        [Parameter(Mandatory)]
        $Response,
 
        [Parameter(Mandatory)]
        [string]$Name,
 
        [string[]]$ExpectedIds = @(),
 
        [string[]]$ExcludedIds = @(),
 
        [int]$ExpectedCount = -1
    )
 
    Assert-Response `
        -Response $Response `
        -ExpectedStatus 200 `
        -Name "$Name response"
 
    $items = @($Response.Content | ConvertFrom-Json)
    $ids = @($items | ForEach-Object { [string]$_.id })
 
    foreach ($expectedId in $ExpectedIds) {
        if ($expectedId -notin $ids) {
            throw "$Name did not contain expected ID '$expectedId'."
        }
    }
 
    foreach ($excludedId in $ExcludedIds) {
        if ($excludedId -in $ids) {
            throw "$Name leaked excluded ID '$excludedId'."
        }
    }
 
    if ($ExpectedCount -ge 0 -and $ids.Count -ne $ExpectedCount) {
        throw (
            "$Name returned $($ids.Count) rows; expected $ExpectedCount. " +
            "IDs: $($ids -join ', ')"
        )
    }
 
    Write-Pass "$Name contents"
}
 
function Assert-ListOmitsProperty {
    param(
        [Parameter(Mandatory)]
        $Response,
 
        [Parameter(Mandatory)]
        [string]$PropertyName,
 
        [Parameter(Mandatory)]
        [string]$Name
    )
 
    $items = @($Response.Content | ConvertFrom-Json)
    foreach ($item in $items) {
        if ($item.PSObject.Properties.Name -contains $PropertyName) {
            throw "$Name exposed forbidden property '$PropertyName'."
        }
    }
 
    Write-Pass $Name
}
 
function Get-SmokeAccessToken {
    param(
        [Parameter(Mandatory)]
        [string]$Email,
 
        [Parameter(Mandatory)]
        [string]$RoleName
    )
 
    $response = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/login" `
        -ContentType "application/x-www-form-urlencoded" `
        -Body @{
            username = $Email
            password = $script:smokePassword
        }
 
    Assert-Response `
        -Response $response `
        -ExpectedStatus 200 `
        -Name "$RoleName login"
 
    $body = $response.Content | ConvertFrom-Json
    if ($body.token_type -ne "bearer") {
        throw "$RoleName login returned token_type '$($body.token_type)'."
    }
 
    if ([string]::IsNullOrWhiteSpace([string]$body.access_token)) {
        throw "$RoleName login returned an empty access token."
    }
 
    [string]$body.access_token
}
 
function New-BearerHeaders {
    param(
        [Parameter(Mandatory)]
        [string]$Token
    )
 
    @{
        Authorization = "Bearer $Token"
    }
}
 
function Show-ManualSession {
    param(
        [Parameter(Mandatory)]
        $SeedData
    )
 
    Write-Host ""
    Write-Host "Manual authentication test environment is ready." `
        -ForegroundColor Cyan
    Write-Host "API:     $script:baseUri"
    Write-Host "Swagger: $script:baseUri/docs"
    Write-Host ""
    Write-Host "Use the password entered when this script started."
    Write-Host "Access tokens and passwords are not printed."
    Write-Host ""
    Write-Host "Test accounts:" -ForegroundColor Cyan
    @(
        [pscustomobject]@{
            Role = "super_admin"
            Email = $SeedData.super_admin_email
            Department = "-"
        },
        [pscustomobject]@{
            Role = "department_admin A"
            Email = $SeedData.department_admin_a_email
            Department = $SeedData.department_a_id
        },
        [pscustomobject]@{
            Role = "department_admin B"
            Email = $SeedData.department_admin_b_email
            Department = $SeedData.department_b_id
        },
        [pscustomobject]@{
            Role = "doctor A"
            Email = $SeedData.doctor_user_a_email
            Department = $SeedData.department_a_id
        },
        [pscustomobject]@{
            Role = "viewer A"
            Email = $SeedData.viewer_a_email
            Department = $SeedData.department_a_id
        }
    ) | Format-Table -AutoSize | Out-Host
 
    Write-Host "Tenant resources:" -ForegroundColor Cyan
    @(
        [pscustomobject]@{
            Resource = "Department"
            DepartmentA = $SeedData.department_a_id
            DepartmentB = $SeedData.department_b_id
        },
        [pscustomobject]@{
            Resource = "Team"
            DepartmentA = $SeedData.team_a_id
            DepartmentB = $SeedData.team_b_id
        },
        [pscustomobject]@{
            Resource = "Position"
            DepartmentA = $SeedData.position_a_id
            DepartmentB = $SeedData.position_b_id
        },
        [pscustomobject]@{
            Resource = "Shift"
            DepartmentA = $SeedData.shift_a_id
            DepartmentB = $SeedData.shift_b_id
        },
        [pscustomobject]@{
            Resource = "Doctor"
            DepartmentA = $SeedData.doctor_a_id
            DepartmentB = $SeedData.doctor_b_id
        },
        [pscustomobject]@{
            Resource = "ShiftAssignment"
            DepartmentA = $SeedData.assignment_a_id
            DepartmentB = $SeedData.assignment_b_id
        }
    ) | Format-Table -AutoSize | Out-Host
 
    Write-Host "Manual threat-check sequence:" -ForegroundColor Cyan
    Write-Host "1. Log in through POST /api/v1/auth/login."
    Write-Host "2. Authorize Swagger with the returned access token."
    Write-Host "3. List Department A resources and confirm no Department B IDs."
    Write-Host "4. Request each Department B UUID as Department A and expect 404."
    Write-Host "5. Attempt cross-tenant writes and confirm no records are added."
    Write-Host "6. As Doctor A, access Doctor A unavailability and expect success."
    Write-Host "7. As Doctor A, access Doctor B unavailability and expect 403."
    Write-Host "8. Confirm /doctors/roster omits email."
    Write-Host "9. Confirm Department A admin Doctor detail includes name."
    Write-Host "10. Confirm super-admin routine tenant writes return 403."
    Write-Host ""
}
 
try {
    Set-Location $backendRoot
 
    if ($Manual) {
        if ($null -eq $ManualPassword) {
            $ManualPassword = Read-Host `
                "Enter the password for all disposable smoke-test accounts" `
                -AsSecureString
        }
 
        $manualCredential = [pscredential]::new(
            "authentication-smoke",
            $ManualPassword
        )
        $smokePassword = $manualCredential.GetNetworkCredential().Password
 
        if ([string]::IsNullOrWhiteSpace($smokePassword)) {
            throw "The manual smoke-test password cannot be empty."
        }
    }
    else {
        $smokePassword = "Authentication-Smoke-Only-Password-2026!"
    }
 
    $uvCommand = Get-Command uv -ErrorAction Stop
    $pythonOutput = & $uvCommand.Source run --no-sync python -c (
        "import sys; print(sys.executable)"
    )
 
    if ($LASTEXITCODE -ne 0) {
        throw "Could not resolve the project Python interpreter."
    }
 
    $pythonPath = [string]($pythonOutput | Select-Object -Last 1)
    $pythonPath = $pythonPath.Trim()
 
    if (-not (Test-Path -LiteralPath $pythonPath)) {
        throw "Project Python interpreter not found at '$pythonPath'."
    }
 
    $env:DATABASE_URL = "sqlite:///$databaseName"
    $env:SECRET_KEY = (
        "Authentication-Smoke-JWT-Key-Only-For-Disposable-Testing-2026"
    )
    $env:LOG_LEVEL = "WARNING"
    $env:AUTH_SMOKE_PASSWORD = $smokePassword
 
    $seedCode = @'
import json
import os
from datetime import date
 
from sqlmodel import Session
 
from src.auth.bootstrap import create_super_admin
from src.db.connection import engine, init_db
from src.department.repository import create_department
from src.department.schemas import DepartmentCreate
from src.doctor.repository import create_doctor
from src.position.repository import create_position
from src.shift.repository import create_shift, create_shift_assignment
from src.shift.schemas import ShiftCreate, ShiftAssignmentCreate
from src.team.repository import create_team
from src.user import services as user_services
from src.user.models import UserRole
from src.user.schemas import UserAccountCreate
 
 
password = os.environ["AUTH_SMOKE_PASSWORD"]
init_db()
 
with Session(engine) as session:
    department_a = create_department(
        session,
        DepartmentCreate(name="Authentication Department A", code="AUTHA"),
    )
    department_b = create_department(
        session,
        DepartmentCreate(name="Authentication Department B", code="AUTHB"),
    )
 
    team_a = create_team(
        session=session,
        name="Shared Clinical Team",
        department_id=department_a.id,
    )
    team_b = create_team(
        session=session,
        name="Shared Clinical Team",
        department_id=department_b.id,
    )
 
    position_a = create_position(
        session=session,
        position_name="Shared Clinical Position",
        duty_days=[0, 1, 2, 3, 4, 5, 6],
        department_id=department_a.id,
    )
    position_b = create_position(
        session=session,
        position_name="Shared Clinical Position",
        duty_days=[0, 1, 2, 3, 4, 5, 6],
        department_id=department_b.id,
    )
 
    shift_a = create_shift(
        session=session,
        shift_data=ShiftCreate(
            name="Shared Clinical Shift",
            doctors_per_shift=2,
            grants_day_off=False,
            position_id=position_a.id,
        ),
    )
    shift_b = create_shift(
        session=session,
        shift_data=ShiftCreate(
            name="Shared Clinical Shift",
            doctors_per_shift=2,
            grants_day_off=False,
            position_id=position_b.id,
        ),
    )
 
    doctor_a = create_doctor(
        session=session,
        name="Shared Doctor Name",
        team_id=team_a.id,
        department_id=department_a.id,
    )
    doctor_b = create_doctor(
        session=session,
        name="Shared Doctor Name",
        team_id=team_b.id,
        department_id=department_b.id,
    )
 
    assignment_a = create_shift_assignment(
        session=session,
        shift_id=shift_a.id,
        shift_assignment_data=ShiftAssignmentCreate(
            doctor_id=doctor_a.id,
            date=date(2026, 10, 10),
        ),
    )
    assignment_b = create_shift_assignment(
        session=session,
        shift_id=shift_b.id,
        shift_assignment_data=ShiftAssignmentCreate(
            doctor_id=doctor_b.id,
            date=date(2026, 10, 10),
        ),
    )
 
    super_admin = create_super_admin(
        session=session,
        email="auth-smoke-super-admin@example.com",
        full_name="Authentication Smoke Super Admin",
        password=password,
    )
    department_admin_a = user_services.create_user_account(
        session=session,
        account_data=UserAccountCreate(
            email="auth-smoke-admin-a@example.com",
            full_name="Authentication Smoke Admin A",
            password=password,
            role=UserRole.DEPARTMENT_ADMIN,
            department_id=department_a.id,
            doctor_id=None,
        ),
    )
    department_admin_b = user_services.create_user_account(
        session=session,
        account_data=UserAccountCreate(
            email="auth-smoke-admin-b@example.com",
            full_name="Authentication Smoke Admin B",
            password=password,
            role=UserRole.DEPARTMENT_ADMIN,
            department_id=department_b.id,
            doctor_id=None,
        ),
    )
    doctor_user_a = user_services.create_user_account(
        session=session,
        account_data=UserAccountCreate(
            email="auth-smoke-doctor-a@example.com",
            full_name="Authentication Smoke Doctor A",
            password=password,
            role=UserRole.DOCTOR,
            department_id=department_a.id,
            doctor_id=doctor_a.id,
        ),
    )
    viewer_a = user_services.create_user_account(
        session=session,
        account_data=UserAccountCreate(
            email="auth-smoke-viewer-a@example.com",
            full_name="Authentication Smoke Viewer A",
            password=password,
            role=UserRole.VIEWER,
            department_id=department_a.id,
            doctor_id=None,
        ),
    )
 
    print(
        "SMOKE_DATA="
        + json.dumps(
            {
                "department_a_id": str(department_a.id),
                "department_b_id": str(department_b.id),
                "team_a_id": str(team_a.id),
                "team_b_id": str(team_b.id),
                "position_a_id": str(position_a.id),
                "position_b_id": str(position_b.id),
                "shift_a_id": str(shift_a.id),
                "shift_b_id": str(shift_b.id),
                "doctor_a_id": str(doctor_a.id),
                "doctor_b_id": str(doctor_b.id),
                "assignment_a_id": str(assignment_a.id),
                "assignment_b_id": str(assignment_b.id),
                "super_admin_email": super_admin.email,
                "department_admin_a_id": str(department_admin_a.id),
                "department_admin_a_email": department_admin_a.email,
                "department_admin_b_id": str(department_admin_b.id),
                "department_admin_b_email": department_admin_b.email,
                "doctor_user_a_email": doctor_user_a.email,
                "viewer_a_email": viewer_a.email,
            }
        )
    )
'@
 
    $seedOutput = $seedCode | & $pythonPath -
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to seed the disposable database."
    }
 
    $seedLine = $seedOutput |
        Where-Object { $_ -like "SMOKE_DATA=*" } |
        Select-Object -Last 1
 
    if (-not $seedLine) {
        throw "Seeder did not return smoke-test metadata."
    }
 
    $seedData = $seedLine.Substring("SMOKE_DATA=".Length) |
        ConvertFrom-Json
 
    $serverProcess = Start-Process `
        -FilePath $pythonPath `
        -ArgumentList @(
            "-m",
            "uvicorn",
            "src.main:app",
            "--host",
            "127.0.0.1",
            "--port",
            [string]$Port,
            "--log-level",
            "warning"
        ) `
        -WorkingDirectory $backendRoot `
        -RedirectStandardOutput $stdoutPath `
        -RedirectStandardError $stderrPath `
        -PassThru
 
    $serverReady = $false
    for ($attempt = 0; $attempt -lt 50; $attempt++) {
        if ($serverProcess.HasExited) {
            break
        }
 
        try {
            $healthProbe = Invoke-WebRequest `
                -Uri "$baseUri/health" `
                -SkipHttpErrorCheck `
                -TimeoutSec 1
 
            if ($healthProbe.StatusCode -eq 200) {
                $serverReady = $true
                break
            }
        }
        catch {
            # The server may still be starting.
        }
 
        Start-Sleep -Milliseconds 200
    }
 
    if (-not $serverReady) {
        throw "API server did not become ready on $baseUri."
    }
 
    if ($Manual) {
        Show-ManualSession -SeedData $seedData
 
        if ($ManualWaitSeconds -gt 0) {
            Write-Host (
                "Keeping the manual environment open for " +
                "$ManualWaitSeconds seconds."
            )
            Start-Sleep -Seconds $ManualWaitSeconds
        }
        else {
            [void](Read-Host (
                "Press Enter when manual checks are complete to stop the API " +
                "and delete the disposable database"
            ))
        }
 
        Write-Host ""
        Write-Host "Manual authentication test environment closed." `
            -ForegroundColor Green
        return
    }
 
    Write-Pass "two tenants seeded with overlapping resource names"
 
    $healthResponse = Invoke-ApiRequest -Method "GET" -Path "/health"
    Assert-Response `
        -Response $healthResponse `
        -ExpectedStatus 200 `
        -Name "public health endpoint"
 
    $superAdminToken = Get-SmokeAccessToken `
        -Email $seedData.super_admin_email `
        -RoleName "super-admin"
    $departmentAdminAToken = Get-SmokeAccessToken `
        -Email $seedData.department_admin_a_email `
        -RoleName "department-admin A"
    $departmentAdminBToken = Get-SmokeAccessToken `
        -Email $seedData.department_admin_b_email `
        -RoleName "department-admin B"
    $doctorAToken = Get-SmokeAccessToken `
        -Email $seedData.doctor_user_a_email `
        -RoleName "doctor A"
    $viewerAToken = Get-SmokeAccessToken `
        -Email $seedData.viewer_a_email `
        -RoleName "viewer A"
 
    $superAdminHeaders = New-BearerHeaders $superAdminToken
    $departmentAdminAHeaders = New-BearerHeaders $departmentAdminAToken
    $departmentAdminBHeaders = New-BearerHeaders $departmentAdminBToken
    $doctorAHeaders = New-BearerHeaders $doctorAToken
    $viewerAHeaders = New-BearerHeaders $viewerAToken
 
    $profileResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/auth/me" `
        -Headers $superAdminHeaders
    Assert-Response `
        -Response $profileResponse `
        -ExpectedStatus 200 `
        -Name "authenticated profile"
 
    $anonymousChecks = @(
        @{
            Name = "anonymous auth profile"
            Method = "GET"
            Path = "/api/v1/auth/me"
        },
        @{
            Name = "anonymous user list"
            Method = "GET"
            Path = "/api/v1/users/"
        },
        @{
            Name = "anonymous department list"
            Method = "GET"
            Path = "/api/v1/departments/"
        },
        @{
            Name = "anonymous department detail"
            Method = "GET"
            Path = "/api/v1/departments/$($seedData.department_a_id)"
        },
        @{
            Name = "anonymous team create"
            Method = "POST"
            Path = "/api/v1/teams/"
            Body = @{
                name = "Anonymous Denied Team"
            }
        },
        @{
            Name = "anonymous team list"
            Method = "GET"
            Path = "/api/v1/teams/"
        },
        @{
            Name = "anonymous team detail"
            Method = "GET"
            Path = "/api/v1/teams/$($seedData.team_a_id)"
        },
        @{
            Name = "anonymous position create"
            Method = "POST"
            Path = "/api/v1/positions/"
            Body = @{
                name = "Anonymous Denied Position"
                duty_days = @(0, 1, 2)
            }
        },
        @{
            Name = "anonymous position list"
            Method = "GET"
            Path = "/api/v1/positions/"
        },
        @{
            Name = "anonymous position detail"
            Method = "GET"
            Path = "/api/v1/positions/$($seedData.position_a_id)"
        },
        @{
            Name = "anonymous shift create"
            Method = "POST"
            Path = "/api/v1/shifts/"
            Body = @{
                name = "Anonymous Denied Shift"
                doctors_per_shift = 1
                grants_day_off = $false
                position_id = $seedData.position_a_id
            }
        },
        @{
            Name = "anonymous shift list"
            Method = "GET"
            Path = "/api/v1/shifts/"
        },
        @{
            Name = "anonymous shift detail"
            Method = "GET"
            Path = "/api/v1/shifts/$($seedData.shift_a_id)"
        },
        @{
            Name = "anonymous shift-assignment list"
            Method = "GET"
            Path = "/api/v1/shifts/assignments"
        },
        @{
            Name = "anonymous shift-assignment create"
            Method = "POST"
            Path = "/api/v1/shifts/$($seedData.shift_a_id)/assignments"
            Body = @{
                doctor_id = $seedData.doctor_a_id
                date = "2026-10-11"
            }
        },
        @{
            Name = "anonymous doctor create"
            Method = "POST"
            Path = "/api/v1/doctors/"
            Body = @{
                name = "Anonymous Denied Doctor"
                team_id = $seedData.team_a_id
            }
        },
        @{
            Name = "anonymous doctor list"
            Method = "GET"
            Path = "/api/v1/doctors/"
        },
        @{
            Name = "anonymous doctor roster"
            Method = "GET"
            Path = "/api/v1/doctors/roster"
        },
        @{
            Name = "anonymous doctor detail"
            Method = "GET"
            Path = "/api/v1/doctors/$($seedData.doctor_a_id)"
        },
        @{
            Name = "anonymous pre-assignment create"
            Method = "POST"
            Path = (
                "/api/v1/doctors/$($seedData.doctor_a_id)/pre-assignments"
            )
            Body = @{
                shift_id = $seedData.shift_a_id
                date = "2026-10-12"
            }
        },
        @{
            Name = "anonymous pre-assignment list"
            Method = "GET"
            Path = (
                "/api/v1/doctors/$($seedData.doctor_a_id)/pre-assignments"
            )
        },
        @{
            Name = "anonymous unavailability create"
            Method = "POST"
            Path = (
                "/api/v1/doctors/$($seedData.doctor_a_id)/unavailability"
            )
            Body = @{
                date = "2026-10-01"
            }
        },
        @{
            Name = "anonymous unavailability list"
            Method = "GET"
            Path = (
                "/api/v1/doctors/$($seedData.doctor_a_id)/unavailability"
            )
        },
        @{
            Name = "anonymous doctor-position create"
            Method = "POST"
            Path = "/api/v1/doctors/$($seedData.doctor_a_id)/position"
            Body = @{
                position_id = $seedData.position_a_id
            }
        },
        @{
            Name = "anonymous doctor-position list"
            Method = "GET"
            Path = "/api/v1/doctors/$($seedData.doctor_a_id)/position"
        }
    )
 
    foreach ($check in $anonymousChecks) {
        if ($check.ContainsKey("Body")) {
            $response = Invoke-ApiRequest `
                -Method $check.Method `
                -Path $check.Path `
                -Body $check.Body
        }
        else {
            $response = Invoke-ApiRequest `
                -Method $check.Method `
                -Path $check.Path
        }
 
        Assert-Response `
            -Response $response `
            -ExpectedStatus 401 `
            -ExpectedDetail "Unauthorized" `
            -ExpectBearerHeader `
            -Name $check.Name
    }
 
    $deniedTeamNames = @(
        "Doctor Denied Team",
        "Viewer Denied Team",
        "Super Admin Denied Team"
    )
    $roleChecks = @(
        @{
            Name = "doctor tenant write"
            Headers = $doctorAHeaders
            TeamName = $deniedTeamNames[0]
        },
        @{
            Name = "viewer tenant write"
            Headers = $viewerAHeaders
            TeamName = $deniedTeamNames[1]
        },
        @{
            Name = "super-admin tenant write"
            Headers = $superAdminHeaders
            TeamName = $deniedTeamNames[2]
        }
    )
 
    foreach ($roleCheck in $roleChecks) {
        $response = Invoke-ApiRequest `
            -Method "POST" `
            -Path "/api/v1/teams/" `
            -Headers $roleCheck.Headers `
            -Body @{
                name = $roleCheck.TeamName
            }
 
        Assert-Response `
            -Response $response `
            -ExpectedStatus 403 `
            -ExpectedDetail "Insufficient permissions for this operation." `
            -ExpectNoBearerHeader `
            -Name $roleCheck.Name
    }
 
    $departmentAdminGlobalResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/departments/" `
        -Headers $departmentAdminAHeaders
    Assert-Response `
        -Response $departmentAdminGlobalResponse `
        -ExpectedStatus 403 `
        -ExpectedDetail "Insufficient permissions for this operation." `
        -ExpectNoBearerHeader `
        -Name "department-admin global department read"
 
    $viewerUnavailabilityResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path (
            "/api/v1/doctors/$($seedData.doctor_a_id)/unavailability"
        ) `
        -Headers $viewerAHeaders
    Assert-Response `
        -Response $viewerUnavailabilityResponse `
        -ExpectedStatus 403 `
        -ExpectedDetail "Insufficient permissions for this operation." `
        -ExpectNoBearerHeader `
        -Name "viewer raw unavailability read"
 
    $doctorOwnUnavailabilityResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path (
            "/api/v1/doctors/$($seedData.doctor_a_id)/unavailability"
        ) `
        -Headers $doctorAHeaders `
        -Body @{
            date = "2026-10-01"
        }
    Assert-Response `
        -Response $doctorOwnUnavailabilityResponse `
        -ExpectedStatus 201 `
        -Name "doctor own unavailability write"
 
    $doctorForeignUnavailabilityRead = Invoke-ApiRequest `
        -Method "GET" `
        -Path (
            "/api/v1/doctors/$($seedData.doctor_b_id)/unavailability"
        ) `
        -Headers $doctorAHeaders
    Assert-Response `
        -Response $doctorForeignUnavailabilityRead `
        -ExpectedStatus 403 `
        -ExpectedDetail "Cannot access another doctor's unavailability." `
        -ExpectNoBearerHeader `
        -Name "doctor foreign unavailability read"
 
    $doctorForeignUnavailabilityWrite = Invoke-ApiRequest `
        -Method "POST" `
        -Path (
            "/api/v1/doctors/$($seedData.doctor_b_id)/unavailability"
        ) `
        -Headers $doctorAHeaders `
        -Body @{
            date = "2026-10-02"
        }
    Assert-Response `
        -Response $doctorForeignUnavailabilityWrite `
        -ExpectedStatus 403 `
        -ExpectedDetail "Cannot access another doctor's unavailability." `
        -ExpectNoBearerHeader `
        -Name "doctor foreign unavailability write"
 
    $adminForeignUnavailabilityRead = Invoke-ApiRequest `
        -Method "GET" `
        -Path (
            "/api/v1/doctors/$($seedData.doctor_b_id)/unavailability"
        ) `
        -Headers $departmentAdminAHeaders
    Assert-Response `
        -Response $adminForeignUnavailabilityRead `
        -ExpectedStatus 404 `
        -ExpectNoBearerHeader `
        -Name "department-admin foreign unavailability read"
 
    $suppliedTeamDepartmentResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/teams/" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            name = "Supplied Department Team"
            department_id = $seedData.department_b_id
        }
    Assert-Response `
        -Response $suppliedTeamDepartmentResponse `
        -ExpectedStatus 422 `
        -Name "team create rejects supplied department"
 
    $suppliedPositionDepartmentResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/positions/" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            name = "Supplied Department Position"
            duty_days = @(0, 2, 4)
            department_id = $seedData.department_b_id
        }
    Assert-Response `
        -Response $suppliedPositionDepartmentResponse `
        -ExpectedStatus 422 `
        -Name "position create rejects supplied department"
 
    $suppliedDoctorDepartmentResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/doctors/" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            name = "Supplied Department Doctor"
            team_id = $seedData.team_a_id
            department_id = $seedData.department_b_id
        }
    Assert-Response `
        -Response $suppliedDoctorDepartmentResponse `
        -ExpectedStatus 422 `
        -Name "doctor create rejects supplied department"
 
    $derivedTeamResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/teams/" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            name = "Derived Department Team"
        }
    Assert-Response `
        -Response $derivedTeamResponse `
        -ExpectedStatus 201 `
        -Name "department-derived team create"
    Assert-JsonProperty `
        -Response $derivedTeamResponse `
        -PropertyName "department_id" `
        -ExpectedValue ([string]$seedData.department_a_id) `
        -Name "team stores authenticated department"
    $derivedTeam = $derivedTeamResponse.Content | ConvertFrom-Json
 
    $derivedPositionResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/positions/" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            name = "Derived Department Position"
            duty_days = @(1, 3, 5)
        }
    Assert-Response `
        -Response $derivedPositionResponse `
        -ExpectedStatus 201 `
        -Name "department-derived position create"
    Assert-JsonProperty `
        -Response $derivedPositionResponse `
        -PropertyName "department_id" `
        -ExpectedValue ([string]$seedData.department_a_id) `
        -Name "position stores authenticated department"
    $derivedPosition = $derivedPositionResponse.Content | ConvertFrom-Json
 
    $derivedDoctorResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/doctors/" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            name = "Derived Department Doctor"
            team_id = [string]$derivedTeam.id
        }
    Assert-Response `
        -Response $derivedDoctorResponse `
        -ExpectedStatus 201 `
        -Name "department-derived doctor create"
    Assert-JsonProperty `
        -Response $derivedDoctorResponse `
        -PropertyName "department_id" `
        -ExpectedValue ([string]$seedData.department_a_id) `
        -Name "doctor stores authenticated department"
    $derivedDoctor = $derivedDoctorResponse.Content | ConvertFrom-Json
 
    $derivedShiftResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/shifts/" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            name = "Derived Department Shift"
            doctors_per_shift = 1
            grants_day_off = $false
            position_id = [string]$derivedPosition.id
        }
    Assert-Response `
        -Response $derivedShiftResponse `
        -ExpectedStatus 201 `
        -Name "position-derived shift create"
    Assert-JsonProperty `
        -Response $derivedShiftResponse `
        -PropertyName "position_id" `
        -ExpectedValue ([string]$derivedPosition.id) `
        -Name "shift stores scoped Position"
    $derivedShift = $derivedShiftResponse.Content | ConvertFrom-Json
 
    $foreignShiftCreate = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/shifts/" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            name = "Cross Tenant Denied Shift"
            doctors_per_shift = 1
            grants_day_off = $false
            position_id = $seedData.position_b_id
        }
    Assert-Response `
        -Response $foreignShiftCreate `
        -ExpectedStatus 404 `
        -ExpectNoBearerHeader `
        -Name "foreign Position shift create"
 
    $foreignDoctorCreate = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/doctors/" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            name = "Cross Tenant Denied Doctor"
            team_id = $seedData.team_b_id
        }
    Assert-Response `
        -Response $foreignDoctorCreate `
        -ExpectedStatus 404 `
        -ExpectNoBearerHeader `
        -Name "foreign Team doctor create"
 
    $foreignShiftAssignment = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/shifts/$($seedData.shift_b_id)/assignments" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            doctor_id = $seedData.doctor_a_id
            date = "2026-10-11"
        }
    Assert-Response `
        -Response $foreignShiftAssignment `
        -ExpectedStatus 404 `
        -ExpectNoBearerHeader `
        -Name "foreign Shift assignment create"
 
    $foreignDoctorAssignment = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/shifts/$($seedData.shift_a_id)/assignments" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            doctor_id = $seedData.doctor_b_id
            date = "2026-10-11"
        }
    Assert-Response `
        -Response $foreignDoctorAssignment `
        -ExpectedStatus 404 `
        -ExpectNoBearerHeader `
        -Name "foreign Doctor assignment create"
 
    if ($foreignShiftAssignment.Content -ne $foreignDoctorAssignment.Content) {
        throw "Foreign Shift and Doctor assignment responses are distinguishable."
    }
    Write-Pass "assignment foreign-resource response is indistinguishable"
 
    $foreignPreAssignment = Invoke-ApiRequest `
        -Method "POST" `
        -Path (
            "/api/v1/doctors/$($seedData.doctor_a_id)/pre-assignments"
        ) `
        -Headers $departmentAdminAHeaders `
        -Body @{
            shift_id = $seedData.shift_b_id
            date = "2026-10-12"
        }
    Assert-Response `
        -Response $foreignPreAssignment `
        -ExpectedStatus 404 `
        -ExpectNoBearerHeader `
        -Name "foreign Shift pre-assignment create"
 
    $foreignDoctorPosition = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/doctors/$($seedData.doctor_a_id)/position" `
        -Headers $departmentAdminAHeaders `
        -Body @{
            position_id = $seedData.position_b_id
        }
    Assert-Response `
        -Response $foreignDoctorPosition `
        -ExpectedStatus 404 `
        -ExpectNoBearerHeader `
        -Name "foreign Position doctor association create"
 
    $globalDepartmentsResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/departments/" `
        -Headers $superAdminHeaders
    Assert-ListIds `
        -Response $globalDepartmentsResponse `
        -Name "super-admin global department list" `
        -ExpectedIds @(
            [string]$seedData.department_a_id,
            [string]$seedData.department_b_id
        ) `
        -ExpectedCount 2
 
    $userListAResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/users/" `
        -Headers $departmentAdminAHeaders
    Assert-ListIds `
        -Response $userListAResponse `
        -Name "Department A user list" `
        -ExpectedIds @([string]$seedData.department_admin_a_id) `
        -ExcludedIds @([string]$seedData.department_admin_b_id) `
        -ExpectedCount 3
 
    $doctorListAResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/doctors/" `
        -Headers $departmentAdminAHeaders
    Assert-ListIds `
        -Response $doctorListAResponse `
        -Name "Department A doctor list" `
        -ExpectedIds @(
            [string]$seedData.doctor_a_id,
            [string]$derivedDoctor.id
        ) `
        -ExcludedIds @([string]$seedData.doctor_b_id) `
        -ExpectedCount 2
 
    $rosterAResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/doctors/roster" `
        -Headers $viewerAHeaders
    Assert-ListIds `
        -Response $rosterAResponse `
        -Name "Department A reduced doctor roster" `
        -ExpectedIds @(
            [string]$seedData.doctor_a_id,
            [string]$derivedDoctor.id
        ) `
        -ExcludedIds @([string]$seedData.doctor_b_id) `
        -ExpectedCount 2
    Assert-ListOmitsProperty `
        -Response $rosterAResponse `
        -PropertyName "email" `
        -Name "reduced doctor roster omits email"
 
    $teamListAResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/teams/" `
        -Headers $departmentAdminAHeaders
    Assert-ListIds `
        -Response $teamListAResponse `
        -Name "Department A team list" `
        -ExpectedIds @(
            [string]$seedData.team_a_id,
            [string]$derivedTeam.id
        ) `
        -ExcludedIds @([string]$seedData.team_b_id) `
        -ExpectedCount 2
 
    $positionListAResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/positions/" `
        -Headers $departmentAdminAHeaders
    Assert-ListIds `
        -Response $positionListAResponse `
        -Name "Department A position list" `
        -ExpectedIds @(
            [string]$seedData.position_a_id,
            [string]$derivedPosition.id
        ) `
        -ExcludedIds @([string]$seedData.position_b_id) `
        -ExpectedCount 2
 
    $shiftListAResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/shifts/" `
        -Headers $departmentAdminAHeaders
    Assert-ListIds `
        -Response $shiftListAResponse `
        -Name "Department A shift list" `
        -ExpectedIds @(
            [string]$seedData.shift_a_id,
            [string]$derivedShift.id
        ) `
        -ExcludedIds @([string]$seedData.shift_b_id) `
        -ExpectedCount 2
 
    $assignmentListAResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/shifts/assignments" `
        -Headers $viewerAHeaders
    Assert-ListIds `
        -Response $assignmentListAResponse `
        -Name "Department A assignment list" `
        -ExpectedIds @([string]$seedData.assignment_a_id) `
        -ExcludedIds @([string]$seedData.assignment_b_id) `
        -ExpectedCount 1
 
    $ownDetailChecks = @(
        @{
            Name = "Department A department UUID detail"
            Path = "/api/v1/departments/$($seedData.department_a_id)"
            Headers = $viewerAHeaders
            ExpectedId = [string]$seedData.department_a_id
        },
        @{
            Name = "Department A team UUID detail"
            Path = "/api/v1/teams/$($seedData.team_a_id)"
            Headers = $viewerAHeaders
            ExpectedId = [string]$seedData.team_a_id
        },
        @{
            Name = "Department A position UUID detail"
            Path = "/api/v1/positions/$($seedData.position_a_id)"
            Headers = $viewerAHeaders
            ExpectedId = [string]$seedData.position_a_id
        },
        @{
            Name = "Department A shift UUID detail"
            Path = "/api/v1/shifts/$($seedData.shift_a_id)"
            Headers = $viewerAHeaders
            ExpectedId = [string]$seedData.shift_a_id
        },
        @{
            Name = "Department A doctor UUID detail"
            Path = "/api/v1/doctors/$($seedData.doctor_a_id)"
            Headers = $departmentAdminAHeaders
            ExpectedId = [string]$seedData.doctor_a_id
        }
    )
 
    foreach ($detailCheck in $ownDetailChecks) {
        $response = Invoke-ApiRequest `
            -Method "GET" `
            -Path $detailCheck.Path `
            -Headers $detailCheck.Headers
        Assert-Response `
            -Response $response `
            -ExpectedStatus 200 `
            -Name $detailCheck.Name
        Assert-JsonProperty `
            -Response $response `
            -PropertyName "id" `
            -ExpectedValue $detailCheck.ExpectedId `
            -Name "$($detailCheck.Name) identity"
    }
 
    $doctorDetailAResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/doctors/$($seedData.doctor_a_id)" `
        -Headers $departmentAdminAHeaders
    Assert-JsonProperty `
        -Response $doctorDetailAResponse `
        -PropertyName "name" `
        -ExpectedValue "Shared Doctor Name" `
        -Name "department-admin full doctor response includes name"
 
    $foreignDetailChecks = @(
        @{
            Name = "foreign department detail"
            Path = "/api/v1/departments/$($seedData.department_b_id)"
            Headers = $viewerAHeaders
        },
        @{
            Name = "foreign team detail"
            Path = "/api/v1/teams/$($seedData.team_b_id)"
            Headers = $viewerAHeaders
        },
        @{
            Name = "foreign position detail"
            Path = "/api/v1/positions/$($seedData.position_b_id)"
            Headers = $viewerAHeaders
        },
        @{
            Name = "foreign shift detail"
            Path = "/api/v1/shifts/$($seedData.shift_b_id)"
            Headers = $viewerAHeaders
        },
        @{
            Name = "foreign doctor detail"
            Path = "/api/v1/doctors/$($seedData.doctor_b_id)"
            Headers = $departmentAdminAHeaders
        }
    )
 
    foreach ($detailCheck in $foreignDetailChecks) {
        $response = Invoke-ApiRequest `
            -Method "GET" `
            -Path $detailCheck.Path `
            -Headers $detailCheck.Headers
        Assert-Response `
            -Response $response `
            -ExpectedStatus 404 `
            -ExpectNoBearerHeader `
            -Name $detailCheck.Name
    }
 
    $teamListBResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/teams/" `
        -Headers $departmentAdminBHeaders
    Assert-ListIds `
        -Response $teamListBResponse `
        -Name "Department B team non-mutation" `
        -ExpectedIds @([string]$seedData.team_b_id) `
        -ExcludedIds @([string]$derivedTeam.id) `
        -ExpectedCount 1
 
    $positionListBResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/positions/" `
        -Headers $departmentAdminBHeaders
    Assert-ListIds `
        -Response $positionListBResponse `
        -Name "Department B position non-mutation" `
        -ExpectedIds @([string]$seedData.position_b_id) `
        -ExcludedIds @([string]$derivedPosition.id) `
        -ExpectedCount 1
 
    $shiftListBResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/shifts/" `
        -Headers $departmentAdminBHeaders
    Assert-ListIds `
        -Response $shiftListBResponse `
        -Name "Department B shift non-mutation" `
        -ExpectedIds @([string]$seedData.shift_b_id) `
        -ExcludedIds @([string]$derivedShift.id) `
        -ExpectedCount 1
 
    $doctorListBResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/doctors/" `
        -Headers $departmentAdminBHeaders
    Assert-ListIds `
        -Response $doctorListBResponse `
        -Name "Department B doctor non-mutation" `
        -ExpectedIds @([string]$seedData.doctor_b_id) `
        -ExcludedIds @([string]$derivedDoctor.id) `
        -ExpectedCount 1
 
    $assignmentListBResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/shifts/assignments" `
        -Headers $departmentAdminBHeaders
    Assert-ListIds `
        -Response $assignmentListBResponse `
        -Name "Department B assignment non-mutation" `
        -ExpectedIds @([string]$seedData.assignment_b_id) `
        -ExcludedIds @([string]$seedData.assignment_a_id) `
        -ExpectedCount 1
 
    $preAssignmentListAResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path (
            "/api/v1/doctors/$($seedData.doctor_a_id)/pre-assignments"
        ) `
        -Headers $departmentAdminAHeaders
    Assert-ListIds `
        -Response $preAssignmentListAResponse `
        -Name "foreign pre-assignment write non-mutation" `
        -ExpectedCount 0
 
    $doctorPositionListAResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/doctors/$($seedData.doctor_a_id)/position" `
        -Headers $departmentAdminAHeaders
    Assert-ListIds `
        -Response $doctorPositionListAResponse `
        -Name "foreign Doctor-Position write non-mutation" `
        -ExpectedCount 0
 
    $doctorBUnavailabilityResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path (
            "/api/v1/doctors/$($seedData.doctor_b_id)/unavailability"
        ) `
        -Headers $departmentAdminBHeaders
    Assert-ListIds `
        -Response $doctorBUnavailabilityResponse `
        -Name "foreign unavailability write non-mutation" `
        -ExpectedCount 0
 
    $teamListAItems = @($teamListAResponse.Content | ConvertFrom-Json)
    $teamNames = @($teamListAItems | ForEach-Object { [string]$_.name })
    foreach ($deniedTeamName in $deniedTeamNames) {
        if ($deniedTeamName -in $teamNames) {
            throw "Denied team '$deniedTeamName' was persisted."
        }
    }
    Write-Pass "role-denied team writes were not persisted"
 
    $signupResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/users/signup" `
        -Body @{
            email = "removed-signup@example.com"
            full_name = "Removed Signup"
            password = "NotUsed-Password-2026!"
        }
    Assert-Response `
        -Response $signupResponse `
        -ExpectedStatus 404 `
        -Name "removed public signup"
 
    $userLookupResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/users/removed-user%40example.com"
    Assert-Response `
        -Response $userLookupResponse `
        -ExpectedStatus 404 `
        -Name "removed public user lookup"
 
    $departmentCreateResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/departments/" `
        -Body @{
            name = "Removed Department Creation"
            code = "REMOVED"
        }
    Assert-Response `
        -Response $departmentCreateResponse `
        -ExpectedStatus 405 `
        -Name "removed public department creation"

    # ==========================================================================
    # Phase 7: Scoped Invitations, Department Provisioning & Public Signup Flow
    # ==========================================================================

    # 1. Super Admin provisions a new department and receives the admin invitation
    $provisionResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/invitations/provision" `
        -Headers $superAdminHeaders `
        -Body @{
            department_name = "Phase 7 Neurology Department"
            department_code = "P7-NEURO"
        }
    Assert-Response `
        -Response $provisionResponse `
        -ExpectedStatus 201 `
        -Name "super-admin provision department and admin invitation"
    $provisionData = $provisionResponse.Content | ConvertFrom-Json
    $p7DeptId = [string]$provisionData.department.id
    $p7AdminToken = [string]$provisionData.invitation.raw_token

    if ([string]::IsNullOrWhiteSpace($p7AdminToken)) {
        throw "Department provisioning did not return a raw invitation token."
    }

    # 2. Department Admin registers using the invitation token
    $p7AdminEmail = "phase7-admin-neuro@example.com"
    $adminSignupResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/signup" `
        -Body @{
            invitation_token = $p7AdminToken
            first_name = "Gregory"
            last_name = "House"
            email = $p7AdminEmail
            password = $script:smokePassword
        }
    Assert-Response `
        -Response $adminSignupResponse `
        -ExpectedStatus 201 `
        -Name "department-admin public signup via invitation"
    Assert-JsonProperty `
        -Response $adminSignupResponse `
        -PropertyName "role" `
        -ExpectedValue "department_admin" `
        -Name "invited admin has department_admin role"
    Assert-JsonProperty `
        -Response $adminSignupResponse `
        -PropertyName "department_id" `
        -ExpectedValue $p7DeptId `
        -Name "invited admin assigned to provisioned department"

    # 3. Newly registered Department Admin logs in
    $p7AdminAccessToken = Get-SmokeAccessToken `
        -Email $p7AdminEmail `
        -RoleName "newly registered department-admin"
    $p7AdminHeaders = New-BearerHeaders $p7AdminAccessToken

    # 4. Department Admin issues a staff Doctor invitation (open, auto-provisioning)
    $doctorInviteResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/invitations/staff" `
        -Headers $p7AdminHeaders `
        -Body @{
            role = "doctor"
            doctor_id = $null
        }
    Assert-Response `
        -Response $doctorInviteResponse `
        -ExpectedStatus 201 `
        -Name "department-admin creates doctor staff invitation"
    $doctorInviteData = $doctorInviteResponse.Content | ConvertFrom-Json
    $p7DoctorToken = [string]$doctorInviteData.raw_token

    # 5. Doctor registers using the invitation token (auto-provisions doctor)
    $p7DoctorEmail = "phase7-doctor-cameron@example.com"
    $doctorSignupResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/signup" `
        -Body @{
            invitation_token = $p7DoctorToken
            first_name = "Allison"
            last_name = "Cameron"
            email = $p7DoctorEmail
            password = $script:smokePassword
        }
    Assert-Response `
        -Response $doctorSignupResponse `
        -ExpectedStatus 201 `
        -Name "doctor public signup auto-provisions doctor record"
    Assert-JsonProperty `
        -Response $doctorSignupResponse `
        -PropertyName "role" `
        -ExpectedValue "doctor" `
        -Name "invited doctor has doctor role"
    Assert-JsonProperty `
        -Response $doctorSignupResponse `
        -PropertyName "department_id" `
        -ExpectedValue $p7DeptId `
        -Name "invited doctor assigned to correct department"

    $doctorUserData = $doctorSignupResponse.Content | ConvertFrom-Json
    if ([string]::IsNullOrWhiteSpace([string]$doctorUserData.doctor_id)) {
        throw "Doctor signup did not auto-provision and attach doctor_id."
    }

    # 6. Newly registered Doctor logs in and verifies profile
    $p7DoctorAccessToken = Get-SmokeAccessToken `
        -Email $p7DoctorEmail `
        -RoleName "newly registered doctor"
    $p7DoctorHeaders = New-BearerHeaders $p7DoctorAccessToken

    $p7DoctorProfile = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/auth/me" `
        -Headers $p7DoctorHeaders
    Assert-Response `
        -Response $p7DoctorProfile `
        -ExpectedStatus 200 `
        -Name "newly registered doctor profile check"

    # 7. Negative test: Consumed doctor invitation token cannot be reused
    $reusedSignupResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/signup" `
        -Body @{
            invitation_token = $p7DoctorToken
            first_name = "Imposter"
            last_name = "Doctor"
            email = "imposter@example.com"
            password = $script:smokePassword
        }
    Assert-Response `
        -Response $reusedSignupResponse `
        -ExpectedStatus 400 `
        -Name "signup rejects already consumed invitation token"

    # 8. Negative test: Revoked invitation cannot be consumed
    $revokableInviteResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/invitations/staff" `
        -Headers $p7AdminHeaders `
        -Body @{
            role = "viewer"
        }
    Assert-Response `
        -Response $revokableInviteResponse `
        -ExpectedStatus 201 `
        -Name "department-admin creates revokable viewer invitation"
    $revokableData = $revokableInviteResponse.Content | ConvertFrom-Json
    $revokableId = [string]$revokableData.id
    $revokableToken = [string]$revokableData.raw_token

    $revokeResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/invitations/$revokableId/revoke" `
        -Headers $p7AdminHeaders
    Assert-Response `
        -Response $revokeResponse `
        -ExpectedStatus 200 `
        -Name "department-admin revokes invitation"

    $revokedSignupResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/signup" `
        -Body @{
            invitation_token = $revokableToken
            first_name = "Revoked"
            last_name = "User"
            email = "revoked@example.com"
            password = $script:smokePassword
        }
    Assert-Response `
        -Response $revokedSignupResponse `
        -ExpectedStatus 400 `
        -Name "signup rejects revoked invitation token"

    # 9. Scoped listing check: Department admin sees invitations
    $invitationsListResponse = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/auth/invitations" `
        -Headers $p7AdminHeaders
    Assert-Response `
        -Response $invitationsListResponse `
        -ExpectedStatus 200 `
        -Name "department-admin lists department invitations"

    # 10. Phase 8: Refresh-Session Persistence & Token Rotation Verification
    Write-Host "Verifying Phase 8 Refresh Session & Rotation Lifecycle..." -ForegroundColor Cyan

    # 10.1 Login to obtain access and refresh tokens
    $p8LoginResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/login" `
        -ContentType "application/x-www-form-urlencoded" `
        -Body @{
            username = $seedData.super_admin_email
            password = $script:smokePassword
        }
    Assert-Response `
        -Response $p8LoginResponse `
        -ExpectedStatus 200 `
        -Name "Phase 8 login returns session credentials"

    $p8LoginData = $p8LoginResponse.Content | ConvertFrom-Json
    $rawRefresh1 = [string]$p8LoginData.refresh_token
    if ([string]::IsNullOrWhiteSpace($rawRefresh1)) {
        throw "Phase 8 login did not return a refresh_token."
    }
    Write-Pass "Phase 8 login payload contains refresh_token"

    # 10.2 Token rotation: POST /api/v1/auth/refresh with T_1
    $rotateResponse1 = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/refresh" `
        -Body @{
            refresh_token = $rawRefresh1
        }
    Assert-Response `
        -Response $rotateResponse1 `
        -ExpectedStatus 200 `
        -Name "refresh endpoint rotates T_1 to T_2"

    $rotateData1 = $rotateResponse1.Content | ConvertFrom-Json
    $rawRefresh2 = [string]$rotateData1.refresh_token
    $rotatedAccess = [string]$rotateData1.access_token

    if ([string]::IsNullOrWhiteSpace($rawRefresh2) -or $rawRefresh2 -eq $rawRefresh1) {
        throw "Phase 8 refresh did not return a unique rotated refresh_token."
    }
    if ([string]::IsNullOrWhiteSpace($rotatedAccess)) {
        throw "Phase 8 refresh did not return an access_token."
    }
    Write-Pass "refresh endpoint issued fresh access token and rotated refresh token"

    # 10.3 Verify rotated access token works against protected profile endpoint
    $testHeaders = New-BearerHeaders -Token $rotatedAccess
    $profileResp = Invoke-ApiRequest `
        -Method "GET" `
        -Path "/api/v1/auth/me" `
        -Headers $testHeaders
    Assert-Response `
        -Response $profileResp `
        -ExpectedStatus 200 `
        -Name "new access token queries protected endpoint"

    # 10.4 Reuse Detection: Attacker presents stale T_1 -> rejected with 401
    $replayResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/refresh" `
        -Body @{
            refresh_token = $rawRefresh1
        }
    Assert-Response `
        -Response $replayResponse `
        -ExpectedStatus 401 `
        -Name "replay of stale T_1 is rejected with 401 reuse detected"

    # 10.5 Session Family Invalidation: Legitimate client presents T_2 -> rejected with 401
    $nukedFamilyResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/refresh" `
        -Body @{
            refresh_token = $rawRefresh2
        }
    Assert-Response `
        -Response $nukedFamilyResponse `
        -ExpectedStatus 401 `
        -Name "reuse detection revokes entire session family (T_2 rejected with 401)"

    # 10.6 Logout Flow: Login fresh session, call /logout, verify termination
    $p8Login2Response = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/login" `
        -ContentType "application/x-www-form-urlencoded" `
        -Body @{
            username = $seedData.super_admin_email
            password = $script:smokePassword
        }
    Assert-Response `
        -Response $p8Login2Response `
        -ExpectedStatus 200 `
        -Name "login for logout verification"
    $p8Login2Data = $p8Login2Response.Content | ConvertFrom-Json
    $logoutToken = [string]$p8Login2Data.refresh_token

    $logoutResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/logout" `
        -Body @{
            refresh_token = $logoutToken
        }
    Assert-Response `
        -Response $logoutResponse `
        -ExpectedStatus 200 `
        -Name "logout terminates active refresh session"

    # Verify post-logout token replay is rejected
    $postLogoutResponse = Invoke-ApiRequest `
        -Method "POST" `
        -Path "/api/v1/auth/refresh" `
        -Body @{
            refresh_token = $logoutToken
        }
    Assert-Response `
        -Response $postLogoutResponse `
        -ExpectedStatus 401 `
        -Name "refresh with logged out token is rejected with 401"

    Write-Host "Verifying Phase 11 Endpoint Rate Limiting..."
    $throttledIp = "198.51.100.42"
    $cleanIp = "198.51.100.43"
    $rateLimitHeaders = @{
        "X-Forwarded-For" = $throttledIp
    }

    # Make 5 signup attempts to reach the limit
    for ($i = 1; $i -le 5; $i++) {
        $attemptResp = Invoke-ApiRequest `
            -Method POST `
            -Path "/api/v1/auth/signup" `
            -Headers $rateLimitHeaders `
            -Body @{
                invitation_token = "invalid_token_$i"
                first_name = "Rate"
                last_name = "Limit"
                email = "rate_limit_$i@hospital.org"
                password = "SecurePassword123!"
            }
        Assert-Response `
            -Response $attemptResp `
            -ExpectedStatus 404 `
            -Name "rate limit warmup signup attempt $i"
    }

    # 6th request from throttled IP must be 429 Too Many Requests
    $throttledResp = Invoke-ApiRequest `
        -Method POST `
        -Path "/api/v1/auth/signup" `
        -Headers $rateLimitHeaders `
        -Body @{
            invitation_token = "invalid_token_6"
            first_name = "Rate"
            last_name = "Limit"
            email = "rate_limit_6@hospital.org"
            password = "SecurePassword123!"
        }
    Assert-Response `
        -Response $throttledResp `
        -ExpectedStatus 429 `
        -Name "signup rate limit triggers HTTP 429"

    # Verify Retry-After header
    $retryAfter = [string]$throttledResp.Headers["Retry-After"]
    if (-not $retryAfter -or [int]$retryAfter -le 0) {
        throw "Rate limited response missing valid Retry-After header. Found: '$retryAfter'"
    }
    Write-Pass -Name "rate limited response includes positive Retry-After header"

    # Verify an independent clean IP is not throttled
    $unthrottledResp = Invoke-ApiRequest `
        -Method POST `
        -Path "/api/v1/auth/signup" `
        -Headers @{ "X-Forwarded-For" = $cleanIp } `
        -Body @{
            invitation_token = "invalid_token_clean"
            first_name = "Rate"
            last_name = "Limit"
            email = "rate_limit_clean@hospital.org"
            password = "SecurePassword123!"
        }
    Assert-Response `
        -Response $unthrottledResp `
        -ExpectedStatus 404 `
        -Name "unthrottled IP is isolated and receives standard 404"


    Write-Host ""
    Write-Host (
        "Authentication smoke test passed: $passedChecks checks."
    ) -ForegroundColor Green

}
catch {
    $exitCode = 1
    Write-Host ""
    Write-Host "[FAIL] $($_.Exception.Message)" -ForegroundColor Red
 
    if (Test-Path -LiteralPath $stderrPath) {
        $serverErrors = Get-Content -LiteralPath $stderrPath -Tail 20
        if ($serverErrors) {
            Write-Host "--- Server stderr ---" -ForegroundColor Yellow
            $serverErrors | ForEach-Object { Write-Host $_ }
        }
    }
}
finally {
    if ($null -ne $serverProcess -and -not $serverProcess.HasExited) {
        Stop-Process -Id $serverProcess.Id
        $serverProcess.WaitForExit(5000) | Out-Null
    }
 
    foreach ($path in @(
        $databasePath,
        "$databasePath-journal",
        "$databasePath-wal",
        "$databasePath-shm",
        $stdoutPath,
        $stderrPath
    )) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
        }
    }
 
    foreach ($name in $originalEnvironment.Keys) {
        [Environment]::SetEnvironmentVariable(
            $name,
            $originalEnvironment[$name],
            "Process"
        )
    }
 
    Set-Location $originalLocation
}
 
exit $exitCode