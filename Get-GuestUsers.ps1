#requires -Version 5.1
<#
.SYNOPSIS
    Microsoft Graph Guest 계정 조회 및 Excel 내보내기
.DESCRIPTION
    - 앱 인증(App-only) 방식
    - 대화형 / 비대화형 로그인 기록 포함
    - Excel 파일로 자동 생성
#>

# ⚙️ ImportExcel 모듈 설치
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
}

param(
    [string]$OutDir = $PWD   # GitHub Actions에서는 현재 작업 폴더
)

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$OutXlsx = Join-Path $OutDir ("GuestUsers_{0}.xlsx" -f $ts)

# 환경 변수에서 앱 인증 정보 가져오기
$ClientId     = $env:AZURE_CLIENT_ID
$TenantId     = $env:AZURE_TENANT_ID
$ClientSecret = $env:AZURE_CLIENT_SECRET

# Microsoft Graph 앱 인증
if (-not (Get-MgContext)) {
    Connect-MgGraph -ClientId $ClientId `
                    -TenantId $TenantId `
                    -ClientSecret $ClientSecret `
                    -Scopes "https://graph.microsoft.com/.default"
}

# 폴더 생성
if ($OutDir -and -not (Test-Path $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir | Out-Null
}

# 베타 프로필로 전환 후 조회
$prevProfile = (Get-MgProfile).Name
Select-MgProfile -Name beta

try {
    $users = Get-MgUser -Filter "userType eq 'Guest'" -All -Property `
        "displayName,mail,otherMails,userPrincipalName,id,signInActivity"
}
finally {
    if ($prevProfile) { Select-MgProfile -Name $prevProfile } else { Select-MgProfile -Name "v1.0" }
}

# 데이터 가공
$rows = $users | ForEach-Object {
    $email = if ($_.Mail) { $_.Mail }
             elseif ($_.OtherMails -and $_.OtherMails.Count) { $_.OtherMails[0] }
             else { $_.UserPrincipalName }

    $s = $_.SignInActivity

    $lastInteractive = if ($s -and $s.LastSuccessfulSignInDateTime) { [datetime]$s.LastSuccessfulSignInDateTime }
                       elseif ($s -and $s.LastSignInDateTime)       { [datetime]$s.LastSignInDateTime }
                       else { $null }

    $lastNonInteractive = if ($s -and $s.LastNonInteractiveSignInDateTime) { [datetime]$s.LastNonInteractiveSignInDateTime }
                          else { $null }

    $lastInteractiveStr = if ($lastInteractive) { $lastInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "로그인 기록 없음" }
    $lastNonInteractiveStr = if ($lastNonInteractive) { $lastNonInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "로그인 기록 없음" }

    $latestLogin = if (-not $lastInteractive -and -not $lastNonInteractive) {
        "로그인 기록 없음"
    } elseif ($lastInteractive -and $lastNonInteractive) {
        ($lastInteractive, $lastNonInteractive | Sort-Object -Descending)[0].ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
    } elseif ($lastInteractive) {
        $lastInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
    } else {
        $lastNonInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
    }

    [pscustomobject]@{
        Name                   = $_.DisplayName
        Email                  = $email
        ObjectId               = $_.Id
        LastSignIn_대화형       = $lastInteractiveStr
        LastSignIn_비대화형     = $lastNonInteractiveStr
        LastSeen               = $latestLogin
    }
}

# Excel 내보내기
$rows | Export-Excel -Path $OutXlsx -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName 'GuestUsers'

[Console]::Write($OutXlsx)
