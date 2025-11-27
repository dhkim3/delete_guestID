#requires -Version 5.1

# 모듈 설치
Install-Module -Name ImportExcel -Force -Scope CurrentUser
Install-Module -Name Microsoft.Graph -Force -Scope CurrentUser

# 로그인
Connect-MgGraph -ClientId $env:AZURE_CLIENT_ID `
                -TenantId $env:AZURE_TENANT_ID `
                -ClientSecret $env:AZURE_CLIENT_SECRET `
                -Scopes "User.Read.All","AuditLog.Read.All"

# 파일명
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$OutXlsx = "GuestUsers_$ts.xlsx"

# Guest 계정 조회
$users = Get-MgUser -Filter "userType eq 'Guest'" -All -Property "displayName,mail,userPrincipalName,id,signInActivity"

# 가공
$rows = $users | ForEach-Object {
    $email = if ($_.Mail) { $_.Mail } else { $_.UserPrincipalName }
    $s = $_.SignInActivity
    $lastInteractive = if ($s -and $s.LastSuccessfulSignInDateTime) { [datetime]$s.LastSuccessfulSignInDateTime } else { $null }
    $lastNonInteractive = if ($s -and $s.LastNonInteractiveSignInDateTime) { [datetime]$s.LastNonInteractiveSignInDateTime } else { $null }
    $latestLogin = if ($lastInteractive -and $lastNonInteractive) { ($lastInteractive, $lastNonInteractive | Sort-Object -Descending)[0] } elseif ($lastInteractive) { $lastInteractive } elseif ($lastNonInteractive) { $lastNonInteractive } else { $null }
    [pscustomobject]@{
        Name = $_.DisplayName
        Email = $email
        ObjectId = $_.Id
        LastSignIn_대화형 = if ($lastInteractive) { $lastInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "로그인 기록 없음" }
        LastSignIn_비대화형 = if ($lastNonInteractive) { $lastNonInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "로그인 기록 없음" }
        LastSeen = if ($latestLogin) { $latestLogin.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "로그인 기록 없음" }
    }
}

# 엑셀 내보내기
$rows | Export-Excel -Path $OutXlsx -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName 'GuestUsers'
