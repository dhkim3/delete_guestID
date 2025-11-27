#requires -Version 5.1

# 모듈 설치 (GitHub Actions 환경에서 매번 설치)
Install-Module -Name ImportExcel -Force -Scope CurrentUser
Install-Module -Name Microsoft.Graph -Force -Scope CurrentUser
Install-Module -Name Microsoft.Graph.Auth -Force -Scope CurrentUser

# ClientSecretCredential 객체 생성 (방법 A)
$ClientSecretCred = [Microsoft.Graph.Auth.ClientSecretCredential]::new(
    $env:AZURE_TENANT_ID,
    $env:AZURE_CLIENT_ID,
    $env:AZURE_CLIENT_SECRET
)

# Microsoft Graph 연결
Connect-MgGraph -ClientSecretCredential $ClientSecretCred -Scopes "User.Read.All","AuditLog.Read.All"

# Excel 파일명
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$OutXlsx = "GuestUsers_$ts.xlsx"

# Guest 계정 조회
$users = Get-MgUser -Filter "userType eq 'Guest'" -All -Property "displayName,mail,otherMails,userPrincipalName,id,signInActivity"

# 데이터 가공
$rows = $users | ForEach-Object {
    $email = if ($_.Mail) { $_.Mail } elseif ($_.OtherMails.Count) { $_.OtherMails[0] } else { $_.UserPrincipalName }
    $s = $_.SignInActivity

    $lastInteractive = if ($s -and $s.LastSuccessfulSignInDateTime) { [datetime]$s.LastSuccessfulSignInDateTime } else { $null }
    $lastNonInteractive = if ($s -and $s.LastNonInteractiveSignInDateTime) { [datetime]$s.LastNonInteractiveSignInDateTime } else { $null }

    $latestLogin = if ($lastInteractive -and $lastNonInteractive) {
        ($lastInteractive, $lastNonInteractive | Sort-Object -Descending)[0]
    } elseif ($lastInteractive) {
        $lastInteractive
    } elseif ($lastNonInteractive) {
        $lastNonInteractive
    } else { $null }

    [pscustomobject]@{
        Name = $_.DisplayName
        Email = $email
        ObjectId = $_.Id
        LastSignIn_대화형 = if ($lastInteractive) { $lastInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "로그인 기록 없음" }
        LastSignIn_비대화형 = if ($lastNonInteractive) { $lastNonInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "로그인 기록 없음" }
        LastSeen = if ($latestLogin) { $latestLogin.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") } else { "로그인 기록 없음" }
    }
}

# Excel로 내보내기
$rows | Export-Excel -Path $OutXlsx -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName 'GuestUsers'
