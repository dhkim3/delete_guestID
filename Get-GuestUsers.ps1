#requires -Version 5.1

param(
  [string]$OutDir = "output"
)

# 폴더 생성
if (-not (Test-Path $OutDir)) {
  New-Item -ItemType Directory -Path $OutDir | Out-Null
}

# 모듈 설치
Install-Module -Name ImportExcel -Force -Scope CurrentUser
Install-Module -Name Microsoft.Graph -Force -Scope CurrentUser

# Microsoft.Identity.Client 필요 없음, Microsoft.Graph.Auth도 설치하지 마세요

# ClientSecretCredential 객체 생성
$ClientSecretCred = [Microsoft.Graph.Auth.ClientSecretCredential]::new(
    $env:AZURE_TENANT_ID,
    $env:AZURE_CLIENT_ID,
    $env:AZURE_CLIENT_SECRET
)

# Graph 연결
Connect-MgGraph -ClientSecretCredential $ClientSecretCred -Scopes "User.Read.All","AuditLog.Read.All"


# 결과 파일명
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$OutXlsx = Join-Path $OutDir ("GuestUsers_{0}.xlsx" -f $ts)

# Guest 조회 (Beta)
$prev = (Get-MgProfile).Name
Select-MgProfile -Name beta

try {
  $users = Get-MgUser -Filter "userType eq 'Guest'" -All -Property `
    "displayName,mail,otherMails,userPrincipalName,id,signInActivity"
}
finally {
  Select-MgProfile -Name $prev
}

# 가공
$rows = $users | ForEach-Object {
  $email = if ($_.Mail) { $_.Mail }
           elseif ($_.OtherMails -and $_.OtherMails.Count) { $_.OtherMails[0] }
           else { $_.UserPrincipalName }

  $s = $_.SignInActivity

  $lastInteractive =
    if ($s -and $s.LastSuccessfulSignInDateTime) { [datetime]$s.LastSuccessfulSignInDateTime }
    elseif ($s -and $s.LastSignInDateTime)       { [datetime]$s.LastSignInDateTime }
    else { $null }

  $lastInteractiveStr = if ($lastInteractive) {
      ($lastInteractive.ToLocalTime()).ToString("yyyy-MM-dd HH:mm:ss")
  } else { "로그인 기록 없음" }

  $lastNonInteractive =
    if ($s -and $s.LastNonInteractiveSignInDateTime) { [datetime]$s.LastNonInteractiveSignInDateTime }
    else { $null }

  $lastNonInteractiveStr = if ($lastNonInteractive) {
      ($lastNonInteractive.ToLocalTime()).ToString("yyyy-MM-dd HH:mm:ss")
  } else { "로그인 기록 없음" }

  $latestLogin =
    if (-not $lastInteractive -and -not $lastNonInteractive) {
      "로그인 기록 없음"
    }
    elseif ($lastInteractive -and $lastNonInteractive) {
      ($lastInteractive, $lastNonInteractive | Sort-Object -Descending)[0].ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
    }
    elseif ($lastInteractive) {
      $lastInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
    }
    else {
      $lastNonInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
    }

  [pscustomobject]@{
    Name                    = $_.DisplayName
    Email                   = $email
    ObjectId                = $_.Id
    LastSignIn_대화형        = $lastInteractiveStr
    LastSignIn_비대화형      = $lastNonInteractiveStr
    LastSeen                = $latestLogin
  }
}

# 엑셀로 저장
$rows | Export-Excel -Path $OutXlsx -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName 'GuestUsers'

Write-Host "Export Completed: $OutXlsx"


