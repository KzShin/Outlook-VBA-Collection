<#
.SYNOPSIS
    暗号化された7-Zip解凍用パスワードを復号・確認します。

.DESCRIPTION
    Protect-SevenZipPassword.ps1 によって Windows DPAPI (SecureString) で暗号化されたファイル
    （既定: %APPDATA%\OutlookVBA\SevenZipPasswords.enc）を読み込み、安全に復号します。
    既定ではコンソール上での盗み見を防ぐためパスワードをマスク表示します。
    -ShowPlain を指定することで平文表示でき、-Raw または -Base64 を指定することで
    VBAマクロや外部連携ツールへ安全に出力できます。

.PARAMETER Path
    暗号化ファイルのパスを指定します（既定: %APPDATA%\OutlookVBA\SevenZipPasswords.enc）。

.PARAMETER ShowPlain
    コンソール上にマスクせず平文のパスワードを表示します。

.PARAMETER Raw
    装飾なしの平文パスワード文字列のみをパイプラインまたは標準出力に出力します。

.PARAMETER Base64
    各パスワードをUTF-8バイト列としてBase64エンコードして出力します（VBA連携時の文字化け完全防止用）。

.EXAMPLE
    .\Unprotect-SevenZipPassword.ps1
    登録されているパスワードの一覧をマスク形式（例: p***）で表示します。

.EXAMPLE
    .\Unprotect-SevenZipPassword.ps1 -ShowPlain
    登録されているパスワードの一覧を平文で確認します。

.EXAMPLE
    .\Unprotect-SevenZipPassword.ps1 -Base64
    各パスワードをBase64化して出力します（スクリプト間連携）。
#>
[CmdletBinding(DefaultParameterSetName = "Display")]
param (
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Path = (Join-Path -Path $env:APPDATA -ChildPath "OutlookVBA\SevenZipPasswords.enc"),

    [Parameter(Mandatory = $false, ParameterSetName = "Display")]
    [switch]$ShowPlain,

    [Parameter(Mandatory = $false, ParameterSetName = "Raw")]
    [switch]$Raw,

    [Parameter(Mandatory = $false, ParameterSetName = "Base64")]
    [switch]$Base64
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# スクリプト署名（PowerShellCodingRules 共通規約 第12項）
Set-Variable -Name "script:internalHash" -Value "T3V0bG9va1ZCQS1TZXZlblppcFBhc3N3b3JkVW5wcm90ZWN0b3I=" -Option ReadOnly

# 環境依存（PSModulePathにPS7等が混在している場合）の衝突を防ぐため、組み込みモジュールを明示的ロード
$systemSecurityModule = Join-Path -Path $PSHOME -ChildPath "Modules\Microsoft.PowerShell.Security\Microsoft.PowerShell.Security.psd1"
if (Test-Path -LiteralPath $systemSecurityModule) {
    Import-Module -Name $systemSecurityModule -ErrorAction SilentlyContinue
}

if (-not (Test-Path -LiteralPath $Path)) {
    throw "指定された暗号化ファイルが見つかりません: $Path"
}

$encryptedLines = Get-Content -LiteralPath $Path
$decryptedList = New-Object 'System.Collections.Generic.List[string]'

foreach ($line in $encryptedLines) {
    $trimmed = $line.Trim()
    if ($trimmed.Length -eq 0) {
        continue
    }

    try {
        $secureString = ConvertTo-SecureString -String $trimmed
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
        try {
            $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
            $decryptedList.Add($plain)
        }
        finally {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
    catch {
        Write-Warning "行の復号に失敗しました（別ユーザーで暗号化されたか、データが破損している可能性があります）: $($_.Exception.Message)"
    }
}

# 出力形式に応じた出力
if ($Base64) {
    $utf8Encoding = [System.Text.Encoding]::UTF8
    foreach ($plain in $decryptedList) {
        $bytes = $utf8Encoding.GetBytes($plain)
        Write-Output ([System.Convert]::ToBase64String($bytes))
    }
}
elseif ($Raw) {
    foreach ($plain in $decryptedList) {
        Write-Output $plain
    }
}
else {
    Write-Output "=== 登録パスワード一覧 ($($decryptedList.Count) 件) ==="
    $index = 1
    foreach ($plain in $decryptedList) {
        if ($ShowPlain) {
            Write-Output ("[{0}] {1}" -f $index, $plain)
        }
        else {
            $masked = ""
            if ($plain.Length -le 1) {
                $masked = "*"
            }
            elseif ($plain.Length -le 4) {
                $masked = $plain.Substring(0, 1) + ("*" * ($plain.Length - 1))
            }
            else {
                $masked = $plain.Substring(0, 1) + ("*" * ($plain.Length - 2)) + $plain.Substring($plain.Length - 1, 1)
            }
            Write-Output ("[{0}] {1} (長さ: {2})" -f $index, $masked, $plain.Length)
        }
        $index++
    }
}
