<#
.SYNOPSIS
    7-Zip解凍用パスワードを暗号化して保存・管理します。

.DESCRIPTION
    平文のパスワードまたは既存のテキストファイル（SevenZipPasswords.txt）からパスワードを読み込み、
    Windows DPAPI (Data Protection API) による暗号化（SecureString）を施して
    暗号化ファイル（既定: %APPDATA%\OutlookVBA\SevenZipPasswords.enc）へ保存します。
    暗号化は実行したWindowsユーザーのアカウントに紐づけられるため、管理者権限のない標準ユーザーでも
    安全にパスワードを保護できます。

.PARAMETER Password
    暗号化して登録するパスワード文字列の配列を指定します。パイプラインからの入力にも対応しています。

.PARAMETER ImportFromTxt
    既存の平文パスワードファイルからインポートして一括暗号化する場合に指定します。

.PARAMETER InputPath
    インポート元の平文ファイルパスを指定します（既定: %APPDATA%\OutlookVBA\SevenZipPasswords.txt）。

.PARAMETER OutputPath
    出力先の暗号化ファイルパスを指定します（既定: %APPDATA%\OutlookVBA\SevenZipPasswords.enc）。

.PARAMETER Append
    既存の暗号化ファイルが存在する場合に、既存の内容を保持したまま末尾に追記します。

.PARAMETER RemoveSource
    インポート完了後、平文の元ファイルを削除する場合に指定します。

.EXAMPLE
    .\Protect-SevenZipPassword.ps1 -Password "P@ssword123", "Secret#{yyyy}"
    指定した2件のパスワードを暗号化して SevenZipPasswords.enc に保存します。

.EXAMPLE
    .\Protect-SevenZipPassword.ps1 -ImportFromTxt
    既定の SevenZipPasswords.txt からパスワードを読み込み、暗号化して SevenZipPasswords.enc に保存します。

.EXAMPLE
    .\Protect-SevenZipPassword.ps1 -ImportFromTxt -RemoveSource -WhatIf
    平文ファイルの暗号化移行および元ファイル削除のシミュレーションを実行します。
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = "Add")]
param (
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ParameterSetName = "Add")]
    [ValidateNotNullOrEmpty()]
    [string[]]$Password,

    [Parameter(Mandatory = $true, ParameterSetName = "Import")]
    [switch]$ImportFromTxt,

    [Parameter(Mandatory = $false, ParameterSetName = "Import")]
    [ValidateNotNullOrEmpty()]
    [string]$InputPath = (Join-Path -Path $env:APPDATA -ChildPath "OutlookVBA\SevenZipPasswords.txt"),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path $env:APPDATA -ChildPath "OutlookVBA\SevenZipPasswords.enc"),

    [Parameter(Mandatory = $false)]
    [switch]$Append,

    [Parameter(Mandatory = $false, ParameterSetName = "Import")]
    [switch]$RemoveSource
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# スクリプト署名（PowerShellCodingRules 共通規約 第12項）
Set-Variable -Name "script:internalHash" -Value "T3V0bG9va1ZCQS1TZXZlblppcFBhc3N3b3JkUHJvdGVjdG9y" -Option ReadOnly

# 環境依存（PSModulePathにPS7等が混在している場合）の衝突を防ぐため、組み込みモジュールを明示的ロード
$systemSecurityModule = Join-Path -Path $PSHOME -ChildPath "Modules\Microsoft.PowerShell.Security\Microsoft.PowerShell.Security.psd1"
if (Test-Path -LiteralPath $systemSecurityModule) {
    Import-Module -Name $systemSecurityModule -ErrorAction SilentlyContinue
}

# 出力先フォルダの存在確認・自動作成
$targetFolder = [System.IO.Path]::GetDirectoryName($OutputPath)
if (-not [string]::IsNullOrEmpty($targetFolder) -and -not (Test-Path -LiteralPath $targetFolder)) {
    if ($PSCmdlet.ShouldProcess($targetFolder, "ディレクトリ作成")) {
        New-Item -Path $targetFolder -ItemType Directory -Force | Out-Null
        Write-Verbose "出力先ディレクトリを作成しました: $targetFolder"
    }
}

# 登録対象パスワードを収集するリスト
$passwordsToProtect = New-Object 'System.Collections.Generic.List[string]'

# 既存ファイルの読み込み（追記モード時）
if ($Append -and (Test-Path -LiteralPath $OutputPath)) {
    Write-Verbose "既存の暗号化ファイルを保持して追記します: $OutputPath"
    $existingLines = Get-Content -LiteralPath $OutputPath -Encoding UTF8
    foreach ($line in $existingLines) {
        $trimmed = $line.Trim()
        if ($trimmed.Length -gt 0) {
            # 既に暗号化されている行はそのまま後で再利用するため保持するか、
            # あるいは復号して追記するか。ここでは既存暗号化行を直接リストに保持
        }
    }
}

# パラメータセットごとの処理
switch ($PSCmdlet.ParameterSetName) {
    "Add" {
        foreach ($p in $Password) {
            if (-not [string]::IsNullOrEmpty($p)) {
                $passwordsToProtect.Add($p)
            }
        }
    }
    "Import" {
        if (-not (Test-Path -LiteralPath $InputPath)) {
            throw "インポート元ファイルが見つかりません: $InputPath"
        }

        Write-Information "平文ファイルからパスワードを読み込んでいます: $InputPath"
        $rawLines = Get-Content -LiteralPath $InputPath -Encoding UTF8

        foreach ($raw in $rawLines) {
            $trimmed = $raw.Trim()
            if ($trimmed.Length -eq 0) {
                continue
            }

            $effectivePassword = ""
            if ($trimmed.Length -ge 2 -and $trimmed.StartsWith('"') -and $trimmed.EndsWith('"')) {
                # ダブルクォート囲みの場合は引用符を除去して採用
                $effectivePassword = $trimmed.Substring(1, $trimmed.Length - 2)
            }
            elseif (-not $trimmed.StartsWith("//")) {
                # // 以外の行（# で始まる行も含む）を採用
                $effectivePassword = $trimmed
            }

            if (-not [string]::IsNullOrEmpty($effectivePassword)) {
                $passwordsToProtect.Add($effectivePassword)
            }
        }

        Write-Information "インポート対象件数: $($passwordsToProtect.Count) 件"
    }
}

if ($passwordsToProtect.Count -eq 0 -and -not $Append) {
    Write-Warning "暗号化対象のパスワードがありません。"
    exit 0
}

# 暗号化処理
$encryptedLines = New-Object 'System.Collections.Generic.List[string]'

# 追記モードで既存ファイルが存在する場合、既存の暗号化行を先頭に引き継ぐ
if ($Append -and (Test-Path -LiteralPath $OutputPath)) {
    $existingEncLines = Get-Content -LiteralPath $OutputPath
    foreach ($el in $existingEncLines) {
        $trimmedEl = $el.Trim()
        if ($trimmedEl.Length -gt 0) {
            $encryptedLines.Add($trimmedEl)
        }
    }
}

# 新規パスワードの暗号化
foreach ($plain in $passwordsToProtect) {
    $secureString = ConvertTo-SecureString -String $plain -AsPlainText -Force
    $encrypted = ConvertFrom-SecureString -SecureString $secureString
    $encryptedLines.Add($encrypted)
}

# ファイル書き込み
if ($PSCmdlet.ShouldProcess($OutputPath, "パスワード暗号化保存 ($($passwordsToProtect.Count) 件追加, 合計 $($encryptedLines.Count) 件)")) {
    # UTF-8 (BOM付き) で保存
    $utf8WithBom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllLines($OutputPath, $encryptedLines, $utf8WithBom)
    Write-Output "パスワードを正常に暗号化保存しました: $OutputPath (合計: $($encryptedLines.Count) 件)"

    # インポート時の元ファイル削除処理
    if ($PSCmdlet.ParameterSetName -eq "Import" -and $RemoveSource) {
        if ($PSCmdlet.ShouldProcess($InputPath, "平文ファイルの削除")) {
            Remove-Item -LiteralPath $InputPath -Force
            Write-Information "平文の元ファイルを削除しました: $InputPath"
        }
    }
}
