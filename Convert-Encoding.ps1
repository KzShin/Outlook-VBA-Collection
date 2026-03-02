<#
.SYNOPSIS
    VBAソースコードの文字コードを変換します。
.DESCRIPTION
    指定したディレクトリ内の .bas, .cls, .frm ファイルの文字コードを
    ShiftJIS (CP932) と UTF-8 (BOMなし) の間で相互変換します。
    事前に文字コードを判定し、すでに変換先と同じ場合はスキップします。
    -WhatIf パラメータを付与すると、実際の変換を行わずにテスト実行できます。
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory=$true, HelpMessage="変換元の文字コードを指定してください (ShiftJIS または UTF8)")]
    [ValidateSet("ShiftJIS", "UTF8")]
    [string]$From,

    [Parameter(Mandatory=$true, HelpMessage="変換後の文字コードを指定してください (ShiftJIS または UTF8)")]
    [ValidateSet("ShiftJIS", "UTF8")]
    [string]$To,

    [Parameter(Mandatory=$false, HelpMessage="対象フォルダのパス (デフォルトは .\src)")]
    [string]$TargetFolder = ".\src"
)

if ($From -eq $To) {
    Write-Warning "変換元と変換後が同じ文字コード（$From）です。不要な変換を防ぐため処理を終了します。"
    exit
}

# 文字コードの定義
$encSJIS = [System.Text.Encoding]::GetEncoding(932)
$encUTF8 = New-Object System.Text.UTF8Encoding($false)

# 文字コード判定用（BOMなしUTF-8として厳密に読み込み、無効なバイト列があれば例外を出す設定）
$strictUTF8 = New-Object System.Text.UTF8Encoding($false, $true)

$encodingFrom = if ($From -eq "ShiftJIS") { $encSJIS } else { $encUTF8 }
$encodingTo   = if ($To -eq "ShiftJIS")   { $encSJIS } else { $encUTF8 }

# 対象フォルダの確認
if (-Not (Test-Path $TargetFolder)) {
    Write-Error "指定されたフォルダが見つかりません: $TargetFolder"
    exit
}

# 変換対象となる拡張子を指定
$targetFiles = Get-ChildItem -Path $TargetFolder -Include *.bas, *.cls, *.frm -Recurse

if ($targetFiles.Count -eq 0) {
    Write-Host "対象ファイルが見つかりませんでした。" -ForegroundColor Yellow
    exit
}

Write-Host "=== 文字コード変換を開始します ($From -> $To) ===" -ForegroundColor Cyan

foreach ($file in $targetFiles) {
    
    # 1. ファイルのバイト列を読み込んで現在の文字コードを推測
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $isUTF8 = $true
    try {
        # 厳密なUTF-8としてデコードを試みる
        $null = $strictUTF8.GetString($bytes)
    } catch {
        # Shift-JIS特有のバイト列などでエラーが出た場合はUTF-8ではないと判定
        $isUTF8 = $false
    }

    # 2. スキップ判定
    if ($To -eq "UTF8" -and $isUTF8) {
        Write-Host "[SKIP] $($file.Name) はすでに UTF-8 と推測されるためスキップしました。" -ForegroundColor DarkGray
        continue
    }
    if ($To -eq "ShiftJIS" -and -not $isUTF8) {
        Write-Host "[SKIP] $($file.Name) はすでに Shift-JIS と推測されるためスキップしました。" -ForegroundColor DarkGray
        continue
    }

    # 3. 変換処理 (ShouldProcessにより -WhatIf に対応)
    if ($PSCmdlet.ShouldProcess($file.FullName, "文字コード変換 ($From -> $To)")) {
        try {
            $content = [System.IO.File]::ReadAllText($file.FullName, $encodingFrom)
            [System.IO.File]::WriteAllText($file.FullName, $content, $encodingTo)
            Write-Host "[OK] $($file.Name) を変換しました。" -ForegroundColor Green
        }
        catch {
            Write-Error "[NG] $($file.Name) - $($_.Exception.Message)"
        }
    }
}

Write-Host "=== 処理が完了しました ===" -ForegroundColor Cyan
