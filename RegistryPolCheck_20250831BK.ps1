param(
    [string]$BackupDir = "C:\work\RegistryPolCheck\Backup",
    [int]$MaxBackup = 10,
    [string]$LogDir = "C:\work\RegistryPolCheck\Logs",
    [int]$LogRetentionDays = 7,
    [string]$NotifyBat = "C:\Scripts\TexpostNotify.bat",
    [string]$LGPOPath = "C:\Tools\LGPO.exe"   # 任意: LGPO.exe のパス
)

# ==== 初期化 ====
foreach ($d in @($BackupDir, $LogDir)) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d | Out-Null
    }
}

# ログファイル（日単位）
$logFile = Join-Path $LogDir ("RegistryPol_{0}.log" -f (Get-Date -Format "yyyyMMdd"))

# ログ出力関数
function Write-Log {
    param([string]$msg)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Add-Content -Path $logFile -Value $line
    Write-Output $line
}

# POLファイル検査関数（エラーコードを返す）
function Test-PolFile {
    param($path)

    if (-not (Test-Path $path)) { return "E01001" }   # ファイルなし
    $size = (Get-Item $path).Length
    if ($size -eq 0) { return "E01002" }              # サイズ0

    if (Test-Path $LGPOPath) {
        try {
            $out = & $LGPOPath /parse /m $path 2>&1
            if ($LASTEXITCODE -ne 0 -or $out -match "Error") {
                return "E01003"                       # 解析失敗
            }
        } catch {
            return "E01003"                           # 解析失敗
        }
    }

    return "E00000"                                   # 正常
}

# === 処理対象 ===
$polFiles = @{
    "Machine" = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"
    "User"    = "$env:SystemRoot\System32\GroupPolicy\User\Registry.pol"
}

$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$today     = (Get-Date).ToString("yyyyMMdd")
$problem   = @()

foreach ($key in $polFiles.Keys) {
    $polPath = $polFiles[$key]
    $code = Test-PolFile $polPath

    if ($code -eq "E00000") {
        # --- バックアップ判定 ---
        $todayBackup = Get-ChildItem $BackupDir -File -Filter "*${today}*${key}_Registry.pol" -ErrorAction SilentlyContinue
        if (-not $todayBackup) {
            $name = "${timestamp}_${key}_Registry.pol"
            $backupPath = Join-Path $BackupDir $name
            Copy-Item $polPath $backupPath -Force
            Write-Log "$key Registry.pol 正常 (E00000)、バックアップ作成: $backupPath"
        } else {
            Write-Log "$key Registry.pol 正常 (E00000)、本日のバックアップは既に存在"
        }
    } else {
        # --- エラー検出 ---
        switch ($code) {
            "E01001" { $desc = "ファイル不存在" }
            "E01002" { $desc = "サイズ0" }
            "E01003" { $desc = "解析失敗" }
            default  { $desc = "未知のエラー" }
        }
        Write-Log "警告: $key Registry.pol 異常、状態 = $code ($desc)"
        $problem += "$key: $code ($desc)"
    }
}

# --- 古いバックアップ削除 ---
Get-ChildItem $BackupDir -File -Filter "*Registry.pol" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip $MaxBackup |
    Remove-Item -Force

# --- 古いログ削除 ---
Get-ChildItem $LogDir -File -Filter "RegistryPol_*.log" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
    Remove-Item -Force

# --- エラー通知 ---
if ($problem.Count -gt 0) {
    $msg = "Registry.pol ファイルに問題を検出: " + ($problem -join ", ")
    Write-Log $msg

    if (Test-Path $NotifyBat) {
        Write-Log "外部通知バッチ実行: $NotifyBat"
        & $NotifyBat $msg
    } else {
        Write-Log "通知バッチが存在しません: $NotifyBat"
    }
}
