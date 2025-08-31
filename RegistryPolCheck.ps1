param(
    [string]$BackupDir = "C:\work\RegistryPolCheck\Backup",
    [int]$MaxBackup = 10,
    [string]$LogDir = "C:\work\RegistryPolCheck\Logs",
    [int]$LogRetentionDays = 7,
    [string]$NotifyBat = "C:\Scripts\TexpostNotify.bat",
    [string]$LGPOPath = "C:\Tools\LGPO.exe"   # 任意: LGPO.exe のパス
)

# --- ログ出力関数 ---
function Write-Log {
    param(
        [string]$Level,   # INFO / ERROR
        [string]$Code,    # E00000, E01001 ...
        [string]$Message
    )
    $timestamp = Get-Date -Format "[yyyy-MM-dd] [HH:mm:ss]"
    $line = "$timestamp [$Level] [$Code] $Message"
    Add-Content -Path $LogFile -Value $line
    Write-Output $line
}

# --- ファイル操作ラッパー ---
function Invoke-FileAction {
    param(
        [string]$Action,   # "Get" / "Remove" / "Copy"
        [hashtable]$Params # ???命令的参数
    )
    try {
        switch ($Action) {
            "Get"    { return Get-ChildItem @Params }
            "Remove" { Remove-Item @Params -ErrorAction Stop }
            "Copy"   { Copy-Item @Params -ErrorAction Stop }
            default  { throw "未知操作: $Action" }
        }
    } catch {
        switch ($Action) {
            "Copy"   { $errCode = "E02001" }
            "Remove" { $errCode = "E02002" }
            "Get"    { $errCode = "E02003" }
            default  { $errCode = "E09997" }
        }
        Write-Log -Level "ERROR" -Code $errCode -Message "$Action 実行失敗 - $($_.Exception.Message)"
        return $null
    }
}

# ==== 初期化 ====
foreach ($d in @($BackupDir, $LogDir)) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d | Out-Null
    }
}

# ログファイル（日単位）
$logFile = Join-Path $LogDir ("RegistryPol_{0}.log" -f (Get-Date -Format "yyyyMMdd"))

# --- POLファイル検査関数 ---
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
}

$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$today     = (Get-Date).ToString("yyyyMMdd")
$problem   = @()

# 検査開始ログ
Write-Log -Level "INFO" -Code "E00000" -Message "検査開始"

# --- 各POLファイル検査 ---
foreach ($key in $polFiles.Keys) {
    $polPath = $polFiles[$key]
    $code = Test-PolFile $polPath

    if ($code -eq "E00000") {
        Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol 検査実施 → 正常"

        # --- バックアップ判定 ---
        $todayBackup = Get-ChildItem $BackupDir -File -Filter "*${today}*${key}_Registry.pol" -ErrorAction SilentlyContinue

        if (-not $todayBackup) {
            $name = "${timestamp}_${key}_Registry.pol"
            $backupPath = Join-Path $BackupDir $name
            Invoke-FileAction -Action "Copy" -Params @{ Path = $polPath; Destination = $backupPath; Force = $true }
            Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol バックアップ処理 → 作成: $backupPath"
        } else {
            Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol バックアップ処理 → 本日のバックアップは既に存在"
        }
    } else {
        switch ($code) {
            "E01001" { $desc = "ファイル不存在" }
            "E01002" { $desc = "サイズ0" }
            "E01003" { $desc = "解析失敗" }
            default  { $desc = "未知のエラー" }
        }
        Write-Log -Level "ERROR" -Code $code -Message "$key Registry.pol 検査実施 → 異常、状態 = $desc"
        $problem += "$($key): $code ($desc)"
    }
}

# 検査終了ログ
Write-Log -Level "INFO" -Code "E00000" -Message "検査終了"

# --- 古いバックアップ削除 ---
Get-ChildItem $BackupDir -File -Filter "*Registry.pol" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip $MaxBackup |
    ForEach-Object {
        Invoke-FileAction -Action "Remove" -Params @{ Path = $_.FullName; Force = $true }
        Write-Log -Level "INFO" -Code "E00000" -Message "古いバックアップ削除 → $($_.FullName)"
    }

# --- 古いログ削除 ---
Get-ChildItem $LogDir -File -Filter "RegistryPol_*.log" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
    ForEach-Object {
        Invoke-FileAction -Action "Remove" -Params @{ Path = $_.FullName; Force = $true }
        Write-Log -Level "INFO" -Code "E00000" -Message "古いログ削除 → $($_.FullName)"
    }

# --- エラー通知 ---
if ($problem.Count -gt 0) {
    $msg = "Registry.pol ファイルに問題を検出: " + ($problem -join ", ")
    Write-Log -Level "ERROR" -Code "E09999" -Message $msg

    if (Test-Path $NotifyBat) {
        Write-Log -Level "INFO" -Code "E00000" -Message "外部通知バッチ実行: $NotifyBat"
        & $NotifyBat $msg
    } else {
        Write-Log -Level "ERROR" -Code "E09998" -Message "通知バッチが存在しません: $NotifyBat"
    }
}
