# RegistryPolCheck.ps1
# 処理概要： 本スクリプトは Windows サーバー上で利用されるローカルグループポリシー設定ファイル Registry.pol の健全性を確認し、異常を検知した場合にログ出力および通知を行うツールです。
# また、正常な場合は日次でバックアップを取得し、古いバックアップやログを自動的に削除します。
# 設定可能なパラメーターを利用して、バックアップ保存数やログ保存期間、通知バッチのパスなどをカスタマイズできます。
# LGPO.exe を利用して Registry.pol の内容を解析し、異常がないかを確認します。
# ただし LGPO.exe は必須ではなく、存在しない場合はファイルの存在とサイズのみをチェックします。
# LGPO.exe は Microsoft から無償で提供されており、必要に応じて導入してください。
# https://www.microsoft.com/en-us/download/details.aspx?id=55319
# なお、LGPO.exe を利用する場合は、スクリプト内の該当部分のコメントアウトを解除してください。

# 機能概要：
# - Registry.pol ファイルの存在確認とサイズチェック
# - LGPO.exe を利用した内容解析（オプション）
# - 日次バックアップの取得（重複防止）
# - 古いバックアップとログの自動削除
# - ログファイルへの詳細なログ出力（INFO / ERROR レベル）
# - 異常検知時の外部通知バッチの実行

# エラーレベルとコード例：
# - I00000: 正常終了
# - E01001: ファイル存在しません
# - E01002: サイズ0
# - E01003: 解析失敗
# - E02001: バックアップ失敗
# - E02002: 古いバックアップ削除失敗
# - E02003: 古いログ削除失敗
# - E09996: 不明なファイル操作エラー
# - E09997: 通知バッチ実行失敗
# - E09998: 通知バッチが存在しない
# - E09999: その他の未知のエラー

# 実行方法：
# - スクリプト: powershell.exe -ExecutionPolicy RemoteSigned -File "C:\work\RegistryPolCheck\RegistryPolCheck.ps1"

param(
    [string]$BackupDir = "C:\work\RegistryPolCheck\Backup",
    [int]$MaxBackup = 7,
    [string]$LogDir = "C:\work\RegistryPolCheck\Logs",
    [int]$LogRetentionDays = 7,
    [string]$NotifyBat = "C:\Scripts\TexpostNotify.bat",
    [string]$LGPOPath = "C:\Tools\LGPO.exe"   # 任意: LGPO.exe のパス
)

# --- ログ出力関数 ---
function Write-Log {
    param(
        [string]$Level,   # INFO / ERROR
        [string]$Code,    # I00000, E01001 ...
        [string]$Message
    )
    $timestamp = Get-Date -Format "[yyyy-MM-dd] [HH:mm:ss]"
    $line = "$timestamp [$Level] [$Code] $Message"
    Add-Content -Path $LogFile -Value $line
    Write-Output $line
}

# --- POLファイル検査関数 ---
function Test-PolFile {
    param($path)

    if (-not (Test-Path $path)) { return "E01001" }   # ファイルなし
    $size = (Get-Item $path).Length
    if ($size -eq 0) { return "E01002" }              # サイズ0

    # --- LGPO.exe の利用について ---
    # LGPO.exe は Microsoft 提供の Local Group Policy Object Utility であり、
    # Registry.pol の内容を解析するために使用可能。
    # ただし本サーバーに LGPO.exe が存在しない可能性があるため、
    # 一時的に利用部分をコメントアウトしている。
    # 将来的に内容チェックを行う場合は、LGPO.exe を導入し再度有効化すること。
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
    return "I00000"                                   # 正常
}

# --- ファイル操作ラッパー ---
function Invoke-FileAction {
    param(
        [string]$Action,   # "Get" / "Remove" / "Copy"
        [hashtable]$Params # コマンドに渡される引数
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
            "Copy"   { $errCode = "W01001"; $level = "WARN" }
            "Remove" { $errCode = "W01002"; $level = "WARN" }
            "Get"    { $errCode = "W01003"; $level = "WARN" }
            default  { $errCode = "E09996"; $level = "ERROR" }
        }
        $msg = "ファイル操作エラー - $Action : $($_.Exception.Message)"
        Write-Log -Level $level -Code $errCode -Message $msg

        # 通知は ERROR のときだけ
        if ($level -eq "ERROR") {
            $notifyArgs = "-s", "2", "-i", "`"A $msg`"", "-c", "kibt"
            & $NotifyBat @notifyArgs
        }

        return [pscustomobject]@{
            Result  = $null
            ErrCode = $errCode
            Level   = $level
            Message = $msg
        }
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

# === 処理対象 ===
$polFiles = @{
    "Machine" = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"
}

$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$today     = (Get-Date).ToString("yyyyMMdd")
$problem   = @()

# 検査開始ログ
Write-Log -Level "INFO" -Code "I00000" -Message "検査開始"

# --- 各POLファイル検査 ---
foreach ($key in $polFiles.Keys) {
    $polPath = $polFiles[$key]
    $code = Test-PolFile $polPath

    if ($code -eq "I00000") {
        Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol 検査実施 → 正常"

        # --- バックアップ判定 ---
        $todayBackup = Get-ChildItem $BackupDir -File -Filter "*${today}*${key}_Registry.pol" -ErrorAction SilentlyContinue

        if (-not $todayBackup) {
            $name = "${timestamp}_${key}_Registry.pol"
            $backupPath = Join-Path $BackupDir $name
            $result = Invoke-FileAction -Action "Copy" -Params @{ Path = $polPath; Destination = $backupPath; Force = $true }
            if ($result -is [pscustomobject] -and $result.ErrCode) {
                # エラー発生
                if ($result.Level -eq "WARN") {
                    $msg = "バックアップ処理 → 失敗: $($result.Message)"
                    Write-Log -Level $result.Level -Code $result.ErrCode -Message $msg
                    $notifyArgs = "-s", "1", "-i", "`"A $msg`"", "-c", "kibt"
                    & $NotifyBat @notifyArgs                    
                }
            } else {
                Write-Log -Level "INFO" -Code "I00000" -Message "$key Registry.pol バックアップ処理 → 作成: $backupPath"
            }
        } else {
            Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol バックアップ処理 → 本日のバックアップは既に存在"
        }
    } else {
        switch ($code) {
            "E01001" { $desc = "ファイル存在しません" }
            "E01002" { $desc = "サイズ0" }
            "E01003" { $desc = "LGPO.exe による Registry.pol 解析に失敗しました" }
            default  { $desc = "未知のエラー" }
        }
        Write-Log -Level "ERROR" -Code $code -Message "$key Registry.pol 検査実施 → 異常、状態 = $desc"
        $problem += "$($key): $code ($desc)"
    }
}

# 検査終了ログ
Write-Log -Level "INFO" -Code "I00000" -Message "検査終了"

# --- 古いバックアップ削除 ---
Get-ChildItem $BackupDir -File -Filter "*Registry.pol" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip $MaxBackup |
    ForEach-Object {
        $result = Invoke-FileAction -Action "Remove" -Params @{ Path = $_.FullName; Force = $true }
        if ($result -is [pscustomobject] -and $result.ErrCode) {
            # エラー発生
            if ($result.Level -eq "WARN") {
                $msg = "古いバックアップ削除 → 失敗: $($result.Message)"
                Write-Log -Level $result.Level -Code $result.ErrCode -Message $msg
                $notifyArgs = "-s", "1", "-i", "`"A $msg`"", "-c", "kibt"
                & $NotifyBat @notifyArgs
            }
        } else {
            Write-Log -Level "INFO" -Code "I00000" -Message "古いバックアップ削除 → $($_.FullName)"
        }
    }

# --- 古いログ削除 ---
Get-ChildItem $LogDir -File -Filter "RegistryPol_*.log" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
    ForEach-Object {
        $result = Invoke-FileAction -Action "Remove" -Params @{ Path = $_.FullName; Force = $true }
        if ($result -is [pscustomobject] -and $result.ErrCode) {
            # エラー発生
            if ($result.Level -eq "WARN") {
                $msg = "古いログ削除 → 失敗: $($result.Message)"
                Write-Log -Level $result.Level -Code $result.ErrCode -Message $msg
                $notifyArgs = "-s", "1", "-i", "`"A $msg`"", "-c", "kibt"
                & $NotifyBat @notifyArgs
            }
        } else {
            Write-Log -Level "INFO" -Code "I00000" -Message "古いログ削除 → $($_.FullName)"
        }
    }

# --- エラー通知 ---
if ($problem.Count -gt 0) {
    $msg = "Registry.pol ファイルに問題を検出: " + ($problem -join ", ")
    Write-Log -Level "ERROR" -Code "E09999" -Message $msg

    if (Test-Path $NotifyBat) {
        Write-Log -Level "INFO" -Code "I00000" -Message "外部通知バッチ実行: $NotifyBat"
        $notifyArgs = "-s", "2", "-i", "`"A $msg`"", "-c", "kibt"
        & $NotifyBat @notifyArgs
        if ($LASTEXITCODE -ne 0) {
            Write-Log -Level "ERROR" -Code "E09997" -Message "通知バッチの実行に失敗しました。戻り値異常: ExitCode=$LASTEXITCODE"
        } else {
            Write-Log -Level "INFO" -Code "I00000" -Message "通知バッチの実行に成功しました： $NotifyBat （レベル： MINOR）"
        }
    } else {
        Write-Log -Level "ERROR" -Code "E09998" -Message "外部通知バッチが存在しません: $NotifyBat"
    }
}
