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
# - W01001: バックアップ失敗
# - W01002: 古いバックアップ削除失敗
# - W01003: 古いログ削除失敗
# - W01004: 管理者権限なし警告
# - E01001: ファイル存在しません
# - E01002: サイズ0
# - E01003: 解析失敗
# - E09994: 必須ディレクトリ作成失敗
# - E09995: 不明なファイル操作エラー
# - E09996: 通知バッチ実行異常
# - E09997: 通知バッチが存在しない
# - E09998: Registry.pol に問題検出
# - E09999: 予期しないエラー

# 実行方法：
# - スクリプト: powershell.exe -ExecutionPolicy RemoteSigned -File "C:\work\RegistryPolCheck\RegistryPolCheck.ps1"

# --- パラメーター定義 ---
[CmdletBinding()]
param(
    [ValidateScript({Test-Path $_ -IsValid})]
    [string]$BackupDir = "C:\work\RegistryPolCheck\Backup",

    [ValidateRange(1, 30)]
    [int]$MaxBackup = 7,

    [ValidateScript({Test-Path $_ -IsValid})]
    [string]$LogDir = "C:\work\RegistryPolCheck\Logs",

    [ValidateRange(1, 30)]
    [int]$LogRetentionDays = 7,

    [ValidateScript({Test-Path $_ -IsValid})]
    [string]$NotifyBat = "C:\Scripts\TexpostNotify.bat",

    [ValidateScript({Test-Path $_ -IsValid})]
    [string]$LGPOPath = "C:\Tools\LGPO.exe"   # 任意: LGPO.exe のパス
)

# --- グローバル変数 ---
$script:LogFile = $null  # ログファイルパス（実行時に設定される）
$NotifyTimeoutSeconds = 30
$LGPOTimeoutSeconds   = 30

# --- 管理者権限チェック ---
# 管理者権限がない場合は警告ログを出力するが、処理は継続する
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $msg = "このスクリプトは管理者権限で実行することを推奨します。権限不足により一部の操作が失敗する可能性があります。"
    # Invoke-Notify -Level "WARN" -Message $msg
    # Write-Log -Level "WARN" -Code "W01004" -Message $msg
}

# グローバルエラー処理設定
$ErrorActionPreference = "Continue"  # エラーが発生した場合でも実行を継続する
$ProgressPreference = "SilentlyContinue"  # タスクスケジューラでの表示問題を回避するため、プログレスバーを無効化します。

# --- ログ出力関数 ---
function Write-Log {
    param(
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level,
        [string]$Code,
        [string]$Message
    )
    
    if (-not $script:LogFile) {
        # throw "LogFile変数が初期化されていません"
        Write-Output "[FALLBACK-LOG] $Message"
        return        
    }
    
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    $line = "$timestamp [$Level] [$Code] $Message"
    
    try {
        Add-Content -Path $script:LogFile -Value $line -ErrorAction Stop
        Write-Output $line
    } catch {
        Write-Error "ログファイルへの書き込みができません: $_"
    }
}

# --- 通知バッチ実行関数 ---
function Invoke-Notify {
    param(
        [ValidateSet("WARN", "ERROR", "MINOR")]
        [string]$Level,   # "ERROR" or "WARN"
        [string]$Message
    )

    # Level → s値 のマッピング
    $map = @{
        "WARN"  = "1"        
        "ERROR" = "2"
        "MINOR" = "3"
    }
    $s = $map[$Level]
    if (-not $s) { $s = "3" }  # fallback: 未定義LevelはMINOR扱い

    if (Test-Path $NotifyBat) {
		Write-Log -Level "INFO" -Code "I00000" -Message "外部通知バッチ実行: $NotifyBat (レベル： $Level)"
        try {
            $notifyArgs = "-s", $s, "-i", "`"A $Message`"", "-c", "kibt"
            $job = Start-Job -ScriptBlock {
                param($bat, $notifyArgs)
                & $bat @notifyArgs
                return $LASTEXITCODE
            } -ArgumentList $NotifyBat, $notifyArgs
            
            if (Wait-Job $job -Timeout 30) {
                $exitCode = Receive-Job $job
                Remove-Job $job -ErrorAction SilentlyContinue

                if ($exitCode -ne 0) {
                    Write-Log -Level "ERROR" -Code "E09996" -Message "通知バッチの実行に失敗しました。戻り値異常: ExitCode=$exitCode | $NotifyBat （レベル： $Level）"
                } else {
                    Write-Log -Level "INFO" -Code "I00000" -Message "通知バッチの実行に成功しました： $NotifyBat （レベル： $Level）"
                }

            } else {
                Stop-Job $job -ErrorAction SilentlyContinue
                Remove-Job $job -Force -ErrorAction SilentlyContinue
                Write-Log -Level "ERROR" -Code "E09996" -Message "通知バッチの実行がタイムアウトしました（ $NotifyTimeoutSeconds 秒）: $NotifyBat"
            }         
        }
        catch {
            Write-Log -Level "ERROR" -Code "E09996" -Message "通知バッチ実行中に例外が発生: $($_.Exception.Message)"
        }
    } else {
        Write-Log -Level "ERROR" -Code "E09997" -Message "外部通知バッチが存在しません: $NotifyBat"
    }
}

# --- POLファイル検査関数 ---
function Test-PolFile {
    param($path, [ref]$ErrorDetails)

    if (-not (Test-Path $path)) { return "E01001" }   # ファイルなし
    $size = (Get-Item $path).Length
    if ($size -eq 0) { return "E01002" }              # サイズ0

    # --- LGPO.exe の利用について ---
    # LGPO.exe は Microsoft 提供の Local Group Policy Object Utility であり、
    # Registry.pol の内容を解析するために使用可能。
    # ただし本サーバーに LGPO.exe が存在しない可能性があるため、
    # 一時的に利用部分をコメントアウトしている。
    # 将来的に内容チェックを行う場合は、LGPO.exe を導入し再度有効化すること。
    if ($LGPOPath -and (Test-Path $LGPOPath)) {
        try {
            <#
            $process = Start-Process -FilePath $LGPOPath -ArgumentList "/parse /q /m $polPath" -PassThru
            if (-not (Wait-Process -Id $process.Id -Timeout $LGPOTimeoutSeconds -ErrorAction SilentlyContinue)) {
                Stop-Process -Id $process.Id -Force
                throw "LGPO実行がタイムアウトしました（${LGPOTimeoutSeconds}秒）"
            }
            $exitCode = $process.ExitCode
            if ($exitCode -ne 0) {
                throw "LGPO解析に失敗しました。ExitCode=$exitCode"
            }
            #>
            $job = Start-Job -ScriptBlock {
                param($lgpoPath, $polPath)
                Start-Process -FilePath $lgpoPath -ArgumentList "/parse", "/m", $polPath -Wait -PassThru -NoNewWindow -RedirectStandardOutput $env:TEMP\lgpo_out.txt -RedirectStandardError $env:TEMP\lgpo_err.txt
            } -ArgumentList $LGPOPath, $path

            if (Wait-Job $job -Timeout 30) {
                $process = Receive-Job $job
                Remove-Job $job -ErrorAction SilentlyContinue
                if ($process.ExitCode -ne 0) {
                    $errorContent = Get-Content $env:TEMP\lgpo_err.txt -ErrorAction SilentlyContinue
                    # Write-Log -Level "ERROR" -Code "E01003" -Message "LGPO解析に失敗しました: $errorContent"
                    $ErrorDetails.Value = "LGPO解析に失敗しました: $errorContent"
                    return "E01003"
                }                               
            } else {
                Stop-Job $job -ErrorAction SilentlyContinue
                Remove-Job $job -Force -ErrorAction SilentlyContinue
                $ErrorDetails.Value = "LGPO実行がタイムアウトしました（ $LGPOTimeoutSeconds 秒）"
                return "E01003"
            }
            
        } catch {
            # Write-Log -Level "ERROR" -Code "E01003" -Message "LGPO実行異常: $($_.Exception.Message)"
            $ErrorDetails.Value = "LGPO実行異常: $($_.Exception.Message)"
            return "E01003"
        } finally {
            # 一時ファイルの削除
            Remove-Item $env:TEMP\lgpo_out.txt, $env:TEMP\lgpo_err.txt -ErrorAction SilentlyContinue
        }
    }

    return "I00000"  # 正常
}

# --- ファイル操作ラッパー ---
function Invoke-FileAction {
    param(
        [string]$Action,   # "Get" / "Remove" / "Copy"
        [hashtable]$Params # コマンドに渡される引数
    )
    try {
        $result = $null
        switch ($Action) {
            "Get"    { $result = Get-ChildItem @Params }
            "Remove" { Remove-Item @Params -ErrorAction Stop; $result = $true }
            "Copy"   { Copy-Item @Params -ErrorAction Stop; $result = $true }
            default  { throw "未知操作: $Action" }
        }

        $msg = "ファイル操作成功 - $Action"

        # 成功時は INFO レベルで返す
        return [pscustomobject]@{
            Result  = $result
            ErrCode = "I00000"
            Level   = "INFO"
            Message = $msg
        }
    } catch {
        switch ($Action) {
            "Copy"   { $errCode = "W01001"; $level = "WARN" }
            "Remove" { $errCode = "W01002"; $level = "WARN" }
            "Get"    { $errCode = "W01003"; $level = "WARN" }
            default  { $errCode = "E09995"; $level = "ERROR" }
        }
        $msg = "ファイル操作エラー - $Action : $($_.Exception.Message)"
        # Write-Log -Level $level -Code $errCode -Message $msg

        <#
        # 通知は ERROR のときだけ
        if ($level -eq "ERROR") {
            Invoke-Notify -Level "ERROR" -Message $msg
        }
        #>
        return [pscustomobject]@{
            Result  = $null
            ErrCode = $errCode
            Level   = $level
            Message = $msg
        }
    }
}

try {
	# ==== 初期化 ====
	foreach ($d in @($BackupDir, $LogDir)) {
		if (-not (Test-Path $d)) {
            try {
                New-Item -ItemType Directory -Path $d -Force -ErrorAction Stop | Out-Null
                Write-Log -Level "INFO" -Code "I00000" -Message "ディレクトリを作成しました: $d"
            } catch {
                $msg = "ディレクトリの作成に失敗しました: $d - $($_.Exception.Message)"
                Invoke-Notify -Level "ERROR" -Message $msg
                Write-Log -Level "ERROR" -Code "E09994" -Message $msg
                throw "必須ディレクトリの作成に失敗したため、スクリプトを終了します。"
            }
		}
	}

	# === 処理対象 ===
	$polFiles = @{
		"Machine" = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"
	}

	# ==== メイン処理 ====
	# 日時情報
	$currentTime = Get-Date
	$timestamp = $currentTime.ToString("yyyyMMdd_HHmmss")
	$today = $currentTime.ToString("yyyyMMdd")

	# ログファイル（日単位）
	$script:LogFile = Join-Path $LogDir ("RegistryPol_{0}.log" -f $currentTime.ToString("yyyyMMdd"))

	# 問題検出リスト
	$problem   = @()

	# 検査開始ログ
	Write-Log -Level "INFO" -Code "I00000" -Message "検査開始"

	# --- 各POLファイル検査 ---
	foreach ($key in $polFiles.Keys) {
		$polPath = $polFiles[$key]
		$errorDetails = ""
		$code = Test-PolFile $polPath ([ref]$errorDetails)

		if ($code -eq "I00000") {
			Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol 検査実施 → 正常"

			# --- バックアップ判定 ---
			$todayBackup = Get-ChildItem $BackupDir -File -Filter "*${today}*${key}_Registry.pol" -ErrorAction SilentlyContinue

			if (-not $todayBackup) {
				$name = "${timestamp}_${key}_Registry.pol"
				$backupPath = Join-Path $BackupDir $name
				$result = Invoke-FileAction -Action "Copy" -Params @{ Path = $polPath; Destination = $backupPath; Force = $true }
				if ($result.Level -eq "INFO") {
					$msg = "$key Registry.pol バックアップ処理 → 成功: $backupPath"
					Write-Log -Level "INFO" -Code "I00000" -Message $msg
				} elseif ($result.Level -eq "WARN") {
					$msg = "$key Registry.pol バックアップ処理 → 失敗: $($result.Message)"
					Write-Log -Level "WARN" -Code $result.ErrCode -Message $msg
					Invoke-Notify -Level "WARN" -Message $msg   
				} elseif ($result.Level -eq "ERROR") {
					Write-Log -Level "ERROR" -Code $result.ErrCode -Message $result.Message
					Invoke-Notify -Level "ERROR" -Message $result.Message
				}

			} else {
				$msg = "$key Registry.pol バックアップ処理 → 本日分は既に存在するためスキップ: $($todayBackup.FullName)"
				Write-Log -Level "INFO" -Code $code -Message $msg
			}
		} else {
			$desc = switch ($code) {
				"E01001" { "ファイル存在しません" }
				"E01002" { "サイズ0" }
				"E01003" { $errorDetails.Value }
				default  { "未知のエラー" }
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
			if ($result.Level -eq "INFO") {
				$msg = "古いバックアップ削除 → $($_.FullName)"
				Write-Log -Level "INFO" -Code "I00000" -Message $msg
			} elseif ($result.Level -eq "WARN") {
				$msg = "古いバックアップ削除 → 失敗: $($result.Message)"
				Write-Log -Level $result.Level -Code $result.ErrCode -Message $msg
				Invoke-Notify -Level $result.Level -Message $msg
			}
		}

	# --- 古いログ削除 ---
	Get-ChildItem $LogDir -File -Filter "RegistryPol_*.log" |
		Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
		ForEach-Object {
			$result = Invoke-FileAction -Action "Remove" -Params @{ Path = $_.FullName; Force = $true }
			if ($result.Level -eq "INFO") {
				$msg = "古いログ削除 → $($_.FullName)"
				Write-Log -Level "INFO" -Code "I00000" -Message $msg
			} elseif ($result.Level -eq "WARN") {
				$msg = "古いログ削除 → 失敗: $($result.Message)"
				Write-Log -Level $result.Level -Code $result.ErrCode -Message $msg
				Invoke-Notify -Level $result.Level -Message $msg
			}
		}

	# --- エラー通知 ---
	if ($problem.Count -gt 0) {
		$msg = "Registry.pol ファイルに問題を検出: " + ($problem -join ", ")
		Write-Log -Level "ERROR" -Code "E09998" -Message $msg
		Invoke-Notify -Level "ERROR" -Message $msg
	}
}
catch {
    $msg = "スクリプト実行中に予期しないエラーが発生しました: $($_.Exception.Message)"
    Invoke-Notify -Level "ERROR" -Message $msg
    Write-Log -Level "ERROR" -Code "E09999" -Message $msg
    exit 1
} finally {
    if ($script:LogFile) {
        Write-Log -Level "INFO" -Code "I00000" -Message "スクリプト実行完了"
    }
}

