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
# - W01002: 古いバックアップ／ログ削除失敗
# - W01003: ファイル取得失敗
# - W01004: 管理者権限なし警告
# - W09997: 互換性警告（LGPO解析で互換性問題検出）
# - E01001: ファイル存在しません
# - E01002: サイズ0
# - E01003: 解析失敗（LGPO実行がタイムアウトまたは異常終了）
# - E01004: アクセス拒否/権限不足（LGPOがアクセス拒否で失敗）
# - E01005: 破損ファイル（LGPOがファイル破損で失敗）
# - E09993: 必須ディレクトリ作成失敗
# - E09994: 不明なファイル操作エラー
# - E09995: 通知バッチ実行異常
# - E09996: 通知バッチが存在しない
# - E09997: Registry.pol に問題検出
# - E09998: 未知のエラー（LGPO解析で未知のエラー検出）
# - E09999: LGPO実行異常（LGPOプロセスの起動失敗）
# - E99999: 予期しないエラー

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

# Set-PSDebug -Trace 2

# 厳密モードの設定
Set-StrictMode -Version 3.0 

# --- グローバル変数 ---
$script:LogFile = $null  # ログファイルパス（実行時に設定される）
$NotifyTimeoutSeconds = 30
$LGPOTimeoutSeconds   = 30

# グローバルエラー処理設定
$ErrorActionPreference = "Continue"  # エラーが発生した場合でも実行を継続する
$ProgressPreference = "SilentlyContinue"  # タスクスケジューラでの表示問題を回避するため、プログレスバーを無効化する

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
        $msg = "ログファイルへの書き込みができません: $_"
        Write-Output $msg
    }
}

# --- 通知バッチ実行関数 ---
function Invoke-Notify {
    param(
        [ValidateSet("WARN", "ERROR", "MINOR")]
        [string]$Level,   # 通知レベル（WARN/ERROR/MINOR）
        [string]$Message
    )

    # Level → -s値 のマッピング
    $map = @{
        "WARN"  = "1"        
        "ERROR" = "2"
        "MINOR" = "3"
    }
    $s = $map[$Level]
    if (-not $s) { $s = "3" }  # fallback: 未定義LevelはMINOR扱い

    if (Test-Path $NotifyBat) {
        # 通知バッチ実行開始ログ
		Write-Log -Level "INFO" -Code "I00000" -Message "外部通知バッチ実行: $NotifyBat (レベル： $Level)"
        try {
            # 引数の組み立て
            $notifyArgs = @("-s", $s, "-i", "`"A $Message`"", "-c", "kibt")

            # Start-Process方式で実行（非同期起動＋PID取得）
            $process = Start-Process -FilePath $NotifyBat -ArgumentList $notifyArgs -PassThru -WindowStyle Hidden
            
			# タイムアウト付きで待機（Wait-Processは終了を待つだけなので終了確認を追加）
			$waited = Wait-Process -Id $process.Id -Timeout $NotifyTimeoutSeconds -ErrorAction SilentlyContinue

			if (-not (Get-Process -Id $process.Id -ErrorAction SilentlyContinue)) {
                # プロセス終了済み → $waited が $false なら「Wait-Process が取りこぼした」
                if (-not $waited) {
                    Write-Log -Level "INFO" -Code "I09998" -Message "Wait-Process が false を返しましたが、実際にはプロセスは終了していました: $NotifyBat"
                }
			    # --- プロセスは既に終了している ---
			    $exitCode = $process.ExitCode
			    if ($exitCode -ne 0) {
			        Write-Log -Level "ERROR" -Code "E09995" -Message "通知バッチの実行に失敗しました。戻り値異常: ExitCode=$exitCode | $NotifyBat （レベル： $Level）"
			    } else {
			        Write-Log -Level "INFO" -Code "I00000" -Message "通知バッチの実行に成功しました： $NotifyBat （レベル： $Level）"
			    }
			} else {
			    # --- プロセスはまだ生きている = 本当にタイムアウト ---
			    try {
			        Stop-Process -Id $process.Id -Force -ErrorAction Stop
			        Write-Log -Level "WARN" -Code "W09999" -Message "タイムアウトした通知バッチを強制終了しました: PID=$($process.Id)"
			    } catch {
			        Write-Log -Level "INFO" -Code "I09999" -Message "通知バッチ終了処理: Stop-Process 実行時に例外発生（既に終了している可能性あり）: $($_.Exception.Message)"
			    }
			    Write-Log -Level "ERROR" -Code "E09995" -Message "通知バッチの実行がタイムアウトしました（$NotifyTimeoutSeconds 秒）: $NotifyBat"
			}
   
        }
        catch {
            # --- 例外処理 ---
            Write-Log -Level "ERROR" -Code "E09995" -Message "通知バッチ実行中に例外が発生: $($_.Exception.Message)"
        }
    } else {
        # --- バッチファイルが存在しない場合 ---
        Write-Log -Level "ERROR" -Code "E09996" -Message "外部通知バッチが存在しません: $NotifyBat"
    }
}

# --- POLファイル検査関数 ---
function Test-PolFile {
    param($path, [ref]$ErrorDetails)

    # 事前初期化（Set-StrictMode を使っている場合に必須）
    $code = $null

    # 存在チェック    
    if (-not (Test-Path $path)) {
        $code = "E01001"
        return $code
    }

    # サイズチェック
    $size = (Get-Item $path).Length
    if ($size -eq 0) {
        $code = "E01002"
        return $code
    }

    # --- LGPO.exe の利用について ---
    # LGPO.exe は Microsoft 提供の Local Group Policy Object Utility であり、
    # Registry.pol の内容を解析するために使用可能。
    # ただし本サーバーに LGPO.exe が存在しない可能性があるため、
    # 一時的に利用部分をコメントアウトしている。
    # 将来的に内容チェックを行う場合は、LGPO.exe を導入し再度有効化すること。

    # LGPO 実行用ファイルパスが指定されており存在する場合のみ実行
    if (-not ($LGPOPath) -or -not (Test-Path $LGPOPath)) {
        # LGPO 未指定または存在しない -> 内容チェックスキップして正常扱い
        $code = "I00000"
        return $code
    }

    # LGPO.exe 実行 (timeout付き)
    $outFile = Join-Path $env:TEMP "lgpo_out.txt"
    $errFile = Join-Path $env:TEMP "lgpo_err.txt"    
	try {
        # Start-Process で非同期起動し、プロセスオブジェクトを取得
        $proc = Start-Process -FilePath $LGPOPath `
            -ArgumentList "/parse","/m",$path `
            -RedirectStandardOutput $outFile `
            -RedirectStandardError  $errFile `
            -NoNewWindow -PassThru -ErrorAction Stop
        <#
        # PowerShell 7.0 以降で利用可能な Wait-Process
		if (-not (Wait-Process -Id $proc.Id -Timeout $LGPOTimeoutSeconds -ErrorAction SilentlyContinue)) {
		    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
		    $ErrorDetails.Value = "LGPO実行がタイムアウトしました（${LGPOTimeoutSeconds}秒）"
		    return "E01003"
		}
        #>

        # .NET の WaitForExit を利用（単位はミリ秒）
        $ok = $proc.WaitForExit($LGPOTimeoutSeconds * 1000)
        if (-not $ok) {
            # --- ② LGPO 実行異常（タイムアウト / プロセスハング） ---
            $code = "E01003"
            $msg = "LGPO実行がタイムアウトしました（${LGPOTimeoutSeconds}秒）"
            Write-Error $msg
            try {
                Stop-Process -Id $proc.Id -Force -ErrorAction Stop
            } catch {
                $stopMsg = "LGPOプロセスの強制終了に失敗しました: $($_.Exception.Message)"
                Write-Error $stopMsg
                $msg = "$msg | $stopMsg"
            }
            $ErrorDetails.Value = $msg
            return $code
        }

        # --- LGPO の終了コードを確認 ---
	    if ($proc.ExitCode -ne 0) {
            $errText = ""
            # エラーファイルの内容を確認
            if (Test-Path $errFile) {
                $errText = (Get-Content -Path $errFile -Raw -ErrorAction SilentlyContinue).Trim()
            }

            # 互換性（古いバージョンなど）による解析不可の判定：
			if ($errText) {
			    # 改行をスペースに置換して1行にまとめ、長すぎる場合は切り詰める
			    $singleLine = ($errText -replace "\r?\n", " ").Trim()
			    if ($singleLine.Length -gt 1000) {
			        $singleLine = $singleLine.Substring(0,1000) + "..."
			    }

			    switch -regex ($singleLine) {
                    # 互換性警告
			        "parse|解析|invalid" {
			            $code = "W09997"
			            $ErrorDetails.Value = "LGPO互換性失敗（Registry.polを解析できませんでした）: $singleLine"
			            break
			        }
                    # 権限エラー
			        "access|denied|拒否|permission" {
			            $code = "E01004"
			            $ErrorDetails.Value = "LGPO権限エラー（アクセス拒否/権限不足）: $singleLine"
			            break
			        }
                    # ファイル破損
			        "corrupt|壊れている" {
			            $code = "E01005"
			            $ErrorDetails.Value = "LGPOファイル破損（Registry.polが壊れています）: $singleLine"
			            break
			        }
                    # その他のエラー
			        default {
			            $code = "E09998"   # 未知のエラー
			            $ErrorDetails.Value = "LGPO不明エラー: $singleLine"
			        }
			    }
			}
			else {
			    # 正常（errText が空なら OK）
			    $code = "I00000"
			}

	    }
	}
	catch {
	    # --- ② プロセス起動自体が失敗 ---
        $code = "E09999"
        $msg = "LGPOプロセスの起動に失敗しました: $($_.Exception.Message)"
	    $ErrorDetails.Value = $msg
	    return $code
	}
	finally {
	    # --- ③ エラーファイルの扱いと ErrorDetails の整備 ---
	    # $code が未定義なら安全のため初期化
	    if (-not $code) { $code = "E09998" }

	    # outFile は常に削除
	    if (Test-Path $outFile) {
	        Remove-Item $outFile -ErrorAction SilentlyContinue
	    }

	    # errFile の扱い
	    if ($code -eq "I00000" -or $code -eq "W09997") {
	        # 正常 or 警告なら errFile も削除
	        if (Test-Path $errFile) {
	            Remove-Item $errFile -ErrorAction SilentlyContinue
	        }
	    }
	    else {
	        # 致命エラー時は errFile を残す
	        if ((Test-Path $errFile) -and (-not $ErrorDetails.Value)) {
	            $errText = (Get-Content -Path $errFile -Raw -ErrorAction SilentlyContinue).Trim()
	            if ($errText) {
	                $singleLine = ($errText -replace "\r?\n", " ").Trim()
	                if ($singleLine.Length -gt 1500) { $singleLine = $singleLine.Substring(0,1500) + "..." }
	                $ErrorDetails.Value = $singleLine
	            }
	        }
	    }

	    # 正常時のみ一時ファイルを削除
	    if ($code -eq "I00000") {
	        Remove-Item $outFile,$errFile -ErrorAction SilentlyContinue
	    } else {
	        # エラー/警告時は errFile があれば ErrorDetails に詰める（ログ出力は main に任せる）
	        if ((Test-Path $errFile) -and (-not $ErrorDetails.Value)) {
	            $errText = (Get-Content -Path $errFile -Raw -ErrorAction SilentlyContinue).Trim()
	            if ($errText) {
	                # 改行をスペースに変換して1行にまとめる（長すぎるときは切り詰め）
	                $singleLine = ($errText -replace "\r?\n", " ").Trim()
	                if ($singleLine.Length -gt 1500) { $singleLine = $singleLine.Substring(0,1500) + "..." }
	                $ErrorDetails.Value = $singleLine
	            }
	        }
	        # errFile はデバッグのため残す（必要なら別途クリーンアップポリシーを作る）
	    }
	}
   
    # --- 正常終了 ---
    if (-not $code) {
        $code = "I00000"
    }
    return $code
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
            default  { $errCode = "E09994"; $level = "ERROR" }
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

# --- メイン処理 ---
try {
    # まず基本的なログファイルのパスを設定します（一時的）
    $tempLogFile = Join-Path $env:TEMP "RegistryPolCheck_Init.log"
    $script:LogFile = $tempLogFile
    Write-Log -Level "INFO" -Code "I00000" -Message "スクリプト実行開始"

    # --- 管理者権限チェック ---
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        $msg = "管理者権限で実行されていません。正常に動作しない可能性があります。"
        Write-Log -Level "WARN" -Code "W01004" -Message $msg
        Invoke-Notify -Level "WARN" -Message $msg
    }

	# ==== 初期化 ====
	foreach ($d in @($BackupDir, $LogDir)) {
		if (-not (Test-Path $d)) {
            try {
                New-Item -ItemType Directory -Path $d -Force -ErrorAction Stop | Out-Null
                Write-Log -Level "INFO" -Code "I00000" -Message "ディレクトリを作成しました: $d"
            } catch {
                $msg = "ディレクトリの作成に失敗しました: $d - $($_.Exception.Message)"
                Write-Log -Level "ERROR" -Code "E09993" -Message $msg
                Invoke-Notify -Level "ERROR" -Message $msg
                # throw "必須ディレクトリの作成に失敗したため、スクリプトを終了します。"
                exit 1
            }
		}
	}

	# === 処理対象 ===
	$polFiles = @{
		"Machine" = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"
	}

	# 日時情報
	$currentTime = Get-Date
	$timestamp = $currentTime.ToString("yyyyMMdd_HHmmss")
	$today = $currentTime.ToString("yyyyMMdd")

	# ログファイル（日単位）
	$script:LogFile = Join-Path $LogDir ("RegistryPol_{0}.log" -f $currentTime.ToString("yyyyMMdd"))

    # 一時ログがある場合は、正式ログファイルにコピーする
    if (Test-Path $tempLogFile) {
        Get-Content $tempLogFile | Add-Content $script:LogFile
        Remove-Item $tempLogFile -ErrorAction SilentlyContinue
    }    

	# 問題検出リスト
	$problem   = @()

	# 検査開始ログ
	Write-Log -Level "INFO" -Code "I00000" -Message "検査開始"

	# --- 各POLファイル検査 ---
	foreach ($key in $polFiles.Keys) {
		$polPath = $polFiles[$key]
        $errorDetails = [ref]("")
        $code = Test-PolFile $polPath $errorDetails


		if (($code -eq "I00000") -or ($code -eq "W09997")) {
		    # --- 正常または警告ありで処理継続 ---
		    if ($code -eq "W09997") {
		        Write-Log -Level "WARN" -Code $code -Message "$key Registry.pol 検査実施 → 警告、状態 = $($errorDetails.Value)"
                Write-Log -Level "INFO" -Code "I00000" -Message "$key Registry.pol 検査実施 → 正常（互換性警告あり）"
		    } else {
		        Write-Log -Level "INFO" -Code "I00000" -Message "$key Registry.pol 検査実施 → 正常"
		    }

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
                # 配列に強制変換し、最初の要素を取得する
                $backupPath = @($todayBackup)[0].FullName
				$msg = "$key Registry.pol バックアップ処理 → 本日分は既に存在するためスキップ: $($backupPath)"
				Write-Log -Level "INFO" -Code "I00000" -Message $msg
			}
		} else {
			$desc = switch ($code) {
				"E01001" { "ファイル存在しません" }
				"E01002" { "サイズ0" }
                "E01003" { "解析失敗: $($ErrorDetails.Value)" }
                "E01004" { "アクセス拒否/権限不足: $($ErrorDetails.Value)" }
                "E01005" { "破損ファイル: $($ErrorDetails.Value)" }
                "E09998" { "未知のエラー: $($ErrorDetails.Value)" }
                "E09999" { "LGPO実行異常: $($ErrorDetails.Value)" }
                default { "不明コード ($code): $($ErrorDetails.Value)" }
			}
            # エラーログ出力
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
		Write-Log -Level "ERROR" -Code "E09997" -Message $msg
		Invoke-Notify -Level "ERROR" -Message $msg
	}
}
catch {
    $msg = "スクリプト実行中に予期しないエラーが発生しました: $($_.Exception.Message)"
    Write-Log -Level "ERROR" -Code "E99999" -Message $msg    
    Invoke-Notify -Level "ERROR" -Message $msg
    exit 1   # タスクスケジューラが失敗と判定できるように、非ゼロ終了コードで終了する
} finally {
    # 一時ログ (Init用) を削除する
    Remove-Item $tempLogFile -ErrorAction SilentlyContinue

    # スクリプト完了ログ
    if ($script:LogFile) {
        Write-Log -Level "INFO" -Code "I00000" -Message "スクリプト実行完了"
    }
}

