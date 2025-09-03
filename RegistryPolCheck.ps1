# RegistryPolCheck.ps1
# �����T�v�F �{�X�N���v�g�� Windows �T�[�o�[��ŗ��p����郍�[�J���O���[�v�|���V�[�ݒ�t�@�C�� Registry.pol �̌��S�����m�F���A�ُ�����m�����ꍇ�Ƀ��O�o�͂���ђʒm���s���c�[���ł��B
# �܂��A����ȏꍇ�͓����Ńo�b�N�A�b�v���擾���A�Â��o�b�N�A�b�v�⃍�O�������I�ɍ폜���܂��B
# �ݒ�\�ȃp�����[�^�[�𗘗p���āA�o�b�N�A�b�v�ۑ����⃍�O�ۑ����ԁA�ʒm�o�b�`�̃p�X�Ȃǂ��J�X�^�}�C�Y�ł��܂��B
# LGPO.exe �𗘗p���� Registry.pol �̓��e����͂��A�ُ킪�Ȃ������m�F���܂��B
# ������ LGPO.exe �͕K�{�ł͂Ȃ��A���݂��Ȃ��ꍇ�̓t�@�C���̑��݂ƃT�C�Y�݂̂��`�F�b�N���܂��B
# LGPO.exe �� Microsoft ���疳���Œ񋟂���Ă���A�K�v�ɉ����ē������Ă��������B
# https://www.microsoft.com/en-us/download/details.aspx?id=55319
# �Ȃ��ALGPO.exe �𗘗p����ꍇ�́A�X�N���v�g���̊Y�������̃R�����g�A�E�g���������Ă��������B

# �@�\�T�v�F
# - Registry.pol �t�@�C���̑��݊m�F�ƃT�C�Y�`�F�b�N
# - LGPO.exe �𗘗p�������e��́i�I�v�V�����j
# - �����o�b�N�A�b�v�̎擾�i�d���h�~�j
# - �Â��o�b�N�A�b�v�ƃ��O�̎����폜
# - ���O�t�@�C���ւ̏ڍׂȃ��O�o�́iINFO / ERROR ���x���j
# - �ُ팟�m���̊O���ʒm�o�b�`�̎��s

# �G���[���x���ƃR�[�h��F
# - I00000: ����I��
# - W01001: �o�b�N�A�b�v���s
# - W01002: �Â��o�b�N�A�b�v�폜���s
# - W01003: �Â����O�폜���s
# - W01004: �Ǘ��Ҍ����Ȃ��x��
# - E01001: �t�@�C�����݂��܂���
# - E01002: �T�C�Y0
# - E01003: ��͎��s
# - E09994: �K�{�f�B���N�g���쐬���s
# - E09995: �s���ȃt�@�C������G���[
# - E09996: �ʒm�o�b�`���s�ُ�
# - E09997: �ʒm�o�b�`�����݂��Ȃ�
# - E09998: Registry.pol �ɖ�茟�o
# - E09999: �\�����Ȃ��G���[

# ���s���@�F
# - �X�N���v�g: powershell.exe -ExecutionPolicy RemoteSigned -File "C:\work\RegistryPolCheck\RegistryPolCheck.ps1"

# --- �p�����[�^�[��` ---
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
    [string]$LGPOPath = "C:\Tools\LGPO.exe"   # �C��: LGPO.exe �̃p�X
)

# --- �O���[�o���ϐ� ---
$script:LogFile = $null  # ���O�t�@�C���p�X�i���s���ɐݒ肳���j
$NotifyTimeoutSeconds = 30
$LGPOTimeoutSeconds   = 30

# --- �Ǘ��Ҍ����`�F�b�N ---
# �Ǘ��Ҍ������Ȃ��ꍇ�͌x�����O���o�͂��邪�A�����͌p������
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    $msg = "���̃X�N���v�g�͊Ǘ��Ҍ����Ŏ��s���邱�Ƃ𐄏����܂��B�����s���ɂ��ꕔ�̑��삪���s����\��������܂��B"
    # Invoke-Notify -Level "WARN" -Message $msg
    # Write-Log -Level "WARN" -Code "W01004" -Message $msg
}

# �O���[�o���G���[�����ݒ�
$ErrorActionPreference = "Continue"  # �G���[�����������ꍇ�ł����s���p������
$ProgressPreference = "SilentlyContinue"  # �^�X�N�X�P�W���[���ł̕\������������邽�߁A�v���O���X�o�[�𖳌������܂��B

# --- ���O�o�͊֐� ---
function Write-Log {
    param(
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level,
        [string]$Code,
        [string]$Message
    )
    
    if (-not $script:LogFile) {
        # throw "LogFile�ϐ�������������Ă��܂���"
        Write-Output "[FALLBACK-LOG] $Message"
        return        
    }
    
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    $line = "$timestamp [$Level] [$Code] $Message"
    
    try {
        Add-Content -Path $script:LogFile -Value $line -ErrorAction Stop
        Write-Output $line
    } catch {
        Write-Error "���O�t�@�C���ւ̏������݂��ł��܂���: $_"
    }
}

# --- �ʒm�o�b�`���s�֐� ---
function Invoke-Notify {
    param(
        [ValidateSet("WARN", "ERROR", "MINOR")]
        [string]$Level,   # "ERROR" or "WARN"
        [string]$Message
    )

    # Level �� s�l �̃}�b�s���O
    $map = @{
        "WARN"  = "1"        
        "ERROR" = "2"
        "MINOR" = "3"
    }
    $s = $map[$Level]
    if (-not $s) { $s = "3" }  # fallback: ����`Level��MINOR����

    if (Test-Path $NotifyBat) {
		Write-Log -Level "INFO" -Code "I00000" -Message "�O���ʒm�o�b�`���s: $NotifyBat (���x���F $Level)"
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
                    Write-Log -Level "ERROR" -Code "E09996" -Message "�ʒm�o�b�`�̎��s�Ɏ��s���܂����B�߂�l�ُ�: ExitCode=$exitCode | $NotifyBat �i���x���F $Level�j"
                } else {
                    Write-Log -Level "INFO" -Code "I00000" -Message "�ʒm�o�b�`�̎��s�ɐ������܂����F $NotifyBat �i���x���F $Level�j"
                }

            } else {
                Stop-Job $job -ErrorAction SilentlyContinue
                Remove-Job $job -Force -ErrorAction SilentlyContinue
                Write-Log -Level "ERROR" -Code "E09996" -Message "�ʒm�o�b�`�̎��s���^�C���A�E�g���܂����i $NotifyTimeoutSeconds �b�j: $NotifyBat"
            }         
        }
        catch {
            Write-Log -Level "ERROR" -Code "E09996" -Message "�ʒm�o�b�`���s���ɗ�O������: $($_.Exception.Message)"
        }
    } else {
        Write-Log -Level "ERROR" -Code "E09997" -Message "�O���ʒm�o�b�`�����݂��܂���: $NotifyBat"
    }
}

# --- POL�t�@�C�������֐� ---
function Test-PolFile {
    param($path, [ref]$ErrorDetails)

    if (-not (Test-Path $path)) { return "E01001" }   # �t�@�C���Ȃ�
    $size = (Get-Item $path).Length
    if ($size -eq 0) { return "E01002" }              # �T�C�Y0

    # --- LGPO.exe �̗��p�ɂ��� ---
    # LGPO.exe �� Microsoft �񋟂� Local Group Policy Object Utility �ł���A
    # Registry.pol �̓��e����͂��邽�߂Ɏg�p�\�B
    # �������{�T�[�o�[�� LGPO.exe �����݂��Ȃ��\�������邽�߁A
    # �ꎞ�I�ɗ��p�������R�����g�A�E�g���Ă���B
    # �����I�ɓ��e�`�F�b�N���s���ꍇ�́ALGPO.exe �𓱓����ēx�L�������邱�ƁB
    if ($LGPOPath -and (Test-Path $LGPOPath)) {
        try {
            <#
            $process = Start-Process -FilePath $LGPOPath -ArgumentList "/parse /q /m $polPath" -PassThru
            if (-not (Wait-Process -Id $process.Id -Timeout $LGPOTimeoutSeconds -ErrorAction SilentlyContinue)) {
                Stop-Process -Id $process.Id -Force
                throw "LGPO���s���^�C���A�E�g���܂����i${LGPOTimeoutSeconds}�b�j"
            }
            $exitCode = $process.ExitCode
            if ($exitCode -ne 0) {
                throw "LGPO��͂Ɏ��s���܂����BExitCode=$exitCode"
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
                    # Write-Log -Level "ERROR" -Code "E01003" -Message "LGPO��͂Ɏ��s���܂���: $errorContent"
                    $ErrorDetails.Value = "LGPO��͂Ɏ��s���܂���: $errorContent"
                    return "E01003"
                }                               
            } else {
                Stop-Job $job -ErrorAction SilentlyContinue
                Remove-Job $job -Force -ErrorAction SilentlyContinue
                $ErrorDetails.Value = "LGPO���s���^�C���A�E�g���܂����i $LGPOTimeoutSeconds �b�j"
                return "E01003"
            }
            
        } catch {
            # Write-Log -Level "ERROR" -Code "E01003" -Message "LGPO���s�ُ�: $($_.Exception.Message)"
            $ErrorDetails.Value = "LGPO���s�ُ�: $($_.Exception.Message)"
            return "E01003"
        } finally {
            # �ꎞ�t�@�C���̍폜
            Remove-Item $env:TEMP\lgpo_out.txt, $env:TEMP\lgpo_err.txt -ErrorAction SilentlyContinue
        }
    }

    return "I00000"  # ����
}

# --- �t�@�C�����색�b�p�[ ---
function Invoke-FileAction {
    param(
        [string]$Action,   # "Get" / "Remove" / "Copy"
        [hashtable]$Params # �R�}���h�ɓn��������
    )
    try {
        $result = $null
        switch ($Action) {
            "Get"    { $result = Get-ChildItem @Params }
            "Remove" { Remove-Item @Params -ErrorAction Stop; $result = $true }
            "Copy"   { Copy-Item @Params -ErrorAction Stop; $result = $true }
            default  { throw "���m����: $Action" }
        }

        $msg = "�t�@�C�����쐬�� - $Action"

        # �������� INFO ���x���ŕԂ�
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
        $msg = "�t�@�C������G���[ - $Action : $($_.Exception.Message)"
        # Write-Log -Level $level -Code $errCode -Message $msg

        <#
        # �ʒm�� ERROR �̂Ƃ�����
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
	# ==== ������ ====
	foreach ($d in @($BackupDir, $LogDir)) {
		if (-not (Test-Path $d)) {
            try {
                New-Item -ItemType Directory -Path $d -Force -ErrorAction Stop | Out-Null
                Write-Log -Level "INFO" -Code "I00000" -Message "�f�B���N�g�����쐬���܂���: $d"
            } catch {
                $msg = "�f�B���N�g���̍쐬�Ɏ��s���܂���: $d - $($_.Exception.Message)"
                Invoke-Notify -Level "ERROR" -Message $msg
                Write-Log -Level "ERROR" -Code "E09994" -Message $msg
                throw "�K�{�f�B���N�g���̍쐬�Ɏ��s�������߁A�X�N���v�g���I�����܂��B"
            }
		}
	}

	# === �����Ώ� ===
	$polFiles = @{
		"Machine" = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"
	}

	# ==== ���C������ ====
	# �������
	$currentTime = Get-Date
	$timestamp = $currentTime.ToString("yyyyMMdd_HHmmss")
	$today = $currentTime.ToString("yyyyMMdd")

	# ���O�t�@�C���i���P�ʁj
	$script:LogFile = Join-Path $LogDir ("RegistryPol_{0}.log" -f $currentTime.ToString("yyyyMMdd"))

	# ��茟�o���X�g
	$problem   = @()

	# �����J�n���O
	Write-Log -Level "INFO" -Code "I00000" -Message "�����J�n"

	# --- �ePOL�t�@�C������ ---
	foreach ($key in $polFiles.Keys) {
		$polPath = $polFiles[$key]
		$errorDetails = ""
		$code = Test-PolFile $polPath ([ref]$errorDetails)

		if ($code -eq "I00000") {
			Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol �������{ �� ����"

			# --- �o�b�N�A�b�v���� ---
			$todayBackup = Get-ChildItem $BackupDir -File -Filter "*${today}*${key}_Registry.pol" -ErrorAction SilentlyContinue

			if (-not $todayBackup) {
				$name = "${timestamp}_${key}_Registry.pol"
				$backupPath = Join-Path $BackupDir $name
				$result = Invoke-FileAction -Action "Copy" -Params @{ Path = $polPath; Destination = $backupPath; Force = $true }
				if ($result.Level -eq "INFO") {
					$msg = "$key Registry.pol �o�b�N�A�b�v���� �� ����: $backupPath"
					Write-Log -Level "INFO" -Code "I00000" -Message $msg
				} elseif ($result.Level -eq "WARN") {
					$msg = "$key Registry.pol �o�b�N�A�b�v���� �� ���s: $($result.Message)"
					Write-Log -Level "WARN" -Code $result.ErrCode -Message $msg
					Invoke-Notify -Level "WARN" -Message $msg   
				} elseif ($result.Level -eq "ERROR") {
					Write-Log -Level "ERROR" -Code $result.ErrCode -Message $result.Message
					Invoke-Notify -Level "ERROR" -Message $result.Message
				}

			} else {
				$msg = "$key Registry.pol �o�b�N�A�b�v���� �� �{�����͊��ɑ��݂��邽�߃X�L�b�v: $($todayBackup.FullName)"
				Write-Log -Level "INFO" -Code $code -Message $msg
			}
		} else {
			$desc = switch ($code) {
				"E01001" { "�t�@�C�����݂��܂���" }
				"E01002" { "�T�C�Y0" }
				"E01003" { $errorDetails.Value }
				default  { "���m�̃G���[" }
			}
			Write-Log -Level "ERROR" -Code $code -Message "$key Registry.pol �������{ �� �ُ�A��� = $desc"
			$problem += "$($key): $code ($desc)"
		}
	}

	# �����I�����O
	Write-Log -Level "INFO" -Code "I00000" -Message "�����I��"

	# --- �Â��o�b�N�A�b�v�폜 ---
	Get-ChildItem $BackupDir -File -Filter "*Registry.pol" |
		Sort-Object LastWriteTime -Descending |
		Select-Object -Skip $MaxBackup |
		ForEach-Object {
			$result = Invoke-FileAction -Action "Remove" -Params @{ Path = $_.FullName; Force = $true }
			if ($result.Level -eq "INFO") {
				$msg = "�Â��o�b�N�A�b�v�폜 �� $($_.FullName)"
				Write-Log -Level "INFO" -Code "I00000" -Message $msg
			} elseif ($result.Level -eq "WARN") {
				$msg = "�Â��o�b�N�A�b�v�폜 �� ���s: $($result.Message)"
				Write-Log -Level $result.Level -Code $result.ErrCode -Message $msg
				Invoke-Notify -Level $result.Level -Message $msg
			}
		}

	# --- �Â����O�폜 ---
	Get-ChildItem $LogDir -File -Filter "RegistryPol_*.log" |
		Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
		ForEach-Object {
			$result = Invoke-FileAction -Action "Remove" -Params @{ Path = $_.FullName; Force = $true }
			if ($result.Level -eq "INFO") {
				$msg = "�Â����O�폜 �� $($_.FullName)"
				Write-Log -Level "INFO" -Code "I00000" -Message $msg
			} elseif ($result.Level -eq "WARN") {
				$msg = "�Â����O�폜 �� ���s: $($result.Message)"
				Write-Log -Level $result.Level -Code $result.ErrCode -Message $msg
				Invoke-Notify -Level $result.Level -Message $msg
			}
		}

	# --- �G���[�ʒm ---
	if ($problem.Count -gt 0) {
		$msg = "Registry.pol �t�@�C���ɖ������o: " + ($problem -join ", ")
		Write-Log -Level "ERROR" -Code "E09998" -Message $msg
		Invoke-Notify -Level "ERROR" -Message $msg
	}
}
catch {
    $msg = "�X�N���v�g���s���ɗ\�����Ȃ��G���[���������܂���: $($_.Exception.Message)"
    Invoke-Notify -Level "ERROR" -Message $msg
    Write-Log -Level "ERROR" -Code "E09999" -Message $msg
    exit 1
} finally {
    if ($script:LogFile) {
        Write-Log -Level "INFO" -Code "I00000" -Message "�X�N���v�g���s����"
    }
}

