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
# - W01002: �Â��o�b�N�A�b�v�^���O�폜���s
# - W01003: �t�@�C���擾���s
# - W01004: �Ǘ��Ҍ����Ȃ��x��
# - W09997: �݊����x���iLGPO��͂Ō݊�����茟�o�j
# - E01001: �t�@�C�����݂��܂���
# - E01002: �T�C�Y0
# - E01003: ��͎��s�iLGPO���s���^�C���A�E�g�܂��ُ͈�I���j
# - E01004: �A�N�Z�X����/�����s���iLGPO���A�N�Z�X���ۂŎ��s�j
# - E01005: �j���t�@�C���iLGPO���t�@�C���j���Ŏ��s�j
# - E09993: �K�{�f�B���N�g���쐬���s
# - E09994: �s���ȃt�@�C������G���[
# - E09995: �ʒm�o�b�`���s�ُ�
# - E09996: �ʒm�o�b�`�����݂��Ȃ�
# - E09997: Registry.pol �ɖ�茟�o
# - E09998: ���m�̃G���[�iLGPO��͂Ŗ��m�̃G���[���o�j
# - E09999: LGPO���s�ُ�iLGPO�v���Z�X�̋N�����s�j
# - E99999: �\�����Ȃ��G���[

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

# Set-PSDebug -Trace 2

# �������[�h�̐ݒ�
Set-StrictMode -Version 3.0 

# --- �O���[�o���ϐ� ---
$script:LogFile = $null  # ���O�t�@�C���p�X�i���s���ɐݒ肳���j
$NotifyTimeoutSeconds = 30
$LGPOTimeoutSeconds   = 30

# �O���[�o���G���[�����ݒ�
$ErrorActionPreference = "Continue"  # �G���[�����������ꍇ�ł����s���p������
$ProgressPreference = "SilentlyContinue"  # �^�X�N�X�P�W���[���ł̕\������������邽�߁A�v���O���X�o�[�𖳌�������

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
        $msg = "���O�t�@�C���ւ̏������݂��ł��܂���: $_"
        Write-Output $msg
    }
}

# --- �ʒm�o�b�`���s�֐� ---
function Invoke-Notify {
    param(
        [ValidateSet("WARN", "ERROR", "MINOR")]
        [string]$Level,   # �ʒm���x���iWARN/ERROR/MINOR�j
        [string]$Message
    )

    # Level �� -s�l �̃}�b�s���O
    $map = @{
        "WARN"  = "1"        
        "ERROR" = "2"
        "MINOR" = "3"
    }
    $s = $map[$Level]
    if (-not $s) { $s = "3" }  # fallback: ����`Level��MINOR����

    if (Test-Path $NotifyBat) {
        # �ʒm�o�b�`���s�J�n���O
		Write-Log -Level "INFO" -Code "I00000" -Message "�O���ʒm�o�b�`���s: $NotifyBat (���x���F $Level)"
        try {
            # �����̑g�ݗ���
            $notifyArgs = @("-s", $s, "-i", "`"A $Message`"", "-c", "kibt")

            # Start-Process�����Ŏ��s�i�񓯊��N���{PID�擾�j
            $process = Start-Process -FilePath $NotifyBat -ArgumentList $notifyArgs -PassThru -WindowStyle Hidden
            
			# �^�C���A�E�g�t���őҋ@�iWait-Process�͏I����҂����Ȃ̂ŏI���m�F��ǉ��j
			$waited = Wait-Process -Id $process.Id -Timeout $NotifyTimeoutSeconds -ErrorAction SilentlyContinue

			if (-not (Get-Process -Id $process.Id -ErrorAction SilentlyContinue)) {
                # �v���Z�X�I���ς� �� $waited �� $false �Ȃ�uWait-Process ����肱�ڂ����v
                if (-not $waited) {
                    Write-Log -Level "INFO" -Code "I09998" -Message "Wait-Process �� false ��Ԃ��܂������A���ۂɂ̓v���Z�X�͏I�����Ă��܂���: $NotifyBat"
                }
			    # --- �v���Z�X�͊��ɏI�����Ă��� ---
			    $exitCode = $process.ExitCode
			    if ($exitCode -ne 0) {
			        Write-Log -Level "ERROR" -Code "E09995" -Message "�ʒm�o�b�`�̎��s�Ɏ��s���܂����B�߂�l�ُ�: ExitCode=$exitCode | $NotifyBat �i���x���F $Level�j"
			    } else {
			        Write-Log -Level "INFO" -Code "I00000" -Message "�ʒm�o�b�`�̎��s�ɐ������܂����F $NotifyBat �i���x���F $Level�j"
			    }
			} else {
			    # --- �v���Z�X�͂܂������Ă��� = �{���Ƀ^�C���A�E�g ---
			    try {
			        Stop-Process -Id $process.Id -Force -ErrorAction Stop
			        Write-Log -Level "WARN" -Code "W09999" -Message "�^�C���A�E�g�����ʒm�o�b�`�������I�����܂���: PID=$($process.Id)"
			    } catch {
			        Write-Log -Level "INFO" -Code "I09999" -Message "�ʒm�o�b�`�I������: Stop-Process ���s���ɗ�O�����i���ɏI�����Ă���\������j: $($_.Exception.Message)"
			    }
			    Write-Log -Level "ERROR" -Code "E09995" -Message "�ʒm�o�b�`�̎��s���^�C���A�E�g���܂����i$NotifyTimeoutSeconds �b�j: $NotifyBat"
			}
   
        }
        catch {
            # --- ��O���� ---
            Write-Log -Level "ERROR" -Code "E09995" -Message "�ʒm�o�b�`���s���ɗ�O������: $($_.Exception.Message)"
        }
    } else {
        # --- �o�b�`�t�@�C�������݂��Ȃ��ꍇ ---
        Write-Log -Level "ERROR" -Code "E09996" -Message "�O���ʒm�o�b�`�����݂��܂���: $NotifyBat"
    }
}

# --- POL�t�@�C�������֐� ---
function Test-PolFile {
    param($path, [ref]$ErrorDetails)

    # ���O�������iSet-StrictMode ���g���Ă���ꍇ�ɕK�{�j
    $code = $null

    # ���݃`�F�b�N    
    if (-not (Test-Path $path)) {
        $code = "E01001"
        return $code
    }

    # �T�C�Y�`�F�b�N
    $size = (Get-Item $path).Length
    if ($size -eq 0) {
        $code = "E01002"
        return $code
    }

    # --- LGPO.exe �̗��p�ɂ��� ---
    # LGPO.exe �� Microsoft �񋟂� Local Group Policy Object Utility �ł���A
    # Registry.pol �̓��e����͂��邽�߂Ɏg�p�\�B
    # �������{�T�[�o�[�� LGPO.exe �����݂��Ȃ��\�������邽�߁A
    # �ꎞ�I�ɗ��p�������R�����g�A�E�g���Ă���B
    # �����I�ɓ��e�`�F�b�N���s���ꍇ�́ALGPO.exe �𓱓����ēx�L�������邱�ƁB

    # LGPO ���s�p�t�@�C���p�X���w�肳��Ă��葶�݂���ꍇ�̂ݎ��s
    if (-not ($LGPOPath) -or -not (Test-Path $LGPOPath)) {
        # LGPO ���w��܂��͑��݂��Ȃ� -> ���e�`�F�b�N�X�L�b�v���Đ��툵��
        $code = "I00000"
        return $code
    }

    # LGPO.exe ���s (timeout�t��)
    $outFile = Join-Path $env:TEMP "lgpo_out.txt"
    $errFile = Join-Path $env:TEMP "lgpo_err.txt"    
	try {
        # Start-Process �Ŕ񓯊��N�����A�v���Z�X�I�u�W�F�N�g���擾
        $proc = Start-Process -FilePath $LGPOPath `
            -ArgumentList "/parse","/m",$path `
            -RedirectStandardOutput $outFile `
            -RedirectStandardError  $errFile `
            -NoNewWindow -PassThru -ErrorAction Stop
        <#
        # PowerShell 7.0 �ȍ~�ŗ��p�\�� Wait-Process
		if (-not (Wait-Process -Id $proc.Id -Timeout $LGPOTimeoutSeconds -ErrorAction SilentlyContinue)) {
		    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
		    $ErrorDetails.Value = "LGPO���s���^�C���A�E�g���܂����i${LGPOTimeoutSeconds}�b�j"
		    return "E01003"
		}
        #>

        # .NET �� WaitForExit �𗘗p�i�P�ʂ̓~���b�j
        $ok = $proc.WaitForExit($LGPOTimeoutSeconds * 1000)
        if (-not $ok) {
            # --- �A LGPO ���s�ُ�i�^�C���A�E�g / �v���Z�X�n���O�j ---
            $code = "E01003"
            $msg = "LGPO���s���^�C���A�E�g���܂����i${LGPOTimeoutSeconds}�b�j"
            Write-Error $msg
            try {
                Stop-Process -Id $proc.Id -Force -ErrorAction Stop
            } catch {
                $stopMsg = "LGPO�v���Z�X�̋����I���Ɏ��s���܂���: $($_.Exception.Message)"
                Write-Error $stopMsg
                $msg = "$msg | $stopMsg"
            }
            $ErrorDetails.Value = $msg
            return $code
        }

        # --- LGPO �̏I���R�[�h���m�F ---
	    if ($proc.ExitCode -ne 0) {
            $errText = ""
            # �G���[�t�@�C���̓��e���m�F
            if (Test-Path $errFile) {
                $errText = (Get-Content -Path $errFile -Raw -ErrorAction SilentlyContinue).Trim()
            }

            # �݊����i�Â��o�[�W�����Ȃǁj�ɂ���͕s�̔���F
			if ($errText) {
			    # ���s���X�y�[�X�ɒu������1�s�ɂ܂Ƃ߁A��������ꍇ�͐؂�l�߂�
			    $singleLine = ($errText -replace "\r?\n", " ").Trim()
			    if ($singleLine.Length -gt 1000) {
			        $singleLine = $singleLine.Substring(0,1000) + "..."
			    }

			    switch -regex ($singleLine) {
                    # �݊����x��
			        "parse|���|invalid" {
			            $code = "W09997"
			            $ErrorDetails.Value = "LGPO�݊������s�iRegistry.pol����͂ł��܂���ł����j: $singleLine"
			            break
			        }
                    # �����G���[
			        "access|denied|����|permission" {
			            $code = "E01004"
			            $ErrorDetails.Value = "LGPO�����G���[�i�A�N�Z�X����/�����s���j: $singleLine"
			            break
			        }
                    # �t�@�C���j��
			        "corrupt|���Ă���" {
			            $code = "E01005"
			            $ErrorDetails.Value = "LGPO�t�@�C���j���iRegistry.pol�����Ă��܂��j: $singleLine"
			            break
			        }
                    # ���̑��̃G���[
			        default {
			            $code = "E09998"   # ���m�̃G���[
			            $ErrorDetails.Value = "LGPO�s���G���[: $singleLine"
			        }
			    }
			}
			else {
			    # ����ierrText ����Ȃ� OK�j
			    $code = "I00000"
			}

	    }
	}
	catch {
	    # --- �A �v���Z�X�N�����̂����s ---
        $code = "E09999"
        $msg = "LGPO�v���Z�X�̋N���Ɏ��s���܂���: $($_.Exception.Message)"
	    $ErrorDetails.Value = $msg
	    return $code
	}
	finally {
	    # --- �B �G���[�t�@�C���̈����� ErrorDetails �̐��� ---
	    # $code ������`�Ȃ���S�̂��ߏ�����
	    if (-not $code) { $code = "E09998" }

	    # outFile �͏�ɍ폜
	    if (Test-Path $outFile) {
	        Remove-Item $outFile -ErrorAction SilentlyContinue
	    }

	    # errFile �̈���
	    if ($code -eq "I00000" -or $code -eq "W09997") {
	        # ���� or �x���Ȃ� errFile ���폜
	        if (Test-Path $errFile) {
	            Remove-Item $errFile -ErrorAction SilentlyContinue
	        }
	    }
	    else {
	        # �v���G���[���� errFile ���c��
	        if ((Test-Path $errFile) -and (-not $ErrorDetails.Value)) {
	            $errText = (Get-Content -Path $errFile -Raw -ErrorAction SilentlyContinue).Trim()
	            if ($errText) {
	                $singleLine = ($errText -replace "\r?\n", " ").Trim()
	                if ($singleLine.Length -gt 1500) { $singleLine = $singleLine.Substring(0,1500) + "..." }
	                $ErrorDetails.Value = $singleLine
	            }
	        }
	    }

	    # ���펞�݈̂ꎞ�t�@�C�����폜
	    if ($code -eq "I00000") {
	        Remove-Item $outFile,$errFile -ErrorAction SilentlyContinue
	    } else {
	        # �G���[/�x������ errFile ������� ErrorDetails �ɋl�߂�i���O�o�͂� main �ɔC����j
	        if ((Test-Path $errFile) -and (-not $ErrorDetails.Value)) {
	            $errText = (Get-Content -Path $errFile -Raw -ErrorAction SilentlyContinue).Trim()
	            if ($errText) {
	                # ���s���X�y�[�X�ɕϊ�����1�s�ɂ܂Ƃ߂�i��������Ƃ��͐؂�l�߁j
	                $singleLine = ($errText -replace "\r?\n", " ").Trim()
	                if ($singleLine.Length -gt 1500) { $singleLine = $singleLine.Substring(0,1500) + "..." }
	                $ErrorDetails.Value = $singleLine
	            }
	        }
	        # errFile �̓f�o�b�O�̂��ߎc���i�K�v�Ȃ�ʓr�N���[���A�b�v�|���V�[�����j
	    }
	}
   
    # --- ����I�� ---
    if (-not $code) {
        $code = "I00000"
    }
    return $code
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
            default  { $errCode = "E09994"; $level = "ERROR" }
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

# --- ���C������ ---
try {
    # �܂���{�I�ȃ��O�t�@�C���̃p�X��ݒ肵�܂��i�ꎞ�I�j
    $tempLogFile = Join-Path $env:TEMP "RegistryPolCheck_Init.log"
    $script:LogFile = $tempLogFile
    Write-Log -Level "INFO" -Code "I00000" -Message "�X�N���v�g���s�J�n"

    # --- �Ǘ��Ҍ����`�F�b�N ---
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $isAdmin) {
        $msg = "�Ǘ��Ҍ����Ŏ��s����Ă��܂���B����ɓ��삵�Ȃ��\��������܂��B"
        Write-Log -Level "WARN" -Code "W01004" -Message $msg
        Invoke-Notify -Level "WARN" -Message $msg
    }

	# ==== ������ ====
	foreach ($d in @($BackupDir, $LogDir)) {
		if (-not (Test-Path $d)) {
            try {
                New-Item -ItemType Directory -Path $d -Force -ErrorAction Stop | Out-Null
                Write-Log -Level "INFO" -Code "I00000" -Message "�f�B���N�g�����쐬���܂���: $d"
            } catch {
                $msg = "�f�B���N�g���̍쐬�Ɏ��s���܂���: $d - $($_.Exception.Message)"
                Write-Log -Level "ERROR" -Code "E09993" -Message $msg
                Invoke-Notify -Level "ERROR" -Message $msg
                # throw "�K�{�f�B���N�g���̍쐬�Ɏ��s�������߁A�X�N���v�g���I�����܂��B"
                exit 1
            }
		}
	}

	# === �����Ώ� ===
	$polFiles = @{
		"Machine" = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"
	}

	# �������
	$currentTime = Get-Date
	$timestamp = $currentTime.ToString("yyyyMMdd_HHmmss")
	$today = $currentTime.ToString("yyyyMMdd")

	# ���O�t�@�C���i���P�ʁj
	$script:LogFile = Join-Path $LogDir ("RegistryPol_{0}.log" -f $currentTime.ToString("yyyyMMdd"))

    # �ꎞ���O������ꍇ�́A�������O�t�@�C���ɃR�s�[����
    if (Test-Path $tempLogFile) {
        Get-Content $tempLogFile | Add-Content $script:LogFile
        Remove-Item $tempLogFile -ErrorAction SilentlyContinue
    }    

	# ��茟�o���X�g
	$problem   = @()

	# �����J�n���O
	Write-Log -Level "INFO" -Code "I00000" -Message "�����J�n"

	# --- �ePOL�t�@�C������ ---
	foreach ($key in $polFiles.Keys) {
		$polPath = $polFiles[$key]
        $errorDetails = [ref]("")
        $code = Test-PolFile $polPath $errorDetails


		if (($code -eq "I00000") -or ($code -eq "W09997")) {
		    # --- ����܂��͌x������ŏ����p�� ---
		    if ($code -eq "W09997") {
		        Write-Log -Level "WARN" -Code $code -Message "$key Registry.pol �������{ �� �x���A��� = $($errorDetails.Value)"
                Write-Log -Level "INFO" -Code "I00000" -Message "$key Registry.pol �������{ �� ����i�݊����x������j"
		    } else {
		        Write-Log -Level "INFO" -Code "I00000" -Message "$key Registry.pol �������{ �� ����"
		    }

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
                # �z��ɋ����ϊ����A�ŏ��̗v�f���擾����
                $backupPath = @($todayBackup)[0].FullName
				$msg = "$key Registry.pol �o�b�N�A�b�v���� �� �{�����͊��ɑ��݂��邽�߃X�L�b�v: $($backupPath)"
				Write-Log -Level "INFO" -Code "I00000" -Message $msg
			}
		} else {
			$desc = switch ($code) {
				"E01001" { "�t�@�C�����݂��܂���" }
				"E01002" { "�T�C�Y0" }
                "E01003" { "��͎��s: $($ErrorDetails.Value)" }
                "E01004" { "�A�N�Z�X����/�����s��: $($ErrorDetails.Value)" }
                "E01005" { "�j���t�@�C��: $($ErrorDetails.Value)" }
                "E09998" { "���m�̃G���[: $($ErrorDetails.Value)" }
                "E09999" { "LGPO���s�ُ�: $($ErrorDetails.Value)" }
                default { "�s���R�[�h ($code): $($ErrorDetails.Value)" }
			}
            # �G���[���O�o��
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
		Write-Log -Level "ERROR" -Code "E09997" -Message $msg
		Invoke-Notify -Level "ERROR" -Message $msg
	}
}
catch {
    $msg = "�X�N���v�g���s���ɗ\�����Ȃ��G���[���������܂���: $($_.Exception.Message)"
    Write-Log -Level "ERROR" -Code "E99999" -Message $msg    
    Invoke-Notify -Level "ERROR" -Message $msg
    exit 1   # �^�X�N�X�P�W���[�������s�Ɣ���ł���悤�ɁA��[���I���R�[�h�ŏI������
} finally {
    # �ꎞ���O (Init�p) ���폜����
    Remove-Item $tempLogFile -ErrorAction SilentlyContinue

    # �X�N���v�g�������O
    if ($script:LogFile) {
        Write-Log -Level "INFO" -Code "I00000" -Message "�X�N���v�g���s����"
    }
}

