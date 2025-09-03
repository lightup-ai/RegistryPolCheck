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
# - E01001: �t�@�C�����݂��܂���
# - E01002: �T�C�Y0
# - E01003: ��͎��s
# - E02001: �o�b�N�A�b�v���s
# - E02002: �Â��o�b�N�A�b�v�폜���s
# - E02003: �Â����O�폜���s
# - E09996: �s���ȃt�@�C������G���[
# - E09997: �ʒm�o�b�`���s���s
# - E09998: �ʒm�o�b�`�����݂��Ȃ�
# - E09999: ���̑��̖��m�̃G���[

# ���s���@�F
# - �X�N���v�g: powershell.exe -ExecutionPolicy RemoteSigned -File "C:\work\RegistryPolCheck\RegistryPolCheck.ps1"

param(
    [string]$BackupDir = "C:\work\RegistryPolCheck\Backup",
    [int]$MaxBackup = 7,
    [string]$LogDir = "C:\work\RegistryPolCheck\Logs",
    [int]$LogRetentionDays = 7,
    [string]$NotifyBat = "C:\Scripts\TexpostNotify.bat",
    [string]$LGPOPath = "C:\Tools\LGPO.exe"   # �C��: LGPO.exe �̃p�X
)

# --- ���O�o�͊֐� ---
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

# --- POL�t�@�C�������֐� ---
function Test-PolFile {
    param($path)

    if (-not (Test-Path $path)) { return "E01001" }   # �t�@�C���Ȃ�
    $size = (Get-Item $path).Length
    if ($size -eq 0) { return "E01002" }              # �T�C�Y0

    # --- LGPO.exe �̗��p�ɂ��� ---
    # LGPO.exe �� Microsoft �񋟂� Local Group Policy Object Utility �ł���A
    # Registry.pol �̓��e����͂��邽�߂Ɏg�p�\�B
    # �������{�T�[�o�[�� LGPO.exe �����݂��Ȃ��\�������邽�߁A
    # �ꎞ�I�ɗ��p�������R�����g�A�E�g���Ă���B
    # �����I�ɓ��e�`�F�b�N���s���ꍇ�́ALGPO.exe �𓱓����ēx�L�������邱�ƁB
    if (Test-Path $LGPOPath) {
        try {
            $out = & $LGPOPath /parse /m $path 2>&1
            if ($LASTEXITCODE -ne 0 -or $out -match "Error") {
                return "E01003"                       # ��͎��s
            }
        } catch {
            return "E01003"                           # ��͎��s
        }
    }
    return "I00000"                                   # ����
}

# --- �t�@�C�����색�b�p�[ ---
function Invoke-FileAction {
    param(
        [string]$Action,   # "Get" / "Remove" / "Copy"
        [hashtable]$Params # �R�}���h�ɓn��������
    )
    try {
        switch ($Action) {
            "Get"    { return Get-ChildItem @Params }
            "Remove" { Remove-Item @Params -ErrorAction Stop }
            "Copy"   { Copy-Item @Params -ErrorAction Stop }
            default  { throw "���m����: $Action" }
        }
    } catch {
        switch ($Action) {
            "Copy"   { $errCode = "W01001"; $level = "WARN" }
            "Remove" { $errCode = "W01002"; $level = "WARN" }
            "Get"    { $errCode = "W01003"; $level = "WARN" }
            default  { $errCode = "E09996"; $level = "ERROR" }
        }
        $msg = "�t�@�C������G���[ - $Action : $($_.Exception.Message)"
        Write-Log -Level $level -Code $errCode -Message $msg

        # �ʒm�� ERROR �̂Ƃ�����
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

# ==== ������ ====
foreach ($d in @($BackupDir, $LogDir)) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d | Out-Null
    }
}

# ���O�t�@�C���i���P�ʁj
$logFile = Join-Path $LogDir ("RegistryPol_{0}.log" -f (Get-Date -Format "yyyyMMdd"))

# === �����Ώ� ===
$polFiles = @{
    "Machine" = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"
}

$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$today     = (Get-Date).ToString("yyyyMMdd")
$problem   = @()

# �����J�n���O
Write-Log -Level "INFO" -Code "I00000" -Message "�����J�n"

# --- �ePOL�t�@�C������ ---
foreach ($key in $polFiles.Keys) {
    $polPath = $polFiles[$key]
    $code = Test-PolFile $polPath

    if ($code -eq "I00000") {
        Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol �������{ �� ����"

        # --- �o�b�N�A�b�v���� ---
        $todayBackup = Get-ChildItem $BackupDir -File -Filter "*${today}*${key}_Registry.pol" -ErrorAction SilentlyContinue

        if (-not $todayBackup) {
            $name = "${timestamp}_${key}_Registry.pol"
            $backupPath = Join-Path $BackupDir $name
            $result = Invoke-FileAction -Action "Copy" -Params @{ Path = $polPath; Destination = $backupPath; Force = $true }
            if ($result -is [pscustomobject] -and $result.ErrCode) {
                # �G���[����
                if ($result.Level -eq "WARN") {
                    $msg = "�o�b�N�A�b�v���� �� ���s: $($result.Message)"
                    Write-Log -Level $result.Level -Code $result.ErrCode -Message $msg
                    $notifyArgs = "-s", "1", "-i", "`"A $msg`"", "-c", "kibt"
                    & $NotifyBat @notifyArgs                    
                }
            } else {
                Write-Log -Level "INFO" -Code "I00000" -Message "$key Registry.pol �o�b�N�A�b�v���� �� �쐬: $backupPath"
            }
        } else {
            Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol �o�b�N�A�b�v���� �� �{���̃o�b�N�A�b�v�͊��ɑ���"
        }
    } else {
        switch ($code) {
            "E01001" { $desc = "�t�@�C�����݂��܂���" }
            "E01002" { $desc = "�T�C�Y0" }
            "E01003" { $desc = "LGPO.exe �ɂ�� Registry.pol ��͂Ɏ��s���܂���" }
            default  { $desc = "���m�̃G���[" }
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
        if ($result -is [pscustomobject] -and $result.ErrCode) {
            # �G���[����
            if ($result.Level -eq "WARN") {
                $msg = "�Â��o�b�N�A�b�v�폜 �� ���s: $($result.Message)"
                Write-Log -Level $result.Level -Code $result.ErrCode -Message $msg
                $notifyArgs = "-s", "1", "-i", "`"A $msg`"", "-c", "kibt"
                & $NotifyBat @notifyArgs
            }
        } else {
            Write-Log -Level "INFO" -Code "I00000" -Message "�Â��o�b�N�A�b�v�폜 �� $($_.FullName)"
        }
    }

# --- �Â����O�폜 ---
Get-ChildItem $LogDir -File -Filter "RegistryPol_*.log" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
    ForEach-Object {
        $result = Invoke-FileAction -Action "Remove" -Params @{ Path = $_.FullName; Force = $true }
        if ($result -is [pscustomobject] -and $result.ErrCode) {
            # �G���[����
            if ($result.Level -eq "WARN") {
                $msg = "�Â����O�폜 �� ���s: $($result.Message)"
                Write-Log -Level $result.Level -Code $result.ErrCode -Message $msg
                $notifyArgs = "-s", "1", "-i", "`"A $msg`"", "-c", "kibt"
                & $NotifyBat @notifyArgs
            }
        } else {
            Write-Log -Level "INFO" -Code "I00000" -Message "�Â����O�폜 �� $($_.FullName)"
        }
    }

# --- �G���[�ʒm ---
if ($problem.Count -gt 0) {
    $msg = "Registry.pol �t�@�C���ɖ������o: " + ($problem -join ", ")
    Write-Log -Level "ERROR" -Code "E09999" -Message $msg

    if (Test-Path $NotifyBat) {
        Write-Log -Level "INFO" -Code "I00000" -Message "�O���ʒm�o�b�`���s: $NotifyBat"
        $notifyArgs = "-s", "2", "-i", "`"A $msg`"", "-c", "kibt"
        & $NotifyBat @notifyArgs
        if ($LASTEXITCODE -ne 0) {
            Write-Log -Level "ERROR" -Code "E09997" -Message "�ʒm�o�b�`�̎��s�Ɏ��s���܂����B�߂�l�ُ�: ExitCode=$LASTEXITCODE"
        } else {
            Write-Log -Level "INFO" -Code "I00000" -Message "�ʒm�o�b�`�̎��s�ɐ������܂����F $NotifyBat �i���x���F MINOR�j"
        }
    } else {
        Write-Log -Level "ERROR" -Code "E09998" -Message "�O���ʒm�o�b�`�����݂��܂���: $NotifyBat"
    }
}
