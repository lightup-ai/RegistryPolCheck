param(
    [string]$BackupDir = "C:\work\RegistryPolCheck\Backup",
    [int]$MaxBackup = 10,
    [string]$LogDir = "C:\work\RegistryPolCheck\Logs",
    [int]$LogRetentionDays = 7,
    [string]$NotifyBat = "C:\Scripts\TexpostNotify.bat",
    [string]$LGPOPath = "C:\Tools\LGPO.exe"   # �C��: LGPO.exe �̃p�X
)

# --- ���O�o�͊֐� ---
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

# --- �t�@�C�����색�b�p�[ ---
function Invoke-FileAction {
    param(
        [string]$Action,   # "Get" / "Remove" / "Copy"
        [hashtable]$Params # ???���ߓI�Q��
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
            "Copy"   { $errCode = "E02001" }
            "Remove" { $errCode = "E02002" }
            "Get"    { $errCode = "E02003" }
            default  { $errCode = "E09997" }
        }
        Write-Log -Level "ERROR" -Code $errCode -Message "$Action ���s���s - $($_.Exception.Message)"
        return $null
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

# --- POL�t�@�C�������֐� ---
function Test-PolFile {
    param($path)

    if (-not (Test-Path $path)) { return "E01001" }   # �t�@�C���Ȃ�
    $size = (Get-Item $path).Length
    if ($size -eq 0) { return "E01002" }              # �T�C�Y0

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
    return "E00000"                                   # ����
}

# === �����Ώ� ===
$polFiles = @{
    "Machine" = "$env:SystemRoot\System32\GroupPolicy\Machine\Registry.pol"
}

$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$today     = (Get-Date).ToString("yyyyMMdd")
$problem   = @()

# �����J�n���O
Write-Log -Level "INFO" -Code "E00000" -Message "�����J�n"

# --- �ePOL�t�@�C������ ---
foreach ($key in $polFiles.Keys) {
    $polPath = $polFiles[$key]
    $code = Test-PolFile $polPath

    if ($code -eq "E00000") {
        Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol �������{ �� ����"

        # --- �o�b�N�A�b�v���� ---
        $todayBackup = Get-ChildItem $BackupDir -File -Filter "*${today}*${key}_Registry.pol" -ErrorAction SilentlyContinue

        if (-not $todayBackup) {
            $name = "${timestamp}_${key}_Registry.pol"
            $backupPath = Join-Path $BackupDir $name
            Invoke-FileAction -Action "Copy" -Params @{ Path = $polPath; Destination = $backupPath; Force = $true }
            Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol �o�b�N�A�b�v���� �� �쐬: $backupPath"
        } else {
            Write-Log -Level "INFO" -Code $code -Message "$key Registry.pol �o�b�N�A�b�v���� �� �{���̃o�b�N�A�b�v�͊��ɑ���"
        }
    } else {
        switch ($code) {
            "E01001" { $desc = "�t�@�C���s����" }
            "E01002" { $desc = "�T�C�Y0" }
            "E01003" { $desc = "��͎��s" }
            default  { $desc = "���m�̃G���[" }
        }
        Write-Log -Level "ERROR" -Code $code -Message "$key Registry.pol �������{ �� �ُ�A��� = $desc"
        $problem += "$($key): $code ($desc)"
    }
}

# �����I�����O
Write-Log -Level "INFO" -Code "E00000" -Message "�����I��"

# --- �Â��o�b�N�A�b�v�폜 ---
Get-ChildItem $BackupDir -File -Filter "*Registry.pol" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip $MaxBackup |
    ForEach-Object {
        Invoke-FileAction -Action "Remove" -Params @{ Path = $_.FullName; Force = $true }
        Write-Log -Level "INFO" -Code "E00000" -Message "�Â��o�b�N�A�b�v�폜 �� $($_.FullName)"
    }

# --- �Â����O�폜 ---
Get-ChildItem $LogDir -File -Filter "RegistryPol_*.log" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
    ForEach-Object {
        Invoke-FileAction -Action "Remove" -Params @{ Path = $_.FullName; Force = $true }
        Write-Log -Level "INFO" -Code "E00000" -Message "�Â����O�폜 �� $($_.FullName)"
    }

# --- �G���[�ʒm ---
if ($problem.Count -gt 0) {
    $msg = "Registry.pol �t�@�C���ɖ������o: " + ($problem -join ", ")
    Write-Log -Level "ERROR" -Code "E09999" -Message $msg

    if (Test-Path $NotifyBat) {
        Write-Log -Level "INFO" -Code "E00000" -Message "�O���ʒm�o�b�`���s: $NotifyBat"
        & $NotifyBat $msg
    } else {
        Write-Log -Level "ERROR" -Code "E09998" -Message "�ʒm�o�b�`�����݂��܂���: $NotifyBat"
    }
}
