param(
    [string]$BackupDir = "C:\work\RegistryPolCheck\Backup",
    [int]$MaxBackup = 10,
    [string]$LogDir = "C:\work\RegistryPolCheck\Logs",
    [int]$LogRetentionDays = 7,
    [string]$NotifyBat = "C:\Scripts\TexpostNotify.bat",
    [string]$LGPOPath = "C:\Tools\LGPO.exe"   # �C��: LGPO.exe �̃p�X
)

# ==== ������ ====
foreach ($d in @($BackupDir, $LogDir)) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d | Out-Null
    }
}

# ���O�t�@�C���i���P�ʁj
$logFile = Join-Path $LogDir ("RegistryPol_{0}.log" -f (Get-Date -Format "yyyyMMdd"))

# ���O�o�͊֐�
function Write-Log {
    param([string]$msg)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Add-Content -Path $logFile -Value $line
    Write-Output $line
}

# POL�t�@�C�������֐��i�G���[�R�[�h��Ԃ��j
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
    "User"    = "$env:SystemRoot\System32\GroupPolicy\User\Registry.pol"
}

$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$today     = (Get-Date).ToString("yyyyMMdd")
$problem   = @()

foreach ($key in $polFiles.Keys) {
    $polPath = $polFiles[$key]
    $code = Test-PolFile $polPath

    if ($code -eq "E00000") {
        # --- �o�b�N�A�b�v���� ---
        $todayBackup = Get-ChildItem $BackupDir -File -Filter "*${today}*${key}_Registry.pol" -ErrorAction SilentlyContinue
        if (-not $todayBackup) {
            $name = "${timestamp}_${key}_Registry.pol"
            $backupPath = Join-Path $BackupDir $name
            Copy-Item $polPath $backupPath -Force
            Write-Log "$key Registry.pol ���� (E00000)�A�o�b�N�A�b�v�쐬: $backupPath"
        } else {
            Write-Log "$key Registry.pol ���� (E00000)�A�{���̃o�b�N�A�b�v�͊��ɑ���"
        }
    } else {
        # --- �G���[���o ---
        switch ($code) {
            "E01001" { $desc = "�t�@�C���s����" }
            "E01002" { $desc = "�T�C�Y0" }
            "E01003" { $desc = "��͎��s" }
            default  { $desc = "���m�̃G���[" }
        }
        Write-Log "�x��: $key Registry.pol �ُ�A��� = $code ($desc)"
        $problem += "$key: $code ($desc)"
    }
}

# --- �Â��o�b�N�A�b�v�폜 ---
Get-ChildItem $BackupDir -File -Filter "*Registry.pol" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip $MaxBackup |
    Remove-Item -Force

# --- �Â����O�폜 ---
Get-ChildItem $LogDir -File -Filter "RegistryPol_*.log" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
    Remove-Item -Force

# --- �G���[�ʒm ---
if ($problem.Count -gt 0) {
    $msg = "Registry.pol �t�@�C���ɖ������o: " + ($problem -join ", ")
    Write-Log $msg

    if (Test-Path $NotifyBat) {
        Write-Log "�O���ʒm�o�b�`���s: $NotifyBat"
        & $NotifyBat $msg
    } else {
        Write-Log "�ʒm�o�b�`�����݂��܂���: $NotifyBat"
    }
}
