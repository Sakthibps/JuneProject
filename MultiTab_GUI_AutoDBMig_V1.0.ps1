Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# ================= FORM =================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Enterprise SQL Admin Console"
$form.Size = New-Object System.Drawing.Size(950,720)
$form.StartPosition = "CenterScreen"
$form.KeyPreview = $true

$form.Add_KeyDown({ if ($_.KeyCode -eq "Escape") { $form.Close() } })

# ================= TABS =================
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Size="900,650"; $tabs.Location="10,10"

# Tabs
$tabConn   = New-Object System.Windows.Forms.TabPage -Property @{Text="Source Connection"}
$tabTarget = New-Object System.Windows.Forms.TabPage -Property @{Text="Target Connection"}
$tabDB     = New-Object System.Windows.Forms.TabPage -Property @{Text="Database Explorer"}
$tabBackup = New-Object System.Windows.Forms.TabPage -Property @{Text="Backup"}
$tabCopy   = New-Object System.Windows.Forms.TabPage -Property @{Text="Copy Backup"}
$tabRestore= New-Object System.Windows.Forms.TabPage -Property @{Text="Restore"}
$tabImort= New-Object System.Windows.Forms.TabPage -Property @{Text="Imort CSV\JSON"}

# ================= STATUS LABELS =================
$lblConn = New-Object System.Windows.Forms.Label -Property @{Location="20,280";Width=800}
$lblTgt  = New-Object System.Windows.Forms.Label -Property @{Location="20,280";Width=800}
$lblBkp  = New-Object System.Windows.Forms.Label -Property @{Location="20,180";Width=800}
$lblCopy = New-Object System.Windows.Forms.Label -Property @{Location="20,250";Width=800}
$lblRst  = New-Object System.Windows.Forms.Label -Property @{Location="20,220";Width=800}
$lblExpo  = New-Object System.Windows.Forms.Label -Property @{Location="20,220";Width=800}

# ================= STATUS FUNCTIONS =================
function Set-ConnStatus($m,$c="Blue"){ $lblConn.Text="$(Get-Date): $m"; $lblConn.ForeColor=$c }
function Set-TgtStatus($m,$c="Blue"){ $lblTgt.Text="$(Get-Date): $m"; $lblTgt.ForeColor=$c }
function Set-BkpStatus($m,$c="Blue"){ $lblBkp.Text="$(Get-Date): $m"; $lblBkp.ForeColor=$c }
function Set-CopyStatus($m,$c="Blue"){ $lblCopy.Text="$(Get-Date): $m"; $lblCopy.ForeColor=$c }
function Set-RstStatus($m,$c="Blue"){ $lblRst.Text="$(Get-Date): $m"; $lblRst.ForeColor=$c }
#function Set-ExpoStatus($m,$c="Blue"){ $lblRst.Text="$(Get-Date): $m"; $lblExpo.ForeColor=$c }

# ================= CONNECTION BUILDERS =================
function Get-Source {
    $conn = $txtServer.Text.Trim()
    if ($txtInst.Text){ $conn = "$conn\$($txtInst.Text)" }
    if ($txtPort.Text){ $conn = "$conn,$($txtPort.Text)" }
    return $conn
}

function Get-Target {
    $conn = $txtTServer.Text.Trim()
    if ($txtTInst.Text){ $conn = "$conn\$($txtTInst.Text)" }
    if ($txtTPort.Text){ $conn = "$conn,$($txtTPort.Text)" }
    return $conn
}

function Import-Config-CSV($filePath){

    try{
        if(!(Test-Path $filePath)){
            throw "Config file not found: $filePath"
        }

        $data = Import-Csv $filePath | Select-Object -First 1

        # ✅ Source
        $txtServer.Text = $data.SourceServer
        $txtInst.Text   = $data.SourceInstance
        $txtPort.Text   = $data.SourcePort

        # ✅ Target
        $txtTServer.Text = $data.TargetServer
        $txtTInst.Text   = $data.TargetInstance
        $txtTPort.Text   = $data.TargetPort

        # ✅ Backup / Copy / Restore
        $txtPath.Text   = $data.BackupPath
        $txtSrc.Text    = $data.CopySource
        $txtDest.Text   = $data.CopyDest
        $txtFolder.Text = $data.RestoreFolder

        $chkReplace.Checked = [System.Convert]::ToBoolean($data.Replace)

        # ✅ Optional: Pre-select DBs (after load)
        $global:ImportedDBList = $data.Databases -split ';'

        $lblExpo.Text = "$(Get-Date): ✅ Config Loaded from CSV"
        $lblExpo.ForeColor = "Green"

        # ✅ Refresh UI
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()

        # ✅ Auto trigger connection + load DB
        Start-Sleep -Milliseconds 100
        $btnTest.PerformClick()

        Start-Sleep -Milliseconds 100
        $btnLoad.PerformClick()

    }catch{
        $lblExpo.Text = "$(Get-Date): ❌ Failed to load config - $($_.Exception.Message)"
        $lblExpo.ForeColor = "Red"
    }
}


# =========================================================
# ✅ Imort Config TAB
# =========================================================
# ================= Imort Config TAB =================

$txtImortPath = New-Object System.Windows.Forms.TextBox -Property @{
    Location="200,20"; Width=500
}

$tabImort = New-Object System.Windows.Forms.TabPage -Property @{Text="Imort Config"}

$btnImort = New-Object System.Windows.Forms.Button -Property @{
    Text="Import CSV"
    Location="20,100"
}

$btnImort.Add_Click({
#$btnImportCSV.Add_Click({

    try{

        Import-Config-CSV $filePath

    }catch{
        $lblExpo.Text = "$(Get-Date): ❌ $($_.Exception.Message)"
        $lblExpo.ForeColor = "Red"
    }
})

# ================= SQL EXEC =================
function Run-SQL($srv,$qry,$isTarget=$false){
    Invoke-Sqlcmd -ServerInstance $srv -Query $qry -ErrorAction Stop
}

# =========================================================
# ✅ SOURCE CONNECTION TAB
# =========================================================
$txtSServer = New-Object System.Windows.Forms.TextBox -Property @{Location="200,20";Width=250}
$txtInst   = New-Object System.Windows.Forms.TextBox -Property @{Location="200,60";Width=250}
$txtPort   = New-Object System.Windows.Forms.TextBox -Property @{Location="200,100";Width=120}

$btnTest = New-Object System.Windows.Forms.Button -Property @{Text="Test Connection";Location="200,150"}

$btnTest.Add_Click({
    try{
        Run-SQL (Get-Source) "SELECT @@SERVERNAME"
        Set-ConnStatus "✅ Connected: $(Get-Source)" "Green"
        $btnLoad.Enabled = $true
        $btnBackup.Enabled = $true
    }catch{
        Set-ConnStatus "❌ $($_.Exception.Message)" "Red"
    }
})

# =========================================================
# ✅ TARGET CONNECTION TAB
# =========================================================
$txtTServer = New-Object System.Windows.Forms.TextBox -Property @{Location="200,20";Width=250}
$txtTInst   = New-Object System.Windows.Forms.TextBox -Property @{Location="200,60";Width=250}
$txtTPort   = New-Object System.Windows.Forms.TextBox -Property @{Location="200,100";Width=120}

$btnTTest = New-Object System.Windows.Forms.Button -Property @{Text="Test Target";Location="200,150"}

$btnTTest.Add_Click({
    try{
        Run-SQL (Get-Target) "SELECT @@SERVERNAME"
        Set-TgtStatus "✅ Connected: $(Get-Target)" "Green"
    }catch{
        Set-TgtStatus "❌ $($_.Exception.Message)" "Red"
    }
})

# =========================================================
# ✅ DATABASE TAB
# =========================================================
$listDB = New-Object System.Windows.Forms.ListBox -Property @{Size="800,350";Location="20,20";SelectionMode="MultiExtended"}

$btnLoad = New-Object System.Windows.Forms.Button -Property @{Text="Load DBs";Location="20,400";Enabled=$false}

$btnLoad.Add_Click({
    try{
        $listDB.Items.Clear()
        $dbs = Invoke-Sqlcmd -ServerInstance (Get-Source) -Query "SELECT name FROM sys.databases"
        foreach($d in $dbs){
            $listDB.Items.Add($d.name)

            if($global:ImportedDBList -contains $d.name){
                $listDB.SetSelected($listDB.Items.Count - 1, $true)
            }
        }
        Set-ConnStatus "✅ DB Loaded" "Green"
    }catch{
        Set-ConnStatus "❌ Load failed" "Red"
    }
})

# =========================================================
# ✅ BACKUP TAB
# =========================================================
$txtPath = New-Object System.Windows.Forms.TextBox -Property @{Location="150,20";Width=400}
$btnBackup = New-Object System.Windows.Forms.Button -Property @{Text="Backup Selected DBs";Location="20,60";Enabled=$false}
$progress = New-Object System.Windows.Forms.ProgressBar -Property @{Location="20,120";Width=800}

$btnBackup.Add_Click({
    try{
        $i=0;$sel=$listDB.SelectedItems
        foreach($db in $sel){
            $i++
            Run-SQL (Get-Source) "BACKUP DATABASE [$db] TO DISK='$($txtPath.Text)\$db.bak' WITH INIT"
            $progress.Value=[math]::Min(100,($i/$sel.Count)*100)
            Set-BkpStatus "✅ $db completed" "Green"
        }
    }catch{
        Set-BkpStatus "❌ $($_.Exception.Message)" "Red"
    }
})

# =========================================================
# ✅ COPY TAB
# =========================================================
$txtSrc = New-Object System.Windows.Forms.TextBox -Property @{Location="200,20";Width=500}
$txtDest = New-Object System.Windows.Forms.TextBox -Property @{Location="200,60";Width=500}

$btnCopy = New-Object System.Windows.Forms.Button -Property @{Text="Copy Files";Location="20,100"}
$copyProgress = New-Object System.Windows.Forms.ProgressBar -Property @{Location="20,150";Width=800}

$btnCopy.Add_Click({
    try{
        $files = Get-ChildItem $txtSrc.Text -Filter *.bak
        $i=0
        foreach($f in $files){
            $i++
            Copy-Item $f.FullName (Join-Path $txtDest.Text $f.Name) -Force
            $copyProgress.Value=[math]::Min(100,($i/$files.Count)*100)
            Set-CopyStatus "✅ Copied $($f.Name)" "Green"
        }
    }catch{
        Set-CopyStatus "❌ $($_.Exception.Message)" "Red"
    }
})

# =========================================================
# ✅ RESTORE TAB
# =========================================================
# ================= RESTORE TAB =================

$txtFolder = New-Object System.Windows.Forms.TextBox -Property @{Location="200,20";Width=500}

$chkReplace = New-Object System.Windows.Forms.CheckBox -Property @{
    Text="Overwrite Existing Database (WITH REPLACE)"
    Location="20,80"
    Checked=$true
}

$btnRestore = New-Object System.Windows.Forms.Button -Property @{
    Text="Restore All Databases"
    Location="20,120"
}

$restoreProgress = New-Object System.Windows.Forms.ProgressBar -Property @{
    Location="20,180"
    Width=800
    Minimum=0
    Maximum=100
}

$lblRst = New-Object System.Windows.Forms.Label -Property @{
    Location="20,220"
    Width=800
}

function Set-RstStatus($msg,$color="Blue"){
    $lblRst.Text = "$(Get-Date): $msg"
    $lblRst.ForeColor = $color
    $form.Refresh()
}

# ✅ RESTORE LOGIC
$btnRestore.Add_Click({

    try {
        $srv = Get-Target
        $folder = $txtFolder.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($folder)) {
            throw "Backup folder required"
        }

        if (!(Test-Path $folder)) {
            throw "Backup folder not found"
        }

        $files = Get-ChildItem $folder -Filter *.bak

        if ($files.Count -eq 0) {
            throw "No backup files found"
        }

        # ✅ Get default SQL Server paths
        $paths = Invoke-Sqlcmd -ServerInstance $srv -Query "
        SELECT 
            CAST(SERVERPROPERTY('InstanceDefaultDataPath') AS NVARCHAR(400)) AS DataPath,
            CAST(SERVERPROPERTY('InstanceDefaultLogPath')  AS NVARCHAR(400)) AS LogPath
        "

        $dataRoot = $paths.DataPath
        $logRoot  = $paths.LogPath

        $i = 0

        foreach ($file in $files) {

            $db = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $i++

            try {

                Set-RstStatus "Restoring $db ($i/$($files.Count))..."

                # ✅ Get logical file details
                $fileList = Invoke-Sqlcmd -ServerInstance $srv -Query "
                RESTORE FILELISTONLY FROM DISK = '$($file.FullName)'
                "

                $dataLogical = ($fileList | Where-Object { $_.Type -eq "D" }).LogicalName
                $logLogical  = ($fileList | Where-Object { $_.Type -eq "L" }).LogicalName

                # ✅ Map to default instance paths
                $dataFile = "$dataRoot$db.mdf"
                $logFile  = "$logRoot${db}_log.ldf"

                # ✅ Replace option
                $replaceClause = ""
                if ($chkReplace.Checked) {
                    $replaceClause = ", REPLACE"
                }

                # ✅ Force disconnect existing DB
                Invoke-Sqlcmd -ServerInstance $srv -Query "
                IF DB_ID('$db') IS NOT NULL
                ALTER DATABASE [$db] SET SINGLE_USER WITH ROLLBACK IMMEDIATE
                "

                # ✅ Restore with MOVE
                $restoreQuery = @"
RESTORE DATABASE [$db]
FROM DISK = '$($file.FullName)'
WITH 
    MOVE '$dataLogical' TO '$dataFile',
    MOVE '$logLogical'  TO '$logFile'
    $replaceClause,
    RECOVERY
"@

                Invoke-Sqlcmd -ServerInstance $srv -Query $restoreQuery -ErrorAction Stop

                # ✅ Back to multi-user
                Invoke-Sqlcmd -ServerInstance $srv -Query "
                ALTER DATABASE [$db] SET MULTI_USER
                "

                # ✅ Progress update
                $restoreProgress.Value = [math]::Min(100, ($i / $files.Count) * 100)
                $form.Refresh()

                Set-RstStatus "✅ Restored: $db" "Green"
            }
            catch {
                Set-RstStatus "❌ Failed: $db - $($_.Exception.Message)" "Red"
            }
        }

        $restoreProgress.Value = 100
        Set-RstStatus "✅ All restores completed successfully" "Green"
    }
    catch {
        Set-RstStatus "❌ $($_.Exception.Message)" "Red"
    }
})


# ✅ ADD CONTROLS
$tabRestore.Controls.AddRange(@(
    (New-Object System.Windows.Forms.Label -Property @{Text="Backup Folder";Location="20,20"}),
    $txtFolder,
    $chkReplace,
    $btnRestore,
    $restoreProgress,
    $lblRst
))

# ================= ADD CONTROLS =================
$tabConn.Controls.AddRange(@(
(New-Object System.Windows.Forms.Label -Property @{Text="Server";Location="20,20"}),$txtSServer,
(New-Object System.Windows.Forms.Label -Property @{Text="Instance";Location="20,60"}),$txtInst,
(New-Object System.Windows.Forms.Label -Property @{Text="Port";Location="20,100"}),$txtPort,
$btnTest,$lblConn))

$tabTarget.Controls.AddRange(@(
(New-Object System.Windows.Forms.Label -Property @{Text="Server";Location="20,20"}),$txtTServer,
(New-Object System.Windows.Forms.Label -Property @{Text="Instance";Location="20,60"}),$txtTInst,
(New-Object System.Windows.Forms.Label -Property @{Text="Port";Location="20,100"}),$txtTPort,
$btnTTest,$lblTgt))

$tabDB.Controls.AddRange(@($listDB,$btnLoad))
$tabBackup.Controls.AddRange(@((New-Object System.Windows.Forms.Label -Property @{Text="Backup Path";Location="20,20"}),$txtPath,$btnBackup,$progress,$lblBkp))
$tabCopy.Controls.AddRange(@((New-Object System.Windows.Forms.Label -Property @{Text="Source Folder";Location="20,20"}),$txtSrc,(New-Object System.Windows.Forms.Label -Property @{Text="Destination Folder";Location="20,60"}),$txtDest,$btnCopy,$copyProgress,$lblCopy))
$tabRestore.Controls.AddRange(@((New-Object System.Windows.Forms.Label -Property @{Text="Backup Folder";Location="20,20"}),$txtFolder,$btnRestore,$restoreProgress,$lblRst))
$tabImort.Controls.AddRange(@(
    (New-Object System.Windows.Forms.Label -Property @{Text="Output Folder";Location="20,20"}),
    $txtImortPath,
    $btnImort,
    $lblImort
))

$tabs.Controls.AddRange(@($tabConn,$tabTarget,$tabDB,$tabBackup,$tabCopy,$tabRestore,$tabImort))
$form.Controls.Add($tabs)

# ================= RUN =================

$form.Add_Shown({

    # 👉 Change path if needed
    $configFile = "C:\DBMigGUI\ExportConfig\SQL_GUI_Config.csv"

    Import-Config-CSV $configFile
})

[void]$form.ShowDialog()