Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# --- FONCTIONS DE VERIFICATION / INSTALLATION ---

function Check-Python {
    try {
        python --version 2>$null
        if ($LASTEXITCODE -eq 0) { return $true }
        else { return $false }
    } catch { return $false }
}

function Check-PythonModule {
    param([string]$moduleName)
    python -c "import $moduleName" 2>$null
    return ($LASTEXITCODE -eq 0)
}

function Install-Python {
    $pythonInstaller = "$env:TEMP\python-installer.exe"
    $pythonUrl = "https://www.python.org/ftp/python/3.12.2/python-3.12.2-amd64.exe"
    Invoke-WebRequest -Uri $pythonUrl -OutFile $pythonInstaller
    Start-Process -FilePath $pythonInstaller -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait
    Remove-Item $pythonInstaller
}

function Install-PythonModule {
    param([string]$moduleName)
    python -m pip install --upgrade pip
    python -m pip install $moduleName
}

function Check-Ollama {
    $ollamaPath = (Get-Command ollama.exe -ErrorAction SilentlyContinue).Source
    if ($ollamaPath) { return $ollamaPath }
    else {
        $defaultPath = "$env:LOCALAPPDATA\Programs\Ollama\ollama.exe"
        if (Test-Path $defaultPath) { return $defaultPath }
        return $null
    }
}

# --- VERIFICATION PYTHON ET MODULES ---

$pythonInstalled = Check-Python
if (-not $pythonInstalled) {
    $result = [System.Windows.Forms.MessageBox]::Show("Python n'est pas installé. Installer maintenant ?","Alerte","YesNo","Warning")
    if ($result -eq "Yes") {
        Install-Python
        [System.Windows.Forms.MessageBox]::Show("Python installé. Relancez le script.","Info","OK","Information") | Out-Null
        exit
    } else {
        [System.Windows.Forms.MessageBox]::Show("Python est requis pour continuer.","Erreur","OK","Error") | Out-Null
        exit
    }
}

$tkinterInstalled = Check-PythonModule -moduleName "tkinter"
$requestsInstalled = Check-PythonModule -moduleName "requests"

$formCheck = New-Object System.Windows.Forms.Form
$formCheck.Text = "Vérification Python et dépendances"
$formCheck.Size = New-Object System.Drawing.Size(400, 200)
$formCheck.StartPosition = "CenterScreen"

$labelPython = New-Object System.Windows.Forms.Label
$labelPython.Text = "Python : Installé"
$labelPython.Location = New-Object System.Drawing.Point(10, 20)
$labelPython.Size = New-Object System.Drawing.Size(380, 20)
$formCheck.Controls.Add($labelPython)

$tkText = if ($tkinterInstalled) { "Installé" } else { "Non installé" }
$reqText = if ($requestsInstalled) { "Installé" } else { "Non installé" }

$labelTk = New-Object System.Windows.Forms.Label
$labelTk.Text = "tkinter : " + $tkText
$labelTk.Location = New-Object System.Drawing.Point(10, 50)
$labelTk.Size = New-Object System.Drawing.Size(380, 20)
$formCheck.Controls.Add($labelTk)

$labelReq = New-Object System.Windows.Forms.Label
$labelReq.Text = "requests : " + $reqText
$labelReq.Location = New-Object System.Drawing.Point(10, 80)
$labelReq.Size = New-Object System.Drawing.Size(380, 20)
$formCheck.Controls.Add($labelReq)

$btnInstallDeps = New-Object System.Windows.Forms.Button
$btnInstallDeps.Text = "Installer les dépendances manquantes"
$btnInstallDeps.Location = New-Object System.Drawing.Point(80, 120)
$btnInstallDeps.Size = New-Object System.Drawing.Size(230, 30)
$btnInstallDeps.Add_Click({
    if (-not $tkinterInstalled) { Install-PythonModule -moduleName "tkinter" }
    if (-not $requestsInstalled) { Install-PythonModule -moduleName "requests" }
    [System.Windows.Forms.MessageBox]::Show("Dépendances installées.","Info","OK","Information") | Out-Null
    $formCheck.Close()
})
$formCheck.Controls.Add($btnInstallDeps)

$btnContinue = New-Object System.Windows.Forms.Button
$btnContinue.Text = "Continuer"
$btnContinue.Location = New-Object System.Drawing.Point(150, 160)
$btnContinue.Size = New-Object System.Drawing.Size(100, 30)
$btnContinue.Add_Click({ $formCheck.Close() })
$formCheck.Controls.Add($btnContinue)

$formCheck.Topmost = $true
$formCheck.Add_Shown({$formCheck.Activate()})
[void]$formCheck.ShowDialog()

# --- VERIFICATION OLLAMA ET MODELE ---

$formOllama = New-Object System.Windows.Forms.Form
$formOllama.Text = "Vérification Ollama"
$formOllama.Size = New-Object System.Drawing.Size(400, 200)
$formOllama.StartPosition = "CenterScreen"

$ollamaExe = Check-Ollama
$lblOllama = New-Object System.Windows.Forms.Label
if ($ollamaExe) {
    $lblOllama.Text = "Ollama : Installé"
    setx OLLAMA_PATH $ollamaExe | Out-Null
} else {
    $lblOllama.Text = "Ollama : Non installé"
}
$lblOllama.Location = New-Object System.Drawing.Point(10, 20)
$lblOllama.Size = New-Object System.Drawing.Size(380, 20)
$formOllama.Controls.Add($lblOllama)

$ollamaModel = "gpt-oss:20b"
$modelInstalled = $false
if ($ollamaExe) {
    $proc = Start-Process -FilePath $ollamaExe -ArgumentList "list" -NoNewWindow -PassThru -RedirectStandardOutput "$env:TEMP\ollama_models.txt" -Wait
    $installedModels = Get-Content "$env:TEMP\ollama_models.txt"
    if (($installedModels -like "*gpt-oss:20b*")) { $modelInstalled = $true }
}

$lblModel = New-Object System.Windows.Forms.Label
if ($modelInstalled) {
    $lblModel.Text = "Modèle $ollamaModel : Installé"
} else {
    $lblModel.Text = "Modèle $ollamaModel : Non installé"
}
$lblModel.Location = New-Object System.Drawing.Point(10, 50)
$lblModel.Size = New-Object System.Drawing.Size(380, 20)
$formOllama.Controls.Add($lblModel)

$btnContinueOllama = New-Object System.Windows.Forms.Button
$btnContinueOllama.Text = "Continuer"
$btnContinueOllama.Location = New-Object System.Drawing.Point(150, 120)
$btnContinueOllama.Size = New-Object System.Drawing.Size(100, 30)
$btnContinueOllama.Add_Click({ $formOllama.Close() })
$formOllama.Controls.Add($btnContinueOllama)

$formOllama.Topmost = $true
$formOllama.Add_Shown({$formOllama.Activate()})
[void]$formOllama.ShowDialog()

if (-not $ollamaExe) {
    [System.Windows.Forms.MessageBox]::Show("Ollama n'est pas installé. Vous allez être redirigé vers la page d'installation.","Alerte","OK","Warning") | Out-Null
    Start-Process "https://ollama.com/download"
    exit
}
if (-not $modelInstalled) {
    [System.Windows.Forms.MessageBox]::Show("Le modèle $ollamaModel n'est pas installé.`nVeuillez lancer Ollama et entrer :`nollama run $ollamaModel","Info","OK","Information") | Out-Null
    exit
}

# --- INTERFACE CONFIGURATION NOTION ---

$formNotion = New-Object System.Windows.Forms.Form
$formNotion.Text = "Configuration Notion"
$formNotion.Size = New-Object System.Drawing.Size(500,300)
$formNotion.StartPosition = "CenterScreen"

$lblToken = New-Object System.Windows.Forms.Label
$lblToken.Text = "Entrez votre Notion Token :"
$lblToken.Location = New-Object System.Drawing.Point(10,20)
$formNotion.Controls.Add($lblToken)

$txtToken = New-Object System.Windows.Forms.TextBox
$txtToken.Location = New-Object System.Drawing.Point(10,50)
$txtToken.Size = New-Object System.Drawing.Size(450,20)
$formNotion.Controls.Add($txtToken)

$lblDb = New-Object System.Windows.Forms.Label
$lblDb.Text = "Entrez l'ID de la base Notion :"
$lblDb.Location = New-Object System.Drawing.Point(10,90)
$formNotion.Controls.Add($lblDb)

$txtDb = New-Object System.Windows.Forms.TextBox
$txtDb.Location = New-Object System.Drawing.Point(10,120)
$txtDb.Size = New-Object System.Drawing.Size(450,20)
$formNotion.Controls.Add($txtDb)

$btnHelp = New-Object System.Windows.Forms.Button
$btnHelp.Text = "Comment faire ?"
$btnHelp.Location = New-Object System.Drawing.Point(10,160)
$btnHelp.Size = New-Object System.Drawing.Size(120,30)
$btnHelp.Add_Click({ Start-Process "https://developers.notion.com/docs/getting-started" })
$formNotion.Controls.Add($btnHelp)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Enregistrer"
$btnSave.Location = New-Object System.Drawing.Point(200,200)
$btnSave.Size = New-Object System.Drawing.Size(100,30)
$btnSave.Add_Click({
    $token = $txtToken.Text
    $dbid = $txtDb.Text
    if (-not $token -or -not $dbid) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez remplir les deux champs.","Erreur","OK","Error") | Out-Null
        return
    }

    # Vérification API Notion
    try {
        $headers = @{
            "Authorization" = "Bearer $token"
            "Notion-Version" = "2022-06-28"
        }
        $url = "https://api.notion.com/v1/databases/$dbid"
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method GET -ErrorAction Stop

        if ($response.id -ne $null) {
            $scriptPath = Join-Path (Split-Path -Parent $PSCommandPath) "AutoResumeNotion.py"

            if (-not (Test-Path $scriptPath)) {
                [System.Windows.Forms.MessageBox]::Show("AutoResumeNotion.py introuvable.","Erreur","OK","Error") | Out-Null
                return
            }

            $content = Get-Content $scriptPath
            $content = $content -replace 'NOTION_TOKEN = .*', "NOTION_TOKEN = `"$token`""
            $content = $content -replace 'DATABASE_ID = .*', "DATABASE_ID = `"$dbid`""
            Set-Content -Path $scriptPath -Value $content -Encoding UTF8

            [System.Windows.Forms.MessageBox]::Show("Configuration Notion vérifiée et enregistrée avec succès.","Succès","OK","Information") | Out-Null
            $formNotion.Close()
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la vérification Notion.`nVérifiez le Token et l'ID de la base.","Erreur","OK","Error") | Out-Null
        return
    }
})
$formNotion.Controls.Add($btnSave)

$formNotion.Topmost = $true
$formNotion.Add_Shown({$formNotion.Activate()})
[void]$formNotion.ShowDialog()

# --- CREATION DU .BAT ---

$pythonPath = (Get-Command python).Source
$scriptPath = Join-Path (Split-Path -Parent $PSCommandPath) "AutoResumeNotion.py"

if ($ollamaExe) {
    $content = @"
@echo off
REM Boucle jusqu'à ce que Ollama soit lancé
:waitOllama
tasklist /FI "IMAGENAME eq ollama.exe" | find /I "ollama.exe" >nul
if errorlevel 1 (
    REM Ollama n'est pas lancé, on le lance
    start "" "$ollamaExe"
    timeout /t 2 /nobreak >nul
    goto waitOllama
)

REM Ollama est lancé, on continue
"$pythonPath" "$scriptPath"
pause
exit
"@
} else {
    $content = @"
@echo off
"$pythonPath" "$scriptPath"
pause
exit
"@
}

Set-Content -Path "ResumeNotion.bat" -Value $content -Encoding ASCII
[System.Windows.Forms.MessageBox]::Show("ResumeNotion.bat créé (avec gestion de lancement Ollama).","Succès","OK","Information") | Out-Null

$regScript = Join-Path (Split-Path -Parent $scriptPath) "Registry_installer.py"
if (Test-Path $regScript) {
    Start-Process $pythonPath -ArgumentList "`"$regScript`"" -Verb RunAs
} else {
    [System.Windows.Forms.MessageBox]::Show("Registry_installer.py introuvable.","Erreur","OK","Error") | Out-Null
}
