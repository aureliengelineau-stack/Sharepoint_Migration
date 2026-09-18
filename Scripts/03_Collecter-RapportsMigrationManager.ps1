<#
.SYNOPSIS
    Rassemble les rapports Migration Manager (migrations et scans) dans le dossier lu par Power BI.

.DESCRIPTION
    Trois sources sont combinées, dans cet ordre :
      1. Module PowerShell Migration Manager (si disponible) : Connect-MigrationService, puis
         Get-MigrationReport et Get-ScanReport. Module à télécharger sur https://aka.ms/MMPowerShellModule,
         à extraire par exemple dans Kit\Outils\MigrationManager (ou indiquer -ModuleMigrationManager).
      2. Rapports téléchargés à la main dans le centre de migration (Download task report, Summary report,
         Download summary report / scan log, ReportAggregator) : fichiers .csv ou .zip du dossier -DossierImport
         (par défaut : Téléchargements, fichiers des -DepuisJours derniers jours).
      3. Dossier de travail de l'agent sur ce poste (%appdata%\Microsoft\SPMigration), s'il existe.

    Chaque CSV (y compris dans les .zip) est reconnu d'après ses colonnes, dédoublonné (empreinte SHA-256)
    et copié sous un nom normalisé :
      <Sortie>\MigrationManager\Synthese\SummaryReport_<empreinte>.csv
      <Sortie>\MigrationManager\Taches\<TaskID>\ItemReport_<empreinte>.csv, ItemFailureReport_..., ItemSummary_...,
                                                 ScanSummary_..., StructureReport_...
      <Sortie>\MigrationManager\Scans\ScanReport_<empreinte>.csv, ScanLog_...
    et un index <Sortie>\Index_Rapports_MigrationManager.csv est produit.

    Au lancement, une fenêtre Windows demande le dossier de sortie (proposé : -Sortie). -SansFenetre utilise
    -Sortie directement. Rapports conservés 90 jours par Microsoft : collectez-les avant expiration.

.EXAMPLE
    .\03_Collecter-RapportsMigrationManager.ps1
.EXAMPLE
    .\03_Collecter-RapportsMigrationManager.ps1 -ModuleMigrationManager 'C:\Outils\MMPowerShell' -DepuisDate '2026-09-01' -Zip
.EXAMPLE
    .\03_Collecter-RapportsMigrationManager.ps1 -SansModule -DossierImport 'C:\DemoMigration\Telechargements_MM'
#>
[CmdletBinding()]
param(
    [string]$FichierParametres = $(if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Parametres_Client.psd1' } else { 'Parametres_Client.psd1' }),
    [hashtable]$ParametresClient = $(if (Test-Path -LiteralPath $FichierParametres) { Import-PowerShellDataFile -LiteralPath $FichierParametres } else { @{} }),
    [string]$Racine = $(if ($ParametresClient['CheminLocal']) { Split-Path -Parent $ParametresClient['CheminLocal'] } else { 'C:\DemoMigration' }),
    [string]$Sortie = $(if ($ParametresClient['DossierRapports']) { $ParametresClient['DossierRapports'] } else { (Join-Path $Racine 'Rapports') }),
    [string]$ModuleMigrationManager,
    [switch]$SansModule,
    [datetime]$DepuisDate,
    [string]$NomTacheContient,
    [string]$Tags,
    [string]$DossierImport,
    [int]$DepuisJours = 14,
    [string]$DossierAgent,
    [switch]$AvecJournaux,
    [switch]$Zip,
    [switch]$SansFenetre
)

$ErrorActionPreference = 'Stop'

# Paramètres de la mission (Scripts\Parametres_Client.psd1, écrit par 00_Configurer-Parametres.ps1).
# Les valeurs du fichier servent de valeurs par défaut ; un paramètre en ligne de commande reste prioritaire.
if ($ParametresClient.Count -gt 0) {
    $nomMission = if ($ParametresClient['Client']) { $ParametresClient['Client'] } else { '(client non renseigné)' }
    Write-Host ("Mission : {0}   [{1}]" -f $nomMission, (Split-Path -Leaf $FichierParametres)) -ForegroundColor DarkCyan
}
$debut = Get-Date

# ---------------------------------------------------------------- Choix du dossier de sortie
function Select-DossierSortie {
    <#  Ouvre la fenêtre Windows « Rechercher un dossier », pré-positionnée sur $DossierPropose.
        Retourne le chemin choisi, ou $null si l'utilisateur annule.
        Sans interface graphique (hors Windows, session non interactive, échec de la fenêtre) : saisie console,
        puis dossier proposé si aucune console n'est disponible.  #>
    param([string]$Titre, [string]$DossierPropose, [switch]$SansFenetre)

    if ($SansFenetre) { return $DossierPropose }

    $creeTemporairement = $false
    $choix = $null
    $fenetreOk = $false
    if ([System.Environment]::OSVersion.Platform -eq 'Win32NT' -and [System.Environment]::UserInteractive) {
        # Le dossier proposé est créé s'il n'existe pas, pour pouvoir le présélectionner (retiré s'il reste vide et non choisi)
        if ($DossierPropose -and -not (Test-Path -LiteralPath $DossierPropose)) {
            try { New-Item -ItemType Directory -Path $DossierPropose -Force | Out-Null; $creeTemporairement = $true } catch { }
        }
        $bloc = {
            param($Titre, $Initial)
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.Application]::EnableVisualStyles()
            $proprietaire = New-Object System.Windows.Forms.Form -Property @{ TopMost = $true; ShowInTaskbar = $false }
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.Description = $Titre
            $dlg.ShowNewFolderButton = $true
            if ($dlg.PSObject.Properties.Name -contains 'UseDescriptionForTitle') { $dlg.UseDescriptionForTitle = $true }
            if ($Initial -and (Test-Path -LiteralPath $Initial)) {
                if ($dlg.PSObject.Properties.Name -contains 'InitialDirectory') { $dlg.InitialDirectory = $Initial }
                $dlg.SelectedPath = $Initial
            }
            try {
                if ($dlg.ShowDialog($proprietaire) -eq [System.Windows.Forms.DialogResult]::OK) { $dlg.SelectedPath } else { '' }
            } finally { $dlg.Dispose(); $proprietaire.Dispose() }
        }
        try {
            if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA) {
                $choix = [string](& $bloc $Titre $DossierPropose | Select-Object -First 1)
            } else {
                # Les fenêtres Windows Forms exigent un thread STA (ex. certaines sessions PowerShell 7)
                $rs = [runspacefactory]::CreateRunspace()
                $rs.ApartmentState = [System.Threading.ApartmentState]::STA
                $rs.Open()
                $ps = [powershell]::Create()
                $ps.Runspace = $rs
                [void]$ps.AddScript($bloc.ToString()).AddArgument($Titre).AddArgument($DossierPropose)
                try { $choix = [string](@($ps.Invoke()) | Select-Object -First 1) } finally { $ps.Dispose(); $rs.Dispose() }
            }
            $fenetreOk = $true
        } catch {
            Write-Warning "Fenêtre de sélection indisponible : $($_.Exception.Message)"
        }
    }

    if (-not $fenetreOk) {
        try {
            $saisie = Read-Host "$Titre`n  Dossier de sortie [Entrée = $DossierPropose]"
            $choix = if ([string]::IsNullOrWhiteSpace($saisie)) { $DossierPropose } else { $saisie.Trim().Trim('"') }
        } catch {
            $choix = $DossierPropose
        }
    }

    if ($creeTemporairement -and $choix -ne $DossierPropose -and (Test-Path -LiteralPath $DossierPropose) -and
        -not (Get-ChildItem -LiteralPath $DossierPropose -Force -ErrorAction SilentlyContinue)) {
        Remove-Item -LiteralPath $DossierPropose -Force -ErrorAction SilentlyContinue
    }
    if ([string]::IsNullOrWhiteSpace($choix)) { return $null }
    $choix = $choix.TrimEnd('\', '/')
    if ($choix -match '^[A-Za-z]:$') { $choix += '\' }   # racine de lecteur : C:\ et non C:
    return $choix
}

# ---------------------------------------------------------------- Préparation
$Sortie = Select-DossierSortie -Titre 'Choisissez le dossier de sortie des rapports Migration Manager (sous-dossier MigrationManager + index)' -DossierPropose $Sortie -SansFenetre:$SansFenetre
if (-not $Sortie) { Write-Warning 'Sélection annulée : aucun rapport collecté.'; return }
$cibleMM = Join-Path $Sortie 'MigrationManager'
New-Item -ItemType Directory -Path $cibleMM -Force | Out-Null

if (-not $DossierImport) {
    $profil = if ($env:USERPROFILE) { $env:USERPROFILE } else { $HOME }
    $DossierImport = Join-Path $profil 'Downloads'
    $importRecursif = $false
} else { $importRecursif = $true }
if (-not $DossierAgent -and $env:APPDATA) { $DossierAgent = Join-Path $env:APPDATA 'Microsoft\SPMigration' }

$travail = Join-Path ([System.IO.Path]::GetTempPath()) ('MM_Collecte_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $travail -Force | Out-Null
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$sha = [System.Security.Cryptography.SHA256]::Create()
$reGuid = New-Object System.Text.RegularExpressions.Regex '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

# ---------------------------------------------------------------- 1. Module PowerShell Migration Manager
function Find-ModuleMM {
    if ($ModuleMigrationManager) {
        if (Test-Path -LiteralPath $ModuleMigrationManager -PathType Leaf) { return (Resolve-Path -LiteralPath $ModuleMigrationManager).ProviderPath }
        if (Test-Path -LiteralPath $ModuleMigrationManager -PathType Container) {
            $dll = Get-ChildItem -LiteralPath $ModuleMigrationManager -Recurse -Filter 'Microsoft.SharePoint.MigrationManager.PowerShell.dll' -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($dll) { return $dll.FullName }
        }
        Write-Warning "Module Migration Manager introuvable : $ModuleMigrationManager"
        return $null
    }
    $candidats = @((Join-Path $PSScriptRoot '..\Outils'), $PSScriptRoot) | Where-Object { Test-Path -LiteralPath $_ }
    foreach ($c in $candidats) {
        $dll = Get-ChildItem -LiteralPath $c -Recurse -Filter 'Microsoft.SharePoint.MigrationManager.PowerShell.dll' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($dll) { return $dll.FullName }
    }
    return $null
}

$sources = New-Object System.Collections.Generic.List[object]   # dossiers à analyser : @{ Chemin; Origine; Recursif; Filtrer }
if (-not $SansModule) {
    $dll = Find-ModuleMM
    if ($dll) {
        try {
            Get-ChildItem -LiteralPath (Split-Path $dll -Parent) -Recurse -File | Unblock-File -ErrorAction SilentlyContinue
            Import-Module $dll -ErrorAction Stop
            Write-Host 'Connexion au service Migration Manager (compte administrateur SharePoint)...' -ForegroundColor Cyan
            Connect-MigrationService | Out-Null
            $filtre = @{}
            if ($PSBoundParameters.ContainsKey('DepuisDate')) { $filtre.StartTime = $DepuisDate }
            if ($Tags) { $filtre.Tags = $Tags }

            $dMig = Join-Path $travail 'module_migrations'; New-Item -ItemType Directory -Path $dMig -Force | Out-Null
            $pMig = $filtre.Clone(); if ($NomTacheContient) { $pMig.TaskNameContains = $NomTacheContient }
            try { Get-MigrationReport -OutputPath $dMig @pMig | Out-Null; $sources.Add(@{ Chemin = $dMig; Origine = 'Module (Get-MigrationReport)'; Recursif = $true; Filtrer = $false }) }
            catch { Write-Warning "Get-MigrationReport : $($_.Exception.Message)" }

            $dScan = Join-Path $travail 'module_scans'; New-Item -ItemType Directory -Path $dScan -Force | Out-Null
            try { Get-ScanReport -OutputPath $dScan @filtre | Out-Null; $sources.Add(@{ Chemin = $dScan; Origine = 'Module (Get-ScanReport)'; Recursif = $true; Filtrer = $false }) }
            catch { Write-Warning "Get-ScanReport : $($_.Exception.Message)" }
        } catch {
            Write-Warning ("Module Migration Manager non utilisable : {0}`n  Astuce : le module cible Windows PowerShell 5.1 ; relancez depuis « Windows PowerShell » si besoin." -f $_.Exception.Message)
        }
    } else {
        Write-Host 'Module PowerShell Migration Manager non trouvé (https://aka.ms/MMPowerShellModule) : collecte depuis les téléchargements et l''agent.' -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------- 2 et 3. Téléchargements et agent
if (Test-Path -LiteralPath $DossierImport) { $sources.Add(@{ Chemin = $DossierImport; Origine = 'Téléchargement'; Recursif = $importRecursif; Filtrer = -not $importRecursif }) }
else { Write-Warning "Dossier d'import introuvable : $DossierImport" }
if ($DossierAgent -and (Test-Path -LiteralPath $DossierAgent)) { $sources.Add(@{ Chemin = $DossierAgent; Origine = 'Agent (dossier de travail)'; Recursif = $true; Filtrer = $false }) }

# ---------------------------------------------------------------- Reconnaissance des rapports
function Get-Colonnes([string]$Chemin) {
    $sr = New-Object System.IO.StreamReader($Chemin, [System.Text.Encoding]::UTF8, $true)
    try { $ligne = $sr.ReadLine() } finally { $sr.Dispose() }
    if (-not $ligne) { return @() }
    $delim = if (($ligne.Split(';').Count) -gt ($ligne.Split(',').Count)) { ';' } else { ',' }
    return @([regex]::Split($ligne, [regex]::Escape($delim) + '(?=(?:[^"]*"[^"]*")*[^"]*$)') |
        ForEach-Object { ($_.Trim().Trim('"') -replace '[^A-Za-z0-9]', '').ToLowerInvariant() })
}

function Get-TypeRapport([string[]]$c, [string]$nom) {
    $n = ($nom -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
    $a = { param([string[]]$noms) foreach ($x in $noms) { if ($c -contains $x) { return $true } }; return $false }
    # rapports agrégés (ReportAggregator) : classés d'après leurs colonnes, comme les rapports unitaires
    if ($n -like '*structurefailuresummary*') { return 'StructureFailureSummary' }
    if (& $a @('structuretype')) { if ($n -like '*failure*') { return 'StructureFailureReport' } else { return 'StructureReport' } }
    if (& $a @('expectedmigratedfilecount', 'failedreading', 'failedpacking')) { return 'ItemSummary' }
    if ((& $a @('totalscannedfolders')) -and (& $a @('folderswithissues', 'itemswithissues'))) { return 'ScanSummary' }
    if ((& $a @('sourcepath')) -and (& $a @('scanstatus', 'migrationreadiness', 'maxpathlength', 'rootpermissions'))) { return 'ScanReport' }
    $estElement = (& $a @('filename', 'itemname', 'resultcategory')) -and (& $a @('source')) -and (& $a @('status'))
    if ($estElement) {
        if ($n -like '*failuresummary*') { return 'FailureSummary' }
        if ($n -like '*failure*') { return 'ItemFailureReport' }
        return 'ItemReport'
    }
    if ((& $a @('taskname', 'taskid')) -and (& $a @('totalscanneditems', 'migrateditems', 'totaltobemigrateditems', 'migrateditemsincurrentround'))) { return 'SummaryReport' }
    if ($n -like '*scan*' -and (& $a @('sourcepath', 'source', 'path', 'filepath'))) { return 'ScanLog' }
    return $null
}

function Get-TaskId([string]$Chemin, [string[]]$c, [string]$contexte) {
    $idx = [array]::IndexOf($c, 'taskid')
    if ($idx -ge 0) {
        try {
            $l = Import-Csv -LiteralPath $Chemin -Encoding UTF8 | Select-Object -First 1
            $v = @($l.PSObject.Properties)[$idx].Value
            if ($v) { return [string]$v }
        } catch { }
    }
    $m = $reGuid.Match($contexte)
    if ($m.Success) { return $m.Value }
    return $null
}

$candidats = New-Object System.Collections.Generic.List[object]
function Add-Candidat([string]$Fichier, [string]$Origine, [string]$CheminOrigine, [bool]$Explicite) {
    $candidats.Add([pscustomobject]@{ Fichier = $Fichier; Origine = $Origine; CheminOrigine = $CheminOrigine; Explicite = $Explicite })
}
function Expand-ZipCsv([string]$Zip, [string]$Origine, [string]$Contexte, [bool]$Explicite, [int]$Profondeur = 0) {
    $dest = Join-Path $travail ('zip_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Zip)
        try {
            foreach ($e in $archive.Entries) {
                $ext = [System.IO.Path]::GetExtension($e.FullName).ToLowerInvariant()
                if ($ext -ne '.csv' -and -not ($ext -eq '.zip' -and $Profondeur -lt 2)) { continue }
                $relatif = ($e.FullName -replace '[\\/:*?"<>|]+', '_')
                $cible = Join-Path $dest $relatif
                New-Item -ItemType Directory -Path (Split-Path $cible -Parent) -Force | Out-Null
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($e, $cible, $true)
                if ($ext -eq '.zip') { Expand-ZipCsv $cible $Origine "$Contexte!$($e.FullName)" $Explicite ($Profondeur + 1) }
                else { Add-Candidat $cible $Origine "$Contexte!$($e.FullName)" $Explicite }
            }
        } finally { $archive.Dispose() }
    } catch { Write-Warning "Archive illisible ignorée : $Zip ($($_.Exception.Message))" }
}

$limite = (Get-Date).AddDays(-$DepuisJours)
foreach ($s in $sources) {
    $gci = @{ LiteralPath = $s.Chemin; File = $true; ErrorAction = 'SilentlyContinue' }
    if ($s.Recursif) { $gci.Recurse = $true }
    $fichiers = @(Get-ChildItem @gci | Where-Object { $_.Extension -in '.csv', '.zip' })
    if ($s.Filtrer) { $fichiers = @($fichiers | Where-Object { $_.LastWriteTime -ge $limite }) }
    if ($s.Origine -like 'Agent*') { $fichiers = @($fichiers | Where-Object { $_.Extension -eq '.csv' }) }
    $explicite = (-not $s.Filtrer) -and ($s.Origine -notlike 'Agent*')
    foreach ($f in $fichiers) {
        if ($f.Extension -eq '.zip') { Expand-ZipCsv $f.FullName $s.Origine $f.FullName $explicite }
        else { Add-Candidat $f.FullName $s.Origine $f.FullName $explicite }
    }
    Write-Host ("  {0,-28} {1,5} fichier(s) examiné(s) - {2}" -f $s.Origine, $fichiers.Count, $s.Chemin)
}

# ---------------------------------------------------------------- Copie normalisée
$vus = @{}
$index = New-Object System.Collections.Generic.List[object]
$ignores = 0
$deja = 0
$typesTache = @('ItemReport', 'ItemFailureReport', 'ItemSummary', 'ScanSummary', 'StructureReport', 'StructureFailureReport', 'FailureSummary')

# 1re passe : type de chaque CSV, Task ID quand la colonne existe
$analyses = New-Object System.Collections.Generic.List[object]
$tacheParArchive = @{}
foreach ($cand in $candidats) {
    try { $colonnes = Get-Colonnes $cand.Fichier } catch { continue }
    $nomOrigine = [System.IO.Path]::GetFileName(($cand.CheminOrigine -split '!')[-1])
    $type = Get-TypeRapport $colonnes $nomOrigine
    if (-not $type -and -not $cand.Explicite) { $ignores++; continue }   # CSV sans rapport avec la migration dans Téléchargements
    $archive = ($cand.CheminOrigine -split '!')[0]
    $taskId = $null
    if ($type -in $typesTache) {
        $taskId = Get-TaskId $cand.Fichier $colonnes $(if ($cand.CheminOrigine -like '*!*') { $cand.CheminOrigine } else { '' })
        if ($taskId -and $cand.CheminOrigine -like '*!*' -and -not $tacheParArchive.ContainsKey($archive)) { $tacheParArchive[$archive] = $taskId }
    }
    $analyses.Add([pscustomobject]@{ Cand = $cand; Type = $type; TaskId = $taskId; Archive = $archive; NomOrigine = $nomOrigine })
}

# 2e passe : les rapports sans colonne Task ID (ItemSummary, ScanSummary, StructureReport) héritent de la tâche du même zip
foreach ($an in $analyses) {
    $cand = $an.Cand; $type = $an.Type; $nomOrigine = $an.NomOrigine; $taskId = $an.TaskId
    if ($nomOrigine -match '(?i)aggregat') { $taskId = 'Agreges' }        # ReportAggregator : un fichier pour toutes les tâches
    if (-not $taskId -and $type -in $typesTache) {
        if ($tacheParArchive.ContainsKey($an.Archive)) { $taskId = $tacheParArchive[$an.Archive] }
        else { $m = $reGuid.Match($cand.CheminOrigine); if ($m.Success) { $taskId = $m.Value } }
    }

    $fs = [System.IO.File]::OpenRead($cand.Fichier)
    try { $empreinte = [System.BitConverter]::ToString($sha.ComputeHash($fs)).Replace('-', '').ToLowerInvariant() } finally { $fs.Dispose() }
    $cle = "$taskId|$type|$empreinte"            # même contenu possible pour deux tâches ou deux types (ex. relance 100 % en échec)
    if ($vus.ContainsKey($cle)) { continue }
    $vus[$cle] = $true

    $court = $empreinte.Substring(0, 12)
    $dossierTache = Join-Path (Join-Path $cibleMM 'Taches') $(if ($taskId) { $taskId } else { 'SansTaskID' })
    if (-not $type) { $dossier = Join-Path $cibleMM 'NonReconnus'; $nom = '{0}_{1}' -f $court, $nomOrigine }
    elseif ($type -eq 'SummaryReport' -or $type -eq 'AggregateSummary') { $dossier = Join-Path $cibleMM 'Synthese'; $nom = "${type}_$court.csv" }
    elseif ($type -eq 'ScanReport' -or $type -eq 'ScanLog') { $dossier = Join-Path $cibleMM 'Scans'; $nom = "${type}_$court.csv" }
    else { $dossier = $dossierTache; $nom = "${type}_$court.csv" }
    $dest = Join-Path $dossier $nom
    if (Test-Path -LiteralPath $dest) { $deja++; continue }     # déjà collecté lors d'une exécution précédente
    New-Item -ItemType Directory -Path $dossier -Force | Out-Null
    Copy-Item -LiteralPath $cand.Fichier -Destination $dest -Force

    $nbLignes = 0
    $sr = New-Object System.IO.StreamReader($dest)
    try { while ($null -ne $sr.ReadLine()) { $nbLignes++ } } finally { $sr.Dispose() }
    $index.Add([pscustomobject][ordered]@{
        TypeRapport   = $(if ($type) { $type } else { 'Non reconnu' })
        Origine       = $cand.Origine
        TaskID        = $taskId
        Fichier       = $nom
        FichierOrigine = $nomOrigine
        Lignes        = [Math]::Max(0, $nbLignes - 1)
        DateFichier   = (Get-Item -LiteralPath $cand.Fichier).LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss')
        DateCollecte  = $debut.ToString('yyyy-MM-ddTHH:mm:ss')
        Empreinte     = $empreinte
        CheminCopie   = $dest
        CheminOrigine = $cand.CheminOrigine
    })
}

# Index cumulatif : les collectes précédentes sont conservées
$fichierIndex = Join-Path $Sortie 'Index_Rapports_MigrationManager.csv'
$anciens = @()
if (Test-Path -LiteralPath $fichierIndex) {
    $anciens = @(Import-Csv -LiteralPath $fichierIndex -Encoding UTF8 | Where-Object { $e = $_.Empreinte; -not ($index | Where-Object { $_.Empreinte -eq $e }) -and (Test-Path -LiteralPath $_.CheminCopie) })
}
$tout = @($anciens) + $index.ToArray()
$tout | Export-Csv -LiteralPath $fichierIndex -NoTypeInformation -Encoding UTF8

if ($AvecJournaux -and $DossierAgent -and (Test-Path -LiteralPath $DossierAgent)) {
    $logs = @(Get-ChildItem -LiteralPath $DossierAgent -Recurse -File -Include '*.log', '*.txt' -ErrorAction SilentlyContinue)
    $racineAgent = (Resolve-Path -LiteralPath $DossierAgent).ProviderPath.TrimEnd('\', '/')
    foreach ($l in $logs) {
        $dest = Join-Path (Join-Path $Sortie 'MigrationManager_Journaux') $l.FullName.Substring($racineAgent.Length + 1)
        New-Item -ItemType Directory -Path (Split-Path $dest -Parent) -Force | Out-Null
        Copy-Item -LiteralPath $l.FullName -Destination $dest -Force
    }
    Write-Host ("{0} fichier(s) journal copiés." -f $logs.Count)
}
Remove-Item -LiteralPath $travail -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------- Bilan
Write-Host ''
Write-Host '=== Rapports Migration Manager collectés ===' -ForegroundColor Green
if ($index.Count -eq 0 -and $deja -gt 0) {
    Write-Host 'Aucun nouveau rapport depuis la dernière collecte.' -ForegroundColor Yellow
} elseif ($index.Count -eq 0) {
    Write-Warning ("Aucun nouveau rapport. Téléchargez les rapports dans Migration center (Download task report / Summary report) " +
        "vers $DossierImport, ou installez le module PowerShell Migration Manager.")
} else {
    $index | Group-Object TypeRapport | Sort-Object Name | ForEach-Object {
        Write-Host ("  {0,-24} {1,3} fichier(s)  {2,7:N0} lignes" -f $_.Name, $_.Count, ($_.Group | Measure-Object Lignes -Sum).Sum)
    }
    Write-Host ("Tâches : {0}" -f @($index | Where-Object { $_.TaskID -and $_.TaskID -ne 'Agreges' } | Select-Object -ExpandProperty TaskID -Unique).Count)
    if (@($index | Where-Object TaskID -eq 'Agreges').Count) { Write-Host 'Rapports agrégés (ReportAggregator) rangés dans MigrationManager\Taches\Agreges.' }
}
if ($deja) { Write-Host "$deja rapport(s) déjà présent(s) dans le dossier de sortie, non recopié(s)." -ForegroundColor DarkGray }
if ($ignores) { Write-Host "$ignores CSV sans lien avec Migration Manager ignoré(s) dans les téléchargements." -ForegroundColor DarkGray }

if ($Zip) {
    $dossierZip = Split-Path -Path $Sortie -Parent
    if (-not $dossierZip) { $dossierZip = $Sortie }
    $zipPath = Join-Path $dossierZip ("Rapports_Migration_{0:yyyyMMdd_HHmm}.zip" -f (Get-Date))
    $aArchiver = Get-ChildItem -LiteralPath $Sortie -Force | Where-Object { $_.Name -notlike 'Rapports_Migration_*.zip' }
    Compress-Archive -LiteralPath $aArchiver.FullName -DestinationPath $zipPath -Force
    Write-Host "Archive : $zipPath"
}
Write-Host "Dossier Rapports : $Sortie"
