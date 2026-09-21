<#
.SYNOPSIS
    Inventaire du serveur de fichiers source avant migration avec Migration Manager, et préparation du
    fichier de tâches en masse (bulk) à charger dans le centre de migration SharePoint.

.DESCRIPTION
    Parcourt tous les fichiers (y compris cachés et chemins > 260 caractères) et produit dans le dossier de sortie :
      Inventaire_Source.csv               un fichier par ligne, chemin UNC (vu par l'agent) + anomalies SharePoint Online
      Correspondance_Bibliotheques.csv    dossier racine -> bibliothèque cible (chemins UNC)
      MigrationManager_Taches.csv         6 colonnes, à charger dans Migration center > Migrations > Add task > Bulk migration
      Parametres_Migration_Manager.txt    réglages recommandés pour les tâches (extensions exclues, fichiers cachés...)

    Migration Manager n'accepte que des chemins UNC (\\serveur\partage). Si -CheminUNC n'est pas fourni, le script
    cherche le partage SMB qui contient -Source (créé par 01_Preparer-ServeurFichiers.ps1).

    Au lancement, une fenêtre Windows demande le dossier de sortie (proposé : -Sortie).
    -SansFenetre utilise -Sortie directement (exécution planifiée ou automatisée).

.EXAMPLE
    .\02_Inventorier-Source.ps1
.EXAMPLE
    .\02_Inventorier-Source.ps1 -Source 'D:\Partages\Contoso' -CheminUNC '\\SRV-FICHIERS\Contoso' -SiteUrl 'https://contoso.sharepoint.com/sites/Migration' -SansHash
#>
[CmdletBinding()]
param(
    [string]$FichierParametres = $(if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Parametres_Client.psd1' } else { 'Parametres_Client.psd1' }),
    [hashtable]$ParametresClient = $(if (Test-Path -LiteralPath $FichierParametres) { Import-PowerShellDataFile -LiteralPath $FichierParametres } else { @{} }),
    [string]$Racine = $(if ($ParametresClient['CheminLocal']) { Split-Path -Parent $ParametresClient['CheminLocal'] } else { 'C:\DemoMigration' }),
    [string]$Source = $(if ($ParametresClient['CheminLocal']) { $ParametresClient['CheminLocal'] } else { (Join-Path $Racine 'FileServer') }),
    [string]$CheminUNC = [string]$ParametresClient['CheminUNC'],
    [string]$Sortie = $(if ($ParametresClient['DossierRapports']) { $ParametresClient['DossierRapports'] } else { (Join-Path $Racine 'Rapports') }),
    [string]$SiteUrl = $(if ($ParametresClient['SiteUrl']) { $ParametresClient['SiteUrl'] } else { 'https://m365x71797824.sharepoint.com/sites/MigrationFileServer' }),
    [System.Collections.IDictionary]$Correspondance = [ordered]@{
        '01_Clients'           = 'Clients'
        '02_RH_Collaborateurs' = 'RH'
        '03_Finance'           = 'Finance'
        '04_Direction'         = 'Direction'
        '05_Commun'            = 'Commun'
        '06_Archives'          = 'Archives'
    },
    [string[]]$ExtensionsExclues = @('.tmp', '.bak', '.pst'),
    [switch]$SansHash,
    [switch]$AvecEntete,
    [switch]$SansFenetre
)

$ErrorActionPreference = 'Stop'

# Paramètres de la mission (Scripts\Parametres_Client.psd1, écrit par 00_Configurer-Parametres.ps1).
# Les valeurs du fichier servent de valeurs par défaut ; un paramètre en ligne de commande reste prioritaire.
if ($ParametresClient.Count -gt 0) {
    $nomMission = if ($ParametresClient['Client']) { $ParametresClient['Client'] } else { '(client non renseigné)' }
    Write-Host ("Mission : {0}   [{1}]" -f $nomMission, (Split-Path -Leaf $FichierParametres)) -ForegroundColor DarkCyan
}
$isWin = [System.Environment]::OSVersion.Platform -eq 'Win32NT'
$sep = [System.IO.Path]::DirectorySeparatorChar
function Get-LongPath([string]$Path) { if ($isWin -and -not $Path.StartsWith('\\?\')) { if ($Path.StartsWith('\\')) { '\\?\UNC\' + $Path.Substring(2) } else { '\\?\' + $Path } } else { $Path } }
function Remove-LongPrefix([string]$Path) { if ($Path.StartsWith('\\?\UNC\')) { '\\' + $Path.Substring(8) } elseif ($Path.StartsWith('\\?\')) { $Path.Substring(4) } else { $Path } }

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

# ---------------------------------------------------------------- Chemin UNC (Migration Manager n'accepte pas les chemins locaux)
function Resolve-CheminUNC {
    <#  Retourne le chemin UNC (\\ordinateur\partage\...) correspondant à un dossier local, en cherchant
        le partage SMB le plus précis qui le contient. Retourne $null si aucun partage n'est trouvé.  #>
    param([string]$CheminLocal)
    if ($CheminLocal.StartsWith('\\')) { return $CheminLocal.TrimEnd('\') }
    if ([System.Environment]::OSVersion.Platform -ne 'Win32NT' -or -not (Get-Command Get-SmbShare -ErrorAction SilentlyContinue)) { return $null }
    $cible = $CheminLocal.TrimEnd('\')
    $meilleur = $null
    foreach ($s in @(Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { $_.Path -and -not $_.Special -and $_.Name -notlike '*$' })) {
        $p = $s.Path.TrimEnd('\')
        if ($cible -eq $p -or $cible.StartsWith($p + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
            if (-not $meilleur -or $p.Length -gt $meilleur.Path.TrimEnd('\').Length) { $meilleur = $s }
        }
    }
    if (-not $meilleur) { return $null }
    $reste = $cible.Substring($meilleur.Path.TrimEnd('\').Length)
    return ('\\{0}\{1}{2}' -f [System.Environment]::MachineName, $meilleur.Name, $reste)
}

# ---------------------------------------------------------------- Préparation
if (-not (Test-Path -LiteralPath $Source)) { throw "Dossier source introuvable : $Source" }
$Source = (Resolve-Path -LiteralPath $Source).ProviderPath.TrimEnd($sep)
if (-not $CheminUNC) { $CheminUNC = Resolve-CheminUNC $Source }
if (-not $CheminUNC) {
    $CheminUNC = '\\{0}\{1}' -f [System.Environment]::MachineName, (Split-Path -Path $Source -Leaf)
    Write-Warning "Aucun partage SMB trouvé pour $Source. Chemin supposé : $CheminUNC (créez le partage ou utilisez -CheminUNC)."
}
$CheminUNC = $CheminUNC.TrimEnd('\')

$Sortie = Select-DossierSortie -Titre 'Choisissez le dossier de sortie de l''inventaire source (Inventaire_Source.csv, MigrationManager_Taches.csv...)' -DossierPropose $Sortie -SansFenetre:$SansFenetre
if (-not $Sortie) { Write-Warning 'Sélection annulée : aucun fichier produit.'; return }
New-Item -ItemType Directory -Path $Sortie -Force | Out-Null
$sortiePleine = [System.IO.Path]::GetFullPath($Sortie).TrimEnd($sep)
if (($sortiePleine + $sep).StartsWith($Source + $sep, [System.StringComparison]::OrdinalIgnoreCase)) {
    Write-Warning "Le dossier de sortie est situé dans la source : les CSV produits seraient eux-mêmes migrés et inventoriés."
}

$sitePath = ([uri]$SiteUrl).AbsolutePath.TrimEnd('/')          # /sites/MigrationFileServer
$nomsInterdits = @('.lock', 'CON', 'PRN', 'AUX', 'NUL', 'desktop.ini') + (0..9 | ForEach-Object { "COM$_"; "LPT$_" })
$sha = [System.Security.Cryptography.SHA256]::Create()
function Get-CheminUNCFichier([string]$relatif) { $CheminUNC + '\' + ($relatif -replace [regex]::Escape([string]$sep), '\') }

# ---------------------------------------------------------------- Inventaire
Write-Host "Inventaire de $Source (vu par l'agent comme $CheminUNC) ..." -ForegroundColor Cyan
$lignes = New-Object System.Collections.Generic.List[object]
$enum = [System.IO.Directory]::EnumerateFiles((Get-LongPath $Source), '*', [System.IO.SearchOption]::AllDirectories)
foreach ($lp in $enum) {
    $fi = New-Object System.IO.FileInfo($lp)
    $plein = Remove-LongPrefix $fi.FullName
    $relatif = $plein.Substring($Source.Length + 1)
    $premier = $relatif.Split($sep)[0]
    $biblio = if ($Correspondance.Contains($premier)) { $Correspondance[$premier] } else { '(non mappé)' }
    $sousChemin = if ($relatif.Contains($sep)) { $relatif.Substring($premier.Length + 1) } else { $relatif }
    $cheminCible = "$sitePath/$biblio/" + ($sousChemin -replace [regex]::Escape([string]$sep), '/')
    $ext = $fi.Extension.ToLowerInvariant()
    $cache = [bool]($fi.Attributes -band [System.IO.FileAttributes]::Hidden)
    $systeme = [bool]($fi.Attributes -band [System.IO.FileAttributes]::System)
    $unc = Get-CheminUNCFichier $relatif

    $anomalies = @()
    if ($cheminCible.Length -gt 400) { $anomalies += 'Chemin > 400 caractères' }
    if ($fi.Name.StartsWith('~$') -or $fi.Name -like '*_vti_*' -or $nomsInterdits -contains $fi.BaseName -or $nomsInterdits -contains $fi.Name) { $anomalies += 'Nom interdit SharePoint' }
    if ($fi.Name -ne $fi.Name.Trim() -or $fi.Name.EndsWith('.')) { $anomalies += 'Espace ou point en début/fin de nom' }
    if ($cache -or $systeme) { $anomalies += 'Fichier caché/système (migré par défaut)' }
    if ($ExtensionsExclues -contains $ext) { $anomalies += 'Extension exclue' }
    if ($fi.Length -eq 0) { $anomalies += 'Fichier vide' }
    if ($fi.Length -gt 250GB) { $anomalies += 'Fichier > 250 Go' }

    $hash = ''
    if (-not $SansHash) {
        try { $s = $fi.OpenRead(); try { $hash = [System.BitConverter]::ToString($sha.ComputeHash($s)).Replace('-', '') } finally { $s.Dispose() } } catch { $hash = 'ERREUR' }
    }

    $lignes.Add([pscustomobject][ordered]@{
        CheminComplet       = $unc                  # chemin tel qu'il apparaît dans les rapports Migration Manager
        CheminLocal         = $plein
        CheminRelatif       = $relatif
        DossierRacine       = $premier
        Bibliotheque        = $biblio
        CheminCible         = $cheminCible
        Nom                 = $fi.Name
        Extension           = $ext
        TailleOctets        = $fi.Length
        DateCreation        = $fi.CreationTime.ToString('yyyy-MM-ddTHH:mm:ss')
        DateModification    = $fi.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss')
        Cache               = $cache
        Systeme             = $systeme
        LongueurChemin      = $unc.Length
        LongueurCheminCible = $cheminCible.Length
        Profondeur          = $relatif.Split($sep).Count - 1
        Anomalies           = ($anomalies -join ' | ')
        HashSHA256          = $hash
    })
}

# ---------------------------------------------------------------- Sorties
$lignes | Export-Csv -LiteralPath (Join-Path $Sortie 'Inventaire_Source.csv') -NoTypeInformation -Encoding UTF8

$Correspondance.GetEnumerator() | ForEach-Object {
    [pscustomobject][ordered]@{
        DossierSource = $_.Key; Bibliotheque = $_.Value
        CheminSource  = (Get-CheminUNCFichier $_.Key)
        CheminLocal   = (Join-Path $Source $_.Key)
        UrlCible      = "$SiteUrl/$($_.Value)"
    }
} | Export-Csv -LiteralPath (Join-Path $Sortie 'Correspondance_Bibliotheques.csv') -NoTypeInformation -Encoding UTF8

# Fichier de tâches en masse Migration Manager (partages de fichiers) : 6 colonnes
# FileSharePath, (vide), (vide), SharePointSite, DocLibrary, DocSubFolder
$bulk = New-Object System.Collections.Generic.List[string]
if ($AvecEntete) { $bulk.Add('FileSharePath,,,SharePointSite,DocLibrary,DocSubFolder') }
foreach ($e in $Correspondance.GetEnumerator()) {
    $chemin = Get-CheminUNCFichier $e.Key
    if (-not (Test-Path -LiteralPath (Join-Path $Source $e.Key))) { Write-Warning "Dossier absent de la source, ligne ignorée : $($e.Key)"; continue }
    $bulk.Add(('{0},,,{1},{2},' -f $chemin, $SiteUrl, $e.Value))
}
[System.IO.File]::WriteAllLines((Join-Path $Sortie 'MigrationManager_Taches.csv'), [string[]]$bulk, (New-Object System.Text.UTF8Encoding($false)))

$extMM = (($ExtensionsExclues | ForEach-Object { $_.TrimStart('.').ToUpperInvariant() }) -join ':')
$nbCaches = @($lignes | Where-Object { $_.Cache -or $_.Systeme }).Count
@"
Réglages recommandés pour les tâches Migration Manager (Add task > Settings > All settings)
Généré le $(Get-Date -Format 'dd/MM/yyyy HH:mm') pour $CheminUNC

Mission
  Client / environnement ........................... $(if ($ParametresClient['Client']) { $ParametresClient['Client'] } else { '(non renseigné)' })
  Centre d'administration SharePoint ............... $(if ($ParametresClient['CentreAdministration']) { $ParametresClient['CentreAdministration'] } else { '(non renseigné)' })
  Site cible ....................................... $SiteUrl
  Source (UNC / local) ............................. $CheminUNC  /  $Source
  Dossier des rapports ............................. $Sortie

General
  Only perform scanning ............................ Off (On pour un passage de pré-évaluation)
  Preserve file share permissions .................. On

Users
  Microsoft Entra lookup ........................... On
  User mapping file ................................ (aucun, sauf comptes non synchronisés)

Filters
  Migrate hidden files ............................. On (défaut) ; Off pour exclure les $nbCaches fichier(s) caché(s)/système de la source
  Don't migrate files with these extensions ........ $extMM
  Migrate files and folders with invalid characters  On (défaut)

Advanced
  Migration auto rerun ............................. Off
  Migration Manager working folder ................. %appdata%\Microsoft\SPMigration (150 Go libres minimum)

Fichier de tâches en masse : $(Join-Path $Sortie 'MigrationManager_Taches.csv')
"@ | Set-Content -LiteralPath (Join-Path $Sortie 'Parametres_Migration_Manager.txt') -Encoding UTF8

$anom = $lignes | Where-Object Anomalies
Write-Host ''
Write-Host '=== Inventaire source terminé ===' -ForegroundColor Green
Write-Host ("Fichiers : {0:N0}   Volume : {1:N1} Mo   Fichiers avec anomalie : {2}" -f $lignes.Count, (($lignes | Measure-Object TailleOctets -Sum).Sum / 1MB), @($anom).Count)
$lignes | Group-Object Bibliotheque | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-12} {1,5} fichiers  {2,8:N1} Mo" -f $_.Name, $_.Count, (($_.Group | Measure-Object TailleOctets -Sum).Sum / 1MB))
}
Write-Host "Chemin UNC            : $CheminUNC"
Write-Host "Tâches en masse       : $(Join-Path $Sortie 'MigrationManager_Taches.csv') ($(@($bulk | Where-Object { $_ -notlike 'FileSharePath*' }).Count) tâches)"
Write-Host "Extensions à exclure  : $extMM"
Write-Host "Sorties               : $Sortie"
