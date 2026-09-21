<#
.SYNOPSIS
    Crée un faux « serveur de fichiers » Contoso France (données 100 % fictives) et le partage SMB
    nécessaire pour tester la migration vers SharePoint Online avec Migration Manager.

.DESCRIPTION
    1. Décompresse FileServer_Contoso.zip (≈ 1 360 fichiers Office/PDF/images) dans le dossier cible.
    2. Réapplique les dates de création / modification d'origine (2009 → 2026).
    3. Ajoute des cas volontairement problématiques pour la migration (chemin > 400 caractères,
       noms interdits dans SharePoint, fichiers cachés, temporaires, .bak/.pst, fichier vide...).
    4. Écrit la liste des cas plantés dans <Racine>\Cas_de_test.csv.
    5. Crée le partage SMB \\<ordinateur>\<NomPartage> en lecture : Migration Manager n'accepte que des
       chemins UNC. Nécessite une console PowerShell « Exécuter en tant qu'administrateur » (sinon la
       commande à lancer est affichée). -SansPartage pour ne pas le créer.

    Au lancement, une fenêtre Windows demande le dossier de destination (<Racine>) : le serveur de
    démonstration est créé dans <Racine>\FileServer. Le paramètre -Racine sert de dossier proposé ;
    -SansFenetre utilise -Racine directement (exécution planifiée ou automatisée).

    Le compte Windows saisi lors de l'installation de l'agent Migration Manager doit pouvoir lire le partage
    (par défaut : Utilisateurs authentifiés en lecture, Administrateurs en contrôle total).

    Compatible Windows PowerShell 5.1 et PowerShell 7.

.EXAMPLE
    .\01_Preparer-ServeurFichiers.ps1
.EXAMPLE
    .\01_Preparer-ServeurFichiers.ps1 -FichierVolumineuxMo 500
.EXAMPLE
    .\01_Preparer-ServeurFichiers.ps1 -Supprimer      # nettoie tout (y compris les chemins longs)
.EXAMPLE
    .\01_Preparer-ServeurFichiers.ps1 -Racine 'D:\Demo' -SansFenetre
.EXAMPLE
    .\01_Preparer-ServeurFichiers.ps1 -NomPartage 'Contoso_FS' -ComptesLecture 'CONTOSO\svc-migration'
#>
[CmdletBinding()]
param(
    [string]$FichierParametres = $(if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Parametres_Client.psd1' } else { 'Parametres_Client.psd1' }),
    [hashtable]$ParametresClient = $(if (Test-Path -LiteralPath $FichierParametres) { Import-PowerShellDataFile -LiteralPath $FichierParametres } else { @{} }),
    [string]$Racine = $(if ($ParametresClient['CheminLocal']) { Split-Path -Parent $ParametresClient['CheminLocal'] } else { 'C:\DemoMigration' }),
    [string]$ZipPath = $(if ($PSScriptRoot) { Join-Path $PSScriptRoot '..\Donnees' } else { 'Donnees' }),   # dossier des archives FileServer_Contoso*.zip ou archive unique
    [int]$FichierVolumineuxMo = 0,
    [string]$NomPartage = $(if ($ParametresClient['CheminUNC']) { ($ParametresClient['CheminUNC'].TrimEnd('\') -split '\\')[-1] } else { 'FileServer' }),
    [string[]]$ComptesLecture = @(),
    [switch]$SansPartage,
    [switch]$Supprimer,
    [switch]$SansFenetre
)

$ErrorActionPreference = 'Stop'

# Paramètres de la mission (Scripts\Parametres_Client.psd1, écrit par 00_Configurer-Parametres.ps1).
# Les valeurs du fichier servent de valeurs par défaut ; un paramètre en ligne de commande reste prioritaire.
if ($ParametresClient.Count -gt 0) {
    $nomMission = if ($ParametresClient['Client']) { $ParametresClient['Client'] } else { '(client non renseigné)' }
    Write-Host ("Mission : {0}   [{1}]" -f $nomMission, (Split-Path -Leaf $FichierParametres)) -ForegroundColor DarkCyan
}

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

# ---------------------------------------------------------------- Partage SMB
$isWin = [System.Environment]::OSVersion.Platform -eq 'Win32NT'
function Get-NomCompte([string]$Sid) {
    try { (New-Object System.Security.Principal.SecurityIdentifier $Sid).Translate([System.Security.Principal.NTAccount]).Value } catch { $null }
}
function Test-Administrateur {
    if (-not $isWin) { return $false }
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object System.Security.Principal.WindowsPrincipal $id).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}
function New-PartageDemo([string]$Nom, [string]$Chemin, [string[]]$Lecture) {
    if (-not $isWin -or -not (Get-Command New-SmbShare -ErrorAction SilentlyContinue)) {
        Write-Warning 'Partage SMB non créé : Windows et le module SmbShare sont nécessaires.'
        return $null
    }
    $admins = Get-NomCompte 'S-1-5-32-544'                      # Administrateurs (nom localisé)
    if (-not $Lecture -or $Lecture.Count -eq 0) { $Lecture = @(Get-NomCompte 'S-1-5-11') }   # Utilisateurs authentifiés
    $unc = '\\{0}\{1}' -f [System.Environment]::MachineName, $Nom
    $existant = Get-SmbShare -Name $Nom -ErrorAction SilentlyContinue
    if ($existant) {
        if ($existant.Path.TrimEnd('\') -ieq $Chemin.TrimEnd('\')) { Write-Host "Partage existant réutilisé : $unc" -ForegroundColor Cyan; return $unc }
        Write-Warning "Le partage « $Nom » existe déjà vers $($existant.Path). Choisissez un autre -NomPartage."
        return $null
    }
    $commande = "New-SmbShare -Name '$Nom' -Path '$Chemin' -ReadAccess '$($Lecture -join "','")' -FullAccess '$admins'"
    if (-not (Test-Administrateur)) {
        Write-Warning "Partage SMB non créé : relancez PowerShell en tant qu'administrateur, ou exécutez :`n    $commande"
        return $null
    }
    New-SmbShare -Name $Nom -Path $Chemin -ReadAccess $Lecture -FullAccess $admins -Description 'Démo migration Migration Manager (données fictives)' | Out-Null
    Write-Host "Partage SMB créé : $unc (lecture : $($Lecture -join ', '))" -ForegroundColor Cyan
    return $unc
}

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

if ($Supprimer) { $titre = 'Choisissez le dossier qui contient le serveur de démonstration à supprimer (sous-dossier FileServer)' }
else { $titre = 'Choisissez le dossier de destination du serveur de démonstration (un sous-dossier FileServer y sera créé)' }
$Racine = Select-DossierSortie -Titre $titre -DossierPropose $Racine -SansFenetre:$SansFenetre
if (-not $Racine) { Write-Warning 'Sélection annulée : aucune action effectuée.'; return }
Write-Host "Dossier de destination : $Racine" -ForegroundColor Cyan
$FileServer = Join-Path $Racine 'FileServer'

function Get-LongPath([string]$Path) {
    # Préfixe \\?\ : contourne la limite Windows de 260 caractères
    if ([System.Environment]::OSVersion.Platform -ne 'Win32NT' -or $Path.StartsWith('\\?\')) { return $Path }
    return '\\?\' + $Path
}

function New-RandomFile([string]$Path, [int]$SizeMo) {
    $dir = [System.IO.Path]::GetDirectoryName($Path)
    [System.IO.Directory]::CreateDirectory((Get-LongPath $dir)) | Out-Null
    $buffer = New-Object byte[] (1MB)
    $rng = New-Object System.Random
    $fs = [System.IO.File]::Create((Get-LongPath $Path))
    try {
        for ($i = 0; $i -lt $SizeMo; $i++) { $rng.NextBytes($buffer); $fs.Write($buffer, 0, $buffer.Length) }
    } finally { $fs.Dispose() }
}

function Set-Dates([string]$Path, [datetime]$Created, [datetime]$Modified) {
    $lp = Get-LongPath $Path
    [System.IO.File]::SetCreationTime($lp, $Created)
    [System.IO.File]::SetLastWriteTime($lp, $Modified)
}

# ---------------------------------------------------------------- Suppression
if ($Supprimer) {
    if ($isWin -and (Get-Command Get-SmbShare -ErrorAction SilentlyContinue)) {
        $partages = @(Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { $_.Path -and $_.Path.TrimEnd('\') -ieq $FileServer.TrimEnd('\') })
        foreach ($s in $partages) {
            if (Test-Administrateur) { Remove-SmbShare -Name $s.Name -Force; Write-Host "Partage SMB $($s.Name) supprimé." -ForegroundColor Yellow }
            else { Write-Warning "Partage SMB $($s.Name) conservé : relancez en administrateur pour le supprimer (Remove-SmbShare -Name '$($s.Name)')." }
        }
    }
    if (Test-Path -LiteralPath $FileServer) {
        Write-Host "Suppression de $FileServer ..." -ForegroundColor Yellow
        # Retire les attributs caché/système pour pouvoir supprimer
        Get-ChildItem -LiteralPath $FileServer -Recurse -Force -ErrorAction SilentlyContinue |
            ForEach-Object { try { $_.Attributes = 'Normal' } catch { } }
        [System.IO.Directory]::Delete((Get-LongPath $FileServer), $true)
    }
    $cas = Join-Path $Racine 'Cas_de_test.csv'
    if (Test-Path -LiteralPath $cas) { Remove-Item -LiteralPath $cas -Force }
    Write-Host 'Jeu de données supprimé.' -ForegroundColor Green
    return
}

# ---------------------------------------------------------------- Décompression
if (-not (Test-Path -LiteralPath $ZipPath)) { throw "Archive ou dossier introuvable : $ZipPath" }
if (Test-Path -LiteralPath $ZipPath -PathType Container) {
    $archives = @(Get-ChildItem -LiteralPath $ZipPath -Filter 'FileServer_Contoso*.zip' -File | Sort-Object Name)
} else {
    $archives = @(Get-Item -LiteralPath $ZipPath)
}
if ($archives.Count -eq 0) { throw "Aucune archive FileServer_Contoso*.zip dans $ZipPath (extrayez les deux zips du kit dans le même dossier)." }
if (Test-Path -LiteralPath $FileServer) {
    throw "$FileServer existe déjà. Relancez avec -Supprimer pour repartir de zéro."
}
New-Item -ItemType Directory -Path $FileServer -Force | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
foreach ($a in $archives) {
    Write-Host "Décompression de $($a.Name) vers $FileServer ..." -ForegroundColor Cyan
    [System.IO.Compression.ZipFile]::ExtractToDirectory($a.FullName, $FileServer)
}

# ---------------------------------------------------------------- Dates d'origine
$manifestPath = Join-Path $FileServer 'manifest.csv'
if (-not (Test-Path -LiteralPath $manifestPath)) { throw "manifest.csv absent : la partie 1 du jeu de données (FileServer_Contoso_1.zip) manque." }
$manifest = Import-Csv -LiteralPath $manifestPath -Delimiter ';' -Encoding UTF8
$culture = [System.Globalization.CultureInfo]::InvariantCulture
$n = 0
$manquants = 0
foreach ($m in $manifest) {
    $p = Join-Path $FileServer $m.RelativePath
    if (Test-Path -LiteralPath $p) {
        Set-Dates $p ([datetime]::ParseExact($m.CreationTime, 's', $culture)) ([datetime]::ParseExact($m.LastWriteTime, 's', $culture))
        $n++
    } else { $manquants++ }
}
if ($manquants) { Write-Warning "$manquants fichier(s) du manifeste absents : vérifiez que les deux archives (FileServer_Contoso_1.zip et _2.zip) sont dans $ZipPath." }
Remove-Item -LiteralPath $manifestPath -Force
Write-Host "$n fichiers horodatés." -ForegroundColor Cyan

# ---------------------------------------------------------------- Cas de test
$cas = New-Object System.Collections.Generic.List[object]
function Add-Cas([string]$Id, [string]$Chemin, [string]$Probleme, [string]$Attendu) {
    $cas.Add([pscustomobject]@{ Id = $Id; Chemin = $Chemin; Probleme = $Probleme; ComportementAttendu = $Attendu })
}

# 1. Chemin > 400 caractères (limite SharePoint Online sur l'URL décodée)
$segment = 'Projet_Refonte_Systeme_Information_Phase_2_Documentation_Technique_Detaillee_Version_Finale'
$long = Join-Path $FileServer '06_Archives\2018'
1..5 | ForEach-Object { $long = Join-Path $long ("{0:D2}_{1}" -f $_, $segment) }
$longFile = Join-Path $long 'Specification_fonctionnelle_detaillee_module_facturation.txt'
[System.IO.Directory]::CreateDirectory((Get-LongPath $long)) | Out-Null
[System.IO.File]::WriteAllText((Get-LongPath $longFile), 'Document fictif - chemin volontairement trop long.')
Add-Cas 'CT01' $longFile "Chemin de $($longFile.Length) caractères (> 400 dans SharePoint)" 'Échec de migration (chemin trop long)'

# 2. Noms interdits ou bloqués par SharePoint Online
$f = Join-Path $FileServer '01_Clients\CL1000_Boulangeries_Ravel\01_Contrats\~$ntrat_cadre_CL1000.docx'
[System.IO.File]::WriteAllText((Get-LongPath $f), 'lock')
Add-Cas 'CT02' $f 'Nom commençant par ~$ (fichier de verrouillage Office)' 'Non migré (nom non autorisé dans SharePoint)'

$f = Join-Path $FileServer '05_Commun\Procedures\_vti_config_intranet.txt'
[System.IO.File]::WriteAllText((Get-LongPath $f), 'config')
Add-Cas 'CT03' $f 'Nom contenant _vti_' 'Non migré (nom non autorisé dans SharePoint)'

# 3. Fichiers cachés / système
$f = Join-Path $FileServer '05_Commun\Evenements\Seminaire_2025_Annecy\Thumbs.db'
New-RandomFile $f 1
(Get-Item -LiteralPath $f -Force).Attributes = 'Hidden, System'
Add-Cas 'CT04' $f 'Fichier système caché Thumbs.db' 'Migré (Migrate hidden files activé par défaut dans Migration Manager)'

$f = Join-Path $FileServer '04_Direction\Strategie\Notes_personnelles_DG.txt'
[System.IO.File]::WriteAllText((Get-LongPath $f), 'Notes confidentielles fictives.')
(Get-Item -LiteralPath $f -Force).Attributes = 'Hidden'
Add-Cas 'CT05' $f 'Fichier caché' 'Migré par défaut ; ignoré si « Migrate hidden files » est désactivé'

# 4. Fichiers temporaires et extensions à exclure
$f = Join-Path $FileServer '03_Finance\Budgets\2026\~WRL0003.tmp'
New-RandomFile $f 1
Add-Cas 'CT06' $f 'Fichier temporaire Office (.tmp)' 'Filtré si TMP figure dans « Don''t migrate files with these extensions »'

$f = Join-Path $FileServer '06_Archives\Sauvegarde_serveur_2016.bak'
New-RandomFile $f 12
Set-Dates $f ([datetime]'2016-11-04') ([datetime]'2016-11-04')
Add-Cas 'CT07' $f 'Sauvegarde .bak (12 Mo)' 'Filtré si BAK figure dans les extensions exclues'

$f = Join-Path $FileServer '06_Archives\Archive_messagerie_direction.pst'
New-RandomFile $f 8
Set-Dates $f ([datetime]'2017-02-09') ([datetime]'2017-02-09')
Add-Cas 'CT08' $f 'Archive Outlook .pst (8 Mo)' 'Filtré si PST figure dans les extensions exclues'

# 5. Cas qui doivent réussir
$f = Join-Path $FileServer '05_Commun\Modeles\Fichier_vide.txt'
[System.IO.File]::WriteAllText((Get-LongPath $f), '')
Add-Cas 'CT09' $f 'Fichier de 0 octet' 'Migré'

$f = Join-Path $FileServer '05_Commun\Evenements\Seminaire_2025_Annecy\Video_teaser.mp4'
New-RandomFile $f 25
Set-Dates $f ([datetime]'2025-06-20') ([datetime]'2025-06-20')
Add-Cas 'CT10' $f 'Vidéo 25 Mo' 'Migré'

$f = Join-Path $FileServer '04_Direction\Strategie\Plan #1 - Transformation 100% cloud.docx'
Add-Cas 'CT11' $f 'Caractères spéciaux # % & – dans le nom' 'Migré (caractères acceptés par SharePoint Online)'

Add-Cas 'CT12' (Join-Path $FileServer '*\Modele_devis - Copie.docx') 'Doublons (même fichier à 3 emplacements)' 'Migré 3 fois (Migration Manager ne dédoublonne pas)'
Add-Cas 'CT13' (Join-Path $FileServer '06_Archives\2009..2018') 'Fichiers anciens (non modifiés depuis > 8 ans)' 'Migrés, sauf si « Migrate files modified after » est renseigné'

# 6. Cas supplémentaires pour l'audit de migrabilité (05_Audit_Source.ps1)

$f = Join-Path $FileServer '05_Commun\Procedures\Intranet_ancien.aspx'
[System.IO.File]::WriteAllText((Get-LongPath $f), '<%@ Page Language="C#" %><html><body>Ancien intranet fictif</body></html>')
Add-Cas 'CT15' $f 'Page .aspx (type bloqué sans script personnalisé)' 'Échec (type de fichier bloqué)'

$f = Join-Path $FileServer '04_Direction\Raccourci vers Budgets.lnk'
try {
    if (-not $isWin) { throw 'non Windows' }
    $sc = (New-Object -ComObject WScript.Shell).CreateShortcut($f)
    $sc.TargetPath = '\\srv-fichiers\finance\Budgets'
    $sc.Save()
} catch { [System.IO.File]::WriteAllText((Get-LongPath $f), 'Raccourci fictif vers \\srv-fichiers\finance\Budgets') }
Add-Cas 'CT16' $f 'Raccourci Windows vers \\srv-fichiers' 'Migré mais lien cassé'

Add-Cas 'CT17' (Join-Path $FileServer '05_Commun\Procedures\Note_liens_serveur.docx') 'Document Word avec liens vers \\srv-fichiers' 'Migré mais liens non convertis'

$d = Join-Path $FileServer '04_Direction'
$f = Join-Path $d ' Note avec espace initial.txt'
[System.IO.File]::WriteAllText((Get-LongPath $f), 'Nom commençant par un espace.')
Add-Cas 'CT18' $f 'Espace en début de nom' 'Échec (nom non autorisé)'

$f = Join-Path $FileServer '05_Commun\forms\Formulaire_conges.txt'
[System.IO.Directory]::CreateDirectory((Get-LongPath (Split-Path $f -Parent))) | Out-Null
[System.IO.File]::WriteAllText((Get-LongPath $f), 'Formulaire fictif.')
Add-Cas 'CT19' (Split-Path $f -Parent) 'Dossier « forms » à la racine de la bibliothèque Commun' 'Échec (nom réservé à la racine)'

$f = Join-Path $FileServer '03_Finance\Budget;version2.txt'
[System.IO.File]::WriteAllText((Get-LongPath $f), 'Budget fictif.')
Add-Cas 'CT20' $f 'Point-virgule dans le nom' 'Migré (enregistrement Office limité)'

$f = Join-Path $FileServer '05_Commun\Modeles\Installer_modeles.bat'
[System.IO.File]::WriteAllText((Get-LongPath $f), '@echo Script fictif - aucune action')
Add-Cas 'CT21' $f 'Script .bat' 'Migré (à valider : exécutable)'

$f = Join-Path $FileServer '06_Archives\2012\Archive_hors_ligne.pdf'
[System.IO.File]::WriteAllText((Get-LongPath $f), 'Contenu archivé fictif.')
if ($isWin) { try { [System.IO.File]::SetAttributes((Get-LongPath $f), [System.IO.FileAttributes]::Offline) } catch { Write-Warning "Attribut Offline non appliqué : $($_.Exception.Message)" } }
Add-Cas 'CT22' $f 'Fichier marqué hors ligne (HSM)' 'À rappeler avant migration'

function Get-DirAcl($di) { try { $di.GetAccessControl() } catch { [System.IO.FileSystemAclExtensions]::GetAccessControl($di) } }
function Set-DirAcl($di, $acl) { try { $di.SetAccessControl($acl) } catch { [System.IO.FileSystemAclExtensions]::SetAccessControl($di, $acl) } }
if ($isWin) {
    try {
        $inh = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
        $prop = [System.Security.AccessControl.PropagationFlags]::None

        $di = New-Object System.IO.DirectoryInfo (Join-Path $FileServer '02_RH_Collaborateurs\02_Recrutement')
        $acl = Get-DirAcl $di
        $invites = New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-32-546'
        $orphelin = New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-21-1111111111-2222222222-3333333333-1105'
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($invites, 'Modify', $inh, $prop, 'Deny')))
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($orphelin, 'ReadAndExecute', $inh, $prop, 'Allow')))
        Set-DirAcl $di $acl
        Add-Cas 'CT23' $di.FullName 'Refus (Deny) pour Invités + SID orphelin' 'Refus non migré ; SID sans correspondance'

        $di = New-Object System.IO.DirectoryInfo (Join-Path $FileServer '03_Finance\Comptabilite')
        $acl = Get-DirAcl $di
        $auth = New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-11'
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($auth, 'ReadAttributes, ReadExtendedAttributes', $inh, $prop, 'Allow')))
        Set-DirAcl $di $acl
        Add-Cas 'CT24' $di.FullName 'Droit NTFS avancé pour Utilisateurs authentifiés' 'Droit non converti'

        $di = New-Object System.IO.DirectoryInfo (Join-Path $FileServer '04_Direction\Comite_de_direction')
        $acl = Get-DirAcl $di
        $acl.SetAccessRuleProtection($true, $true)
        Set-DirAcl $di $acl
        Add-Cas 'CT25' $di.FullName 'Héritage des autorisations coupé' 'Permission unique SharePoint'
    } catch { Write-Warning "Cas de test d'autorisations non appliqués : $($_.Exception.Message)" }
}

if ($FichierVolumineuxMo -gt 0) {
    $f = Join-Path $FileServer "05_Commun\Evenements\Captation_seminaire_${FichierVolumineuxMo}Mo.mp4"
    Write-Host "Création d'un fichier de $FichierVolumineuxMo Mo ..." -ForegroundColor Cyan
    New-RandomFile $f $FichierVolumineuxMo
    Add-Cas 'CT14' $f "Fichier volumineux ($FichierVolumineuxMo Mo)" 'Migré (limite SharePoint : 250 Go)'
}

$cas | Export-Csv -LiteralPath (Join-Path $Racine 'Cas_de_test.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8

$unc = if ($SansPartage) { $null } else { New-PartageDemo $NomPartage $FileServer $ComptesLecture }

# ---------------------------------------------------------------- Bilan
$all = Get-ChildItem -LiteralPath $FileServer -Recurse -File -Force -ErrorAction SilentlyContinue
$taille = ($all | Measure-Object Length -Sum).Sum
Write-Host ''
Write-Host '=== Serveur de fichiers de démonstration prêt ===' -ForegroundColor Green
Write-Host ("Emplacement : {0}" -f $FileServer)
Write-Host ("Chemin UNC  : {0}" -f $(if ($unc) { $unc } else { '(partage à créer - Migration Manager exige un chemin \\serveur\partage)' }))
Write-Host ("Fichiers    : {0:N0}" -f $all.Count)
Write-Host ("Volume      : {0:N1} Mo" -f ($taille / 1MB))
Get-ChildItem -LiteralPath $FileServer -Directory | ForEach-Object {
    $files = Get-ChildItem -LiteralPath $_.FullName -Recurse -File -Force -ErrorAction SilentlyContinue
    Write-Host ("  {0,-24} {1,5} fichiers  {2,8:N1} Mo" -f $_.Name, $files.Count, (($files | Measure-Object Length -Sum).Sum / 1MB))
}
Write-Host ("Cas de test : {0}" -f (Join-Path $Racine 'Cas_de_test.csv'))
if ($Racine -ne 'C:\DemoMigration') {
    Write-Host ''
    Write-Host "Emplacement personnalisé : pour les scripts 02 et 05, indiquez la source avec -Source '$FileServer'" -ForegroundColor Yellow
}
