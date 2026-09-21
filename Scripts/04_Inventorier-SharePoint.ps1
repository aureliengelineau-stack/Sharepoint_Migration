<#
.SYNOPSIS
    Inventaire des fichiers réellement présents dans les bibliothèques SharePoint Online après la migration
    avec Migration Manager, via Microsoft Graph. Sert au rapprochement source <-> cible dans Power BI.

.DESCRIPTION
    Nécessite le module Microsoft.Graph.Authentication (Install-Module Microsoft.Graph.Authentication -Scope CurrentUser).
    Connexion déléguée avec l'étendue Sites.Read.All (consentement administrateur demandé la première fois).
    Sortie : <Sortie>\Inventaire_SharePoint.csv

    Au lancement, une fenêtre Windows demande le dossier de sortie (proposé : -Sortie). Le fichier
    Correspondance_Bibliotheques.csv est d'abord cherché dans ce dossier, puis dans <Racine>\Rapports.
    -SansFenetre utilise -Sortie directement.

.EXAMPLE
    .\04_Inventorier-SharePoint.ps1
.EXAMPLE
    .\04_Inventorier-SharePoint.ps1 -Sortie 'D:\Exports\Migration' -SansFenetre
#>
[CmdletBinding()]
param(
    [string]$FichierParametres = $(if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Parametres_Client.psd1' } else { 'Parametres_Client.psd1' }),
    [hashtable]$ParametresClient = $(if (Test-Path -LiteralPath $FichierParametres) { Import-PowerShellDataFile -LiteralPath $FichierParametres } else { @{} }),
    [string]$Racine = $(if ($ParametresClient['CheminLocal']) { Split-Path -Parent $ParametresClient['CheminLocal'] } else { 'C:\DemoMigration' }),
    [string]$SiteUrl = $(if ($ParametresClient['SiteUrl']) { $ParametresClient['SiteUrl'] } else { 'https://m365x71797824.sharepoint.com/sites/MigrationFileServer' }),
    [string]$Sortie = $(if ($ParametresClient['DossierRapports']) { $ParametresClient['DossierRapports'] } else { (Join-Path $Racine 'Rapports') }),
    [string]$FichierCorrespondance = (Join-Path $Sortie 'Correspondance_Bibliotheques.csv'),
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


if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Module Microsoft.Graph.Authentication absent. Installez-le : Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}
Import-Module Microsoft.Graph.Authentication

$Sortie = Select-DossierSortie -Titre 'Choisissez le dossier de sortie de l''inventaire SharePoint (Inventaire_SharePoint.csv)' -DossierPropose $Sortie -SansFenetre:$SansFenetre
if (-not $Sortie) { Write-Warning 'Sélection annulée : aucun fichier produit.'; return }
if (-not $PSBoundParameters.ContainsKey('FichierCorrespondance')) {
    $dansSortie = Join-Path $Sortie 'Correspondance_Bibliotheques.csv'
    if (Test-Path -LiteralPath $dansSortie) { $FichierCorrespondance = $dansSortie }
}

$ctx = Get-MgContext
if (-not $ctx -or $ctx.Scopes -notcontains 'Sites.Read.All') {
    Connect-MgGraph -Scopes 'Sites.Read.All' -NoWelcome
}

function Invoke-Graph([string]$Uri) {
    $items = @()
    while ($Uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        if ($resp.PSObject.Properties.Name -contains 'value') { $items += $resp.value } else { return $resp }
        $Uri = $resp.'@odata.nextLink'
    }
    return $items
}

$u = [uri]$SiteUrl
$site = Invoke-Graph ("https://graph.microsoft.com/v1.0/sites/{0}:{1}" -f $u.Host, $u.AbsolutePath.TrimEnd('/'))
Write-Host "Site : $($site.displayName) ($($site.id))" -ForegroundColor Cyan

$biblios = if (Test-Path -LiteralPath $FichierCorrespondance) {
    (Import-Csv -LiteralPath $FichierCorrespondance -Encoding UTF8).Bibliotheque
} else { $null }

$drives = Invoke-Graph "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$select=id,name,webUrl,driveType"
if ($biblios) { $drives = $drives | Where-Object { $biblios -contains $_.name } }
if (-not $drives) { throw "Aucune bibliothèque trouvée sur $SiteUrl" }

$select = 'id,name,size,file,folder,webUrl,createdDateTime,lastModifiedDateTime,fileSystemInfo,createdBy,lastModifiedBy'
$lignes = New-Object System.Collections.Generic.List[object]

foreach ($d in $drives) {
    Write-Host "  Bibliothèque $($d.name) ..." -NoNewline
    $pile = New-Object System.Collections.Stack
    $pile.Push(@{ Id = 'root'; Chemin = '' })
    $n = 0
    while ($pile.Count -gt 0) {
        $cur = $pile.Pop()
        $enfants = Invoke-Graph "https://graph.microsoft.com/v1.0/drives/$($d.id)/items/$($cur.Id)/children?`$top=999&`$select=$select"
        foreach ($it in $enfants) {
            $chemin = if ($cur.Chemin) { "$($cur.Chemin)/$($it.name)" } else { $it.name }
            if ($it.folder) {
                $pile.Push(@{ Id = $it.id; Chemin = $chemin })
            } elseif ($it.file) {
                $n++
                $ext = [System.IO.Path]::GetExtension($it.name).ToLowerInvariant()
                $lignes.Add([pscustomobject][ordered]@{
                    Bibliotheque        = $d.name
                    CheminRelatif       = $chemin
                    CheminCible         = "$($u.AbsolutePath.TrimEnd('/'))/$($d.name)/$chemin"
                    Nom                 = $it.name
                    Extension           = $ext
                    TailleOctets        = $it.size
                    DateCreationSP      = ([datetime]$it.createdDateTime).ToString('yyyy-MM-ddTHH:mm:ss')
                    DateModificationSP  = ([datetime]$it.lastModifiedDateTime).ToString('yyyy-MM-ddTHH:mm:ss')
                    DateModificationFichier = if ($it.fileSystemInfo) { ([datetime]$it.fileSystemInfo.lastModifiedDateTime).ToString('yyyy-MM-ddTHH:mm:ss') } else { '' }
                    CreePar             = $it.createdBy.user.displayName
                    ModifiePar          = $it.lastModifiedBy.user.displayName
                    WebUrl              = $it.webUrl
                    IdElement           = $it.id
                })
            }
        }
    }
    Write-Host " $n fichiers"
}

New-Item -ItemType Directory -Path $Sortie -Force | Out-Null
$lignes | Export-Csv -LiteralPath (Join-Path $Sortie 'Inventaire_SharePoint.csv') -NoTypeInformation -Encoding UTF8
Write-Host ''
Write-Host ("=== Inventaire SharePoint : {0:N0} fichiers, {1:N1} Mo ===" -f $lignes.Count, (($lignes | Measure-Object TailleOctets -Sum).Sum / 1MB)) -ForegroundColor Green
Write-Host "Sortie : $(Join-Path $Sortie 'Inventaire_SharePoint.csv')"
