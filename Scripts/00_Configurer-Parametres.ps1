<#
.SYNOPSIS
    Renseigne les paramètres de la mission (client, centre d'administration, site cible, chemins,
    dossier des rapports) dans Scripts\Parametres_Client.psd1, lu par les scripts 01 à 05.

.DESCRIPTION
    Sans paramètre, le script pose les 5 questions une par une : la valeur actuelle est proposée
    entre crochets, Entrée la conserve. Avec des paramètres, il écrit directement les valeurs
    fournies et laisse les autres inchangées (utile en exécution automatisée).

    Champs écrits dans Parametres_Client.psd1 :
      Client                Client / environnement (en-tête des exécutions)
      CentreAdministration  URL du centre d'administration SharePoint
      SiteUrl               Site cible (scripts 02, 04, 05)
      CheminLocal           Chemin local de la source, vu de ce poste (scripts 02, 05 ; racine du script 01)
      CheminUNC             Chemin UNC de la source, vu par l'agent (vide = détection du partage SMB)
      DossierRapports       Dossier des rapports (scripts 02 à 05, paramètre Power BI DossierRapports)

    Les scripts 01 à 05 relisent ce fichier à chaque lancement. Un paramètre passé en ligne de
    commande reste prioritaire sur le fichier.

    Compatible Windows PowerShell 5.1 et PowerShell 7. Aucun accès réseau, aucune connexion.

.EXAMPLE
    .\00_Configurer-Parametres.ps1
.EXAMPLE
    .\00_Configurer-Parametres.ps1 -Afficher
.EXAMPLE
    .\00_Configurer-Parametres.ps1 -Client 'Fabrikam · production' -CentreAdministration 'https://fabrikam-admin.sharepoint.com' `
        -SiteUrl 'https://fabrikam.sharepoint.com/sites/Migration' -CheminLocal 'D:\Partages\Commun' `
        -CheminUNC '\\SRV-FICHIERS\Commun' -DossierRapports 'D:\Migration\Rapports'
.EXAMPLE
    .\00_Configurer-Parametres.ps1 -Demo        # remet les valeurs de la démonstration Contoso
#>
[CmdletBinding()]
param(
    [string]$Client,
    [string]$CentreAdministration,
    [string]$SiteUrl,
    [string]$CheminLocal,
    [string]$CheminUNC,
    [string]$DossierRapports,
    [string]$Fichier = $(if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Parametres_Client.psd1' } else { 'Parametres_Client.psd1' }),
    [switch]$Demo,
    [switch]$Afficher,
    [switch]$SansQuestions
)

$ErrorActionPreference = 'Stop'

$Defauts = [ordered]@{
    Client               = 'Contoso · tenant de démonstration M365x71797824'
    CentreAdministration = 'https://m365x71797824-admin.sharepoint.com'
    SiteUrl              = 'https://m365x71797824.sharepoint.com/sites/MigrationFileServer'
    CheminLocal          = 'C:\DemoMigration\FileServer'
    CheminUNC            = ''
    DossierRapports      = 'C:\DemoMigration\Rapports'
}

$Questions = [ordered]@{
    Client               = 'Client / environnement'
    CentreAdministration = "URL du centre d'administration SharePoint"
    SiteUrl              = 'Site cible (URL complète)'
    CheminLocal          = 'Chemin local de la source (vu de ce poste)'
    CheminUNC            = "Chemin UNC de la source (vu par l'agent ; vide = détection automatique)"
    DossierRapports      = 'Dossier des rapports'
}

# ------------------------------------------------------------------ lecture de l'existant
$valeurs = [ordered]@{}
foreach ($cle in $Defauts.Keys) { $valeurs[$cle] = $Defauts[$cle] }

if (-not $Demo -and (Test-Path -LiteralPath $Fichier)) {
    try {
        $actuel = Import-PowerShellDataFile -LiteralPath $Fichier
        foreach ($cle in @($valeurs.Keys)) {
            if ($actuel.ContainsKey($cle) -and $null -ne $actuel[$cle]) { $valeurs[$cle] = [string]$actuel[$cle] }
        }
    } catch {
        Write-Warning ("Fichier illisible, valeurs par défaut utilisées : {0}" -f $_.Exception.Message)
    }
}

function Show-Parametres {
    param([System.Collections.IDictionary]$Valeurs, [string]$Chemin)
    Write-Host ''
    Write-Host 'Paramètres de la mission' -ForegroundColor Cyan
    Write-Host ('  Fichier : {0}' -f $Chemin) -ForegroundColor DarkGray
    foreach ($cle in $Valeurs.Keys) {
        $v = if ([string]::IsNullOrWhiteSpace([string]$Valeurs[$cle])) { '(vide)' } else { $Valeurs[$cle] }
        Write-Host ('  {0,-22}{1}' -f $cle, $v)
    }
    Write-Host ''
    Write-Host 'Utilisation par les scripts :' -ForegroundColor Cyan
    Write-Host '  01  racine = dossier parent de CheminLocal, nom du partage = dernier segment de CheminUNC'
    Write-Host '  02  Source = CheminLocal, CheminUNC, SiteUrl, Sortie = DossierRapports'
    Write-Host '  03  Sortie = DossierRapports'
    Write-Host '  04  SiteUrl, Sortie = DossierRapports'
    Write-Host '  05  Source = CheminLocal, CheminUNC, SiteUrl, Sortie = DossierRapports'
    Write-Host '  Power BI : paramètre DossierRapports = ' -NoNewline
    Write-Host $Valeurs['DossierRapports']
}

if ($Afficher) { Show-Parametres -Valeurs $valeurs -Chemin $Fichier; return }

# ------------------------------------------------------------------ valeurs fournies en paramètre
$fournis = @{
    Client               = $PSBoundParameters.ContainsKey('Client')
    CentreAdministration = $PSBoundParameters.ContainsKey('CentreAdministration')
    SiteUrl              = $PSBoundParameters.ContainsKey('SiteUrl')
    CheminLocal          = $PSBoundParameters.ContainsKey('CheminLocal')
    CheminUNC            = $PSBoundParameters.ContainsKey('CheminUNC')
    DossierRapports      = $PSBoundParameters.ContainsKey('DossierRapports')
}
foreach ($cle in @($valeurs.Keys)) {
    if ($fournis[$cle]) { $valeurs[$cle] = (Get-Variable -Name $cle -ValueOnly) }
}

$aDesParametres = ($fournis.Values -contains $true)

# ------------------------------------------------------------------ questions
$interactif = -not $SansQuestions -and -not $Demo -and -not $aDesParametres
if ($interactif -and -not [System.Environment]::UserInteractive) {
    Write-Warning 'Session non interactive : les valeurs actuelles sont conservées.'
    $interactif = $false
}

if ($interactif) {
    Write-Host ''
    Write-Host 'Paramètres de la mission - Entrée conserve la valeur entre crochets.' -ForegroundColor Cyan
    Write-Host ('Fichier écrit : {0}' -f $Fichier) -ForegroundColor DarkGray
    Write-Host ''
    foreach ($cle in @($valeurs.Keys)) {
        $actuelle = [string]$valeurs[$cle]
        $affiche = if ([string]::IsNullOrWhiteSpace($actuelle)) { 'vide' } else { $actuelle }
        $reponse = Read-Host ("{0}`n  [{1}]" -f $Questions[$cle], $affiche)
        if (-not [string]::IsNullOrWhiteSpace($reponse)) { $valeurs[$cle] = $reponse.Trim() }
        elseif ($reponse -eq '-') { $valeurs[$cle] = '' }
    }
}

# ------------------------------------------------------------------ contrôles simples
$avertissements = New-Object System.Collections.Generic.List[string]
if ($valeurs['SiteUrl'] -and $valeurs['SiteUrl'] -notmatch '^https://') {
    $avertissements.Add("Site cible : l'URL devrait commencer par https://")
}
if ($valeurs['CentreAdministration'] -and $valeurs['CentreAdministration'] -notmatch '^https://') {
    $avertissements.Add("Centre d'administration : l'URL devrait commencer par https://")
}
if ($valeurs['CheminUNC'] -and $valeurs['CheminUNC'] -notmatch '^\\\\[^\\]+\\[^\\]+') {
    $avertissements.Add('Chemin UNC : format attendu \\serveur\partage (Migration Manager refuse les chemins locaux)')
}
if ($valeurs['CheminLocal'] -and $valeurs['CheminLocal'] -match '^\\\\') {
    $avertissements.Add('Chemin local : un chemin UNC a été saisi ; renseignez plutôt le champ CheminUNC')
}
foreach ($cle in @('CheminLocal', 'DossierRapports')) {
    $chemin = [string]$valeurs[$cle]
    if ($chemin -and -not (Test-Path -LiteralPath $chemin)) {
        $avertissements.Add(("{0} : le dossier {1} n'existe pas encore sur ce poste" -f $cle, $chemin))
    }
}

# ------------------------------------------------------------------ écriture
function ConvertTo-LitteralPowerShell {
    param([string]$Valeur)
    "'" + ($Valeur -replace "'", "''") + "'"
}

$entete = @"
#
#  Paramètres de la mission - kit Devoteam « Migration Manager »
#  --------------------------------------------------------------
#  Lu automatiquement par les scripts 01 à 05 : ces valeurs deviennent leurs valeurs par défaut.
#  Un paramètre passé en ligne de commande reste prioritaire sur ce fichier.
#  Écrit par 00_Configurer-Parametres.ps1 le $(Get-Date -Format 'dd/MM/yyyy HH:mm').
#
@{

    # Client / environnement : en-tête des exécutions.
    Client = $(ConvertTo-LitteralPowerShell $valeurs['Client'])

    # URL du centre d'administration SharePoint (affichage et rappels).
    CentreAdministration = $(ConvertTo-LitteralPowerShell $valeurs['CentreAdministration'])

    # Site cible : site SharePoint qui reçoit les bibliothèques (scripts 02, 04, 05).
    SiteUrl = $(ConvertTo-LitteralPowerShell $valeurs['SiteUrl'])

    # Chemin local de la source, vu de ce poste (scripts 02 et 05 ; son parent sert de racine au script 01).
    CheminLocal = $(ConvertTo-LitteralPowerShell $valeurs['CheminLocal'])

    # Chemin UNC de la source, vu par l'agent. Vide = détection du partage SMB contenant CheminLocal.
    CheminUNC = $(ConvertTo-LitteralPowerShell $valeurs['CheminUNC'])

    # Dossier des rapports (scripts 02 à 05, paramètre Power BI DossierRapports).
    DossierRapports = $(ConvertTo-LitteralPowerShell $valeurs['DossierRapports'])

}
"@

$dossier = Split-Path -Parent $Fichier
if ($dossier -and -not (Test-Path -LiteralPath $dossier)) { New-Item -ItemType Directory -Path $dossier -Force | Out-Null }

if (Test-Path -LiteralPath $Fichier) {
    $sauvegarde = "$Fichier.bak"
    Copy-Item -LiteralPath $Fichier -Destination $sauvegarde -Force
    Write-Host ("Ancien fichier sauvegardé : {0}" -f (Split-Path -Leaf $sauvegarde)) -ForegroundColor DarkGray
}

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($Fichier, $entete, $utf8Bom)

# relecture de contrôle
$relu = Import-PowerShellDataFile -LiteralPath $Fichier
foreach ($cle in @($valeurs.Keys)) {
    if ([string]$relu[$cle] -ne [string]$valeurs[$cle]) {
        throw ("Écriture incorrecte pour {0} : relu '{1}' au lieu de '{2}'" -f $cle, $relu[$cle], $valeurs[$cle])
    }
}

Write-Host ''
Write-Host ("Paramètres enregistrés : {0}" -f $Fichier) -ForegroundColor Green
Show-Parametres -Valeurs $valeurs -Chemin $Fichier

if ($avertissements.Count -gt 0) {
    Write-Host ''
    Write-Host 'À vérifier :' -ForegroundColor Yellow
    foreach ($a in $avertissements) { Write-Host ("  - {0}" -f $a) -ForegroundColor Yellow }
}

Write-Host ''
Write-Host 'Étape suivante : .\01_Preparer-ServeurFichiers.ps1 (démo) ou .\05_Audit_Source.ps1 -AvecHash (serveur existant)' -ForegroundColor DarkGray
