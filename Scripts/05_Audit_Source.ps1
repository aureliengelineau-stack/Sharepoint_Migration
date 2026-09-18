<#
.SYNOPSIS
    Audit de migrabilité d'un serveur de fichiers vers SharePoint Online avant migration avec Migration Manager.

.DESCRIPTION
    Parcourt la source (fichiers, dossiers, ACL NTFS) et vérifie les contraintes SharePoint Online / Migration Manager :
      - Nommage et chemins  : chemin cible > 400 caractères, caractères interdits, noms réservés, espaces, dossier « forms »…
      - Migration Manager    : source accessible en chemin UNC, espace du dossier de travail de l'agent
      - Taille et types      : > 250 Go, > 15 Go, types bloqués sans script personnalisé, .pst, aperçus > 100 Mo…
      - Accès aux fichiers   : fichiers illisibles, verrouillés, hors ligne (HSM), cachés/système
      - Liens et formats     : raccourcis, liens symboliques, liens \\serveur dans les documents Office, OneNote
      - Autorisations        : héritage coupé, refus (Deny), SID orphelins, comptes locaux, droits avancés
      - Volumétrie           : dossiers > 5 000 éléments, bibliothèques > 100 000 / 300 000, permissions uniques, quota
      - Nettoyage (ROT)      : fichiers anciens, vides, exécutables, doublons (option -AvecHash)

    Sorties (dossier -Sortie) :
      Audit_Source_Controles.csv     référentiel des contrôles (code, gravité, règle, action)
      Audit_Source_Constats.csv      un constat par élément et par contrôle
      Audit_Source_Synthese.csv      une ligne par bibliothèque cible, avec feu Vert / Orange / Rouge
      Audit_Source_Permissions.csv   entrées d'ACL explicites et leur correspondance SharePoint

    Au lancement, une fenêtre Windows demande le dossier de sortie (proposé : -Sortie).
    -SansFenetre utilise -Sortie directement (exécution planifiée ou automatisée).

    Compatible Windows PowerShell 5.1 et PowerShell 7. Lecture seule : aucun fichier source n'est modifié.

.EXAMPLE
    .\05_Audit_Source.ps1
.EXAMPLE
    .\05_Audit_Source.ps1 -Source '\\srv-fichiers\Partages' -SiteUrl 'https://contoso.sharepoint.com/sites/Migration' -AvecHash -QuotaDisponibleGo 1200
.EXAMPLE
    .\05_Audit_Source.ps1 -Sortie 'D:\Exports\Audit' -SansFenetre
#>
[CmdletBinding()]
param(
    [string]$FichierParametres = $(if ($PSScriptRoot) { Join-Path $PSScriptRoot 'Parametres_Client.psd1' } else { 'Parametres_Client.psd1' }),
    [hashtable]$ParametresClient = $(if (Test-Path -LiteralPath $FichierParametres) { Import-PowerShellDataFile -LiteralPath $FichierParametres } else { @{} }),
    [string]$Racine = $(if ($ParametresClient['CheminLocal']) { Split-Path -Parent $ParametresClient['CheminLocal'] } else { 'C:\DemoMigration' }),
    [string]$Source = $(if ($ParametresClient['CheminLocal']) { $ParametresClient['CheminLocal'] } else { (Join-Path $Racine 'FileServer') }),
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
    [string[]]$ExtensionsExclues = @('.bak'),
    [int]$AnneesObsolescence = 5,
    [double]$QuotaDisponibleGo = 0,
    [switch]$AvecHash,
    [switch]$PermissionsFichiers,
    [switch]$SansTestOuverture,
    [switch]$SansAnalyseLiens,
    [switch]$SansROT,
    [string]$CheminUNC = [string]$ParametresClient['CheminUNC'],
    [double]$EspaceAgentMinGo = 150,
    [switch]$SansControlePosteAgent,
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
$isWin = [System.Environment]::OSVersion.Platform -eq 'Win32NT'
$sep = [System.IO.Path]::DirectorySeparatorChar
function Get-LongPath([string]$Path) { if ($isWin -and -not $Path.StartsWith('\\?\')) { if ($Path.StartsWith('\\')) { '\\?\UNC\' + $Path.Substring(2) } else { '\\?\' + $Path } } else { $Path } }
function Remove-LongPrefix([string]$Path) { if ($Path.StartsWith('\\?\UNC\')) { '\\' + $Path.Substring(8) } elseif ($Path.StartsWith('\\?\')) { $Path.Substring(4) } else { $Path } }

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

if (-not (Test-Path -LiteralPath $Source)) { throw "Dossier source introuvable : $Source" }
$Sortie = Select-DossierSortie -Titre 'Choisissez le dossier de sortie de l''audit de migrabilité (Audit_Source_*.csv)' -DossierPropose $Sortie -SansFenetre:$SansFenetre
if (-not $Sortie) { Write-Warning 'Sélection annulée : aucun audit réalisé.'; return }
New-Item -ItemType Directory -Path $Sortie -Force | Out-Null
$Source = (Resolve-Path -LiteralPath $Source).ProviderPath.TrimEnd($sep)
$sortiePleine = [System.IO.Path]::GetFullPath($Sortie).TrimEnd($sep)
if (($sortiePleine + $sep).StartsWith($Source + $sep, [System.StringComparison]::OrdinalIgnoreCase)) {
    Write-Warning "Le dossier de sortie est situé dans la source : les CSV produits seront eux-mêmes inventoriés lors des prochains passages."
}
$sitePath = ([uri]$SiteUrl).AbsolutePath.TrimEnd('/')

# ------------------------------------------------------------------ référentiel des contrôles
$R = 'Bloquant'; $O = 'À traiter'; $I = 'Information'
$controles = @(
    @('CHEMIN_400', 'Nommage et chemins', $R, 'Chemin cible de plus de 400 caractères', 'Chemin décodé complet (site + bibliothèque + dossiers + nom) limité à 400 caractères.', 'Raccourcir les noms, réduire la profondeur ou cibler un niveau plus haut dans la bibliothèque.'),
    @('CHEMIN_350', 'Nommage et chemins', $O, 'Chemin cible entre 350 et 400 caractères', 'Marge de sécurité avant la limite de 400 caractères.', 'Surveiller : tout renommage ou déplacement ultérieur peut faire dépasser la limite.'),
    @('CARACTERES', 'Nommage et chemins', $O, 'Caractères interdits dans le nom', 'Interdits : " * : < > ? / \ | et caractères de contrôle.', 'Renommer avant migration ; sinon vérifier le nom obtenu (« Migrate files and folders with invalid characters », activé par défaut).'),
    @('ESPACES', 'Nommage et chemins', $R, 'Espace en début ou en fin de nom', 'Les espaces de début et de fin ne sont pas autorisés.', 'Renommer l''élément.'),
    @('NOM_RESERVE', 'Nommage et chemins', $R, 'Nom réservé ou interdit', '.lock, CON, PRN, AUX, NUL, COM0-9, LPT0-9, desktop.ini, noms commençant par ~$, _vti_ n''importe où.', 'Renommer ou exclure (fichiers de verrouillage Office : supprimer).'),
    @('FORMS_RACINE', 'Nommage et chemins', $R, 'Dossier « forms » à la racine d''une bibliothèque', 'Le nom forms est réservé à la racine d''une bibliothèque.', 'Renommer le dossier (ex. Formulaires).'),
    @('DEBUT_DOSSIER', 'Nommage et chemins', $R, 'Dossier commençant par ゛ ou ဧ', 'Ces caractères ne peuvent pas débuter un nom de dossier.', 'Renommer le dossier.'),
    @('DIESE_POURCENT', 'Nommage et chemins', $I, 'Caractère # ou % dans le nom', 'Pris en charge par SharePoint Online mais désactivable par l''organisation.', 'Vérifier le paramétrage du tenant.'),
    @('POINT_VIRGULE', 'Nommage et chemins', $I, 'Point-virgule dans le nom', 'Empêche l''enregistrement depuis le Backstage Office vers SharePoint.', 'Renommer si le fichier est édité dans Office.'),
    @('TAILLE_250GO', 'Taille et types', $R, 'Fichier de plus de 250 Go', 'Taille maximale d''un fichier dans SharePoint Online : 250 Go.', 'Conserver hors SharePoint (Azure Files, archivage).'),
    @('TAILLE_15GO', 'Taille et types', $I, 'Fichier entre 15 et 250 Go', 'Pris en charge par Migration Manager (250 Go par fichier) mais long à transférer et consommateur d''espace sur l''agent.', 'Regrouper dans une tâche dédiée et vérifier l''espace du dossier de travail de l''agent.'),
    @('TYPE_SCRIPT', 'Taille et types', $R, 'Type bloqué sans script personnalisé', '.asmx .ascx .aspx .htc .jar .master .swf .xap .xsf refusés sur les sites où le script est bloqué (défaut).', 'Exclure, archiver en .zip ou autoriser le script après analyse de sécurité.'),
    @('PST', 'Taille et types', $O, 'Archive Outlook (.pst)', 'Synchronisation limitée, blocable par stratégie ; SharePoint n''est pas la cible adaptée.', 'Importer dans les boîtes Exchange Online (service d''import Microsoft Purview).'),
    @('EXT_EXCLUE', 'Taille et types', $I, 'Extension exclue par la politique du projet', 'Extension listée dans le paramètre -ExtensionsExclues.', 'Exclure dans les paramètres de la tâche : « Don''t migrate files with these extensions » (format TMP:BAK:PST).'),
    @('APERCU_100MO', 'Taille et types', $I, 'Image ou PDF de plus de 100 Mo', 'Pas de miniature ni d''aperçu au-delà de 100 Mo.', 'Informer les utilisateurs ou optimiser les fichiers.'),
    @('INACCESSIBLE', 'Accès aux fichiers', $R, 'Élément illisible par le compte d''audit', 'Accès refusé : l''agent Migration Manager ne pourra pas le lire si son compte Windows a les mêmes droits.', 'Accorder la lecture au compte Windows de l''agent Migration Manager.'),
    @('VERROUILLE', 'Accès aux fichiers', $O, 'Fichier verrouillé (ouvert)', 'Le fichier est ouvert en écriture exclusive par un autre processus.', 'Migrer hors heures ouvrées ou fermer les sessions.'),
    @('HORS_LIGNE', 'Accès aux fichiers', $O, 'Fichier hors ligne ou archivé (HSM)', 'Attribut Offline ou Recall : le contenu n''est pas présent sur le disque.', 'Rappeler le contenu avant migration ou exclure.'),
    @('CACHE_SYSTEME', 'Accès aux fichiers', $I, 'Élément caché ou système', 'Migré par défaut (« Migrate hidden files » activé) ; l''attribut caché n''est pas conservé dans SharePoint.', 'Décider : inclure, exclure ou supprimer.'),
    @('RACCOURCI', 'Liens et formats', $O, 'Raccourci Windows (.lnk, .url)', 'Migré comme simple fichier : la cible sur le serveur ne sera plus accessible.', 'Remplacer par des liens SharePoint après migration.'),
    @('LIEN_SYMBOLIQUE', 'Liens et formats', $O, 'Lien symbolique ou jonction', 'Point d''analyse NTFS : le contenu cible n''est pas parcouru par l''audit.', 'Migrer la cible réelle et supprimer le lien.'),
    @('LIENS_UNC', 'Liens et formats', $O, 'Document Office avec liens vers le serveur', 'Liens externes vers \\serveur ou file:// (liaisons Excel, liens hypertexte) : non convertis.', 'Mettre à jour les liaisons après migration.'),
    @('ONENOTE', 'Liens et formats', $O, 'Bloc-notes OneNote', 'Bloc-notes limité à 2 Go ; les sections doivent être rouvertes et vérifiées après migration.', 'Tester l''ouverture après migration ou migrer via l''application OneNote.'),
    @('PERM_DENY', 'Autorisations', $O, 'Refus (Deny) explicite', 'Les refus ne sont pas migrés : l''accès peut s''ouvrir dans SharePoint.', 'Revoir la sécurité de l''élément avant migration.'),
    @('PERM_SID_ORPHELIN', 'Autorisations', $O, 'SID orphelin dans l''ACL', 'Compte supprimé : aucune correspondance possible dans Entra ID.', 'Nettoyer l''ACL.'),
    @('PERM_COMPTE_LOCAL', 'Autorisations', $O, 'Compte local ou groupe intégré', 'Pas d''équivalent dans Entra ID (comptes du serveur, BUILTIN, Tout le monde…).', 'Remplacer par des groupes synchronisés ou prévoir un fichier de correspondance.'),
    @('PERM_AVANCEE', 'Autorisations', $I, 'Droit NTFS avancé', 'Seuls Lecture, Écriture/Modification et Contrôle total sont convertis.', 'Vérifier le niveau d''accès obtenu dans SharePoint.'),
    @('PERM_UNIQUE', 'Autorisations', $I, 'Héritage coupé ou droits explicites', 'Deviendra une permission unique SharePoint (si comptes synchronisés).', 'Limiter les permissions uniques ; privilégier les droits au niveau bibliothèque.'),
    @('MM_CHEMIN_UNC', 'Migration Manager', $R, 'Source non accessible en chemin UNC', 'Migration Manager n''accepte que des chemins \\serveur\partage ; les chemins locaux ne sont pas pris en charge.', 'Créer un partage SMB en lecture pour le compte Windows de l''agent (script 01) ou indiquer -CheminUNC.'),
    @('MM_ESPACE_AGENT', 'Migration Manager', $O, "Moins de $EspaceAgentMinGo Go libres pour le dossier de travail de l'agent", 'Le dossier de travail (%appdata%\Microsoft\SPMigration par défaut) nécessite au moins 150 Go libres sur le poste de l''agent.', 'Libérer de l''espace ou déplacer le dossier de travail (Advanced settings) ; contrôle valable si ce poste héberge l''agent.'),
    @('DOSSIER_5000', 'Volumétrie', $O, 'Dossier de plus de 5 000 éléments directs', 'Seuil d''affichage des vues SharePoint : 5 000 éléments.', 'Scinder le dossier ou prévoir des vues indexées.'),
    @('PARTAGE_50000', 'Volumétrie', $I, 'Dossier de plus de 50 000 sous-éléments', 'Un dossier de plus de 50 000 éléments ne peut pas être partagé.', 'Partager à un niveau inférieur.'),
    @('BIBLIO_100K', 'Volumétrie', $O, 'Bibliothèque de plus de 100 000 éléments', 'Au-delà, l''héritage des autorisations de la bibliothèque ne peut plus être rompu.', 'Répartir sur plusieurs bibliothèques ou sites.'),
    @('BIBLIO_300K', 'Volumétrie', $O, 'Bibliothèque de plus de 300 000 éléments', 'Recommandation Microsoft pour la synchronisation OneDrive.', 'Découper, ou déconseiller la synchronisation.'),
    @('PERM_UNIQUES_5000', 'Volumétrie', $O, 'Plus de 5 000 permissions uniques', 'Recommandation : 5 000 permissions uniques par bibliothèque.', 'Simplifier le modèle de droits.'),
    @('PERM_UNIQUES_50000', 'Volumétrie', $R, 'Plus de 50 000 permissions uniques', 'Limite dure : 50 000 permissions uniques par bibliothèque.', 'Refondre le modèle de droits avant migration.'),
    @('QUOTA', 'Volumétrie', $R, 'Volume supérieur au stockage disponible', 'Volume source comparé au paramètre -QuotaDisponibleGo.', 'Trier les données ou acheter du stockage.'),
    @('ANCIEN', 'Nettoyage (ROT)', $I, "Non modifié depuis plus de $AnneesObsolescence ans", 'Donnée potentiellement obsolète (Redundant, Obsolete, Trivial).', 'Archiver ou supprimer après validation métier.'),
    @('FICHIER_VIDE', 'Nettoyage (ROT)', $I, 'Fichier vide (0 octet)', 'Migrable, mais sans contenu.', 'Supprimer si inutile.'),
    @('TEMPORAIRE', 'Nettoyage (ROT)', $O, 'Fichier temporaire', '.tmp, ~*.tmp, .temp : non synchronisés, sans valeur.', 'Supprimer ou exclure.'),
    @('EXECUTABLE', 'Nettoyage (ROT)', $I, 'Exécutable ou script', '.exe .msi .bat .cmd .ps1 .vbs .dll .com .scr', 'Valider le besoin ; souvent à ne pas migrer.'),
    @('DOUBLON', 'Nettoyage (ROT)', $I, 'Contenu en doublon', 'Même empreinte SHA-256 qu''un autre fichier (option -AvecHash).', 'Dédoublonner avant migration.')
)
$refControles = @{}
$ordre = 0
$tableControles = foreach ($c in $controles) {
    $ordre++
    $refControles[$c[0]] = @{ Categorie = $c[1]; Gravite = $c[2] }
    [pscustomobject][ordered]@{ Code = $c[0]; Ordre = $ordre; Categorie = $c[1]; Gravite = $c[2]; Libelle = $c[3]; Regle = $c[4]; Action = $c[5] }
}

# ------------------------------------------------------------------ outils
$constats = New-Object System.Collections.Generic.List[object]
$permissions = New-Object System.Collections.Generic.List[object]
$nomsReserves = @('.lock', 'CON', 'PRN', 'AUX', 'NUL', 'desktop.ini') + (0..9 | ForEach-Object { "COM$_"; "LPT$_" })
$typesScript = @('.asmx', '.ascx', '.aspx', '.htc', '.jar', '.master', '.swf', '.xap', '.xsf')
$typesExec = @('.exe', '.msi', '.bat', '.cmd', '.ps1', '.vbs', '.dll', '.com', '.scr')
$typesApercu = @('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff', '.heic', '.pdf')
$typesOffice = @('.docx', '.docm', '.xlsx', '.xlsm', '.pptx', '.pptm')
$reCaracteres = New-Object System.Text.RegularExpressions.Regex '["*:<>?|\x00-\x1F]'
$reLienExterne = New-Object System.Text.RegularExpressions.Regex 'Target="((?:file:|\\\\)[^"]+)"[^>]*TargetMode="External"|TargetMode="External"[^>]*Target="((?:file:|\\\\)[^"]+)"', 'IgnoreCase'
$dateObsolete = (Get-Date).AddYears(-$AnneesObsolescence)
$sha = if ($AvecHash) { [System.Security.Cryptography.SHA256]::Create() } else { $null }
$zipDisponible = $false
if (-not $SansAnalyseLiens) { try { Add-Type -AssemblyName System.IO.Compression; Add-Type -AssemblyName System.IO.Compression.FileSystem; $zipDisponible = $true } catch { } }
$ATTR_RECALL = 0x40000 -bor 0x400000
$SID_INTEGRES = '^S-1-(1-0|5-(7|11|18|19|20|32-\d+)|3-\d)$'

function Get-Bibliotheque([string]$relatif) {
    $premier = $relatif.Split($sep)[0]
    if ($Correspondance.Contains($premier)) { return @($premier, [string]$Correspondance[$premier]) }
    return @($premier, '(non mappé)')
}

function Add-Constat($code, $element, [string]$detail) {
    $ref = $refControles[$code]
    $constats.Add([pscustomobject][ordered]@{
        Code                = $code
        Gravite             = $ref.Gravite
        Categorie           = $ref.Categorie
        Bibliotheque        = $element.Bibliotheque
        DossierRacine       = $element.DossierRacine
        TypeElement         = $element.Type
        Chemin              = $element.Chemin
        CheminRelatif       = $element.Relatif
        Nom                 = $element.Nom
        Extension           = $element.Extension
        TailleOctets        = $element.Taille
        DateModification    = $element.DateModif
        LongueurCheminCible = $element.LongueurCible
        Detail              = $detail
    })
}

function Get-AclSafe($fsi) {
    if (-not $isWin) { return $null }
    try { return $fsi.GetAccessControl() } catch { }
    try { return [System.IO.FileSystemAclExtensions]::GetAccessControl($fsi) } catch { }
    return $null
}

function Convert-Droits([int64]$droits) {
    $FULL = 2032127; $MODIFY = 197055; $WRITE = 278; $READ = 131209
    if (($droits -band 0x10000000) -ne 0 -or ($droits -band $FULL) -eq $FULL) { return 'Contrôle total' }
    if (($droits -band 0x40000000) -ne 0 -or ($droits -band $MODIFY) -eq $MODIFY -or ($droits -band $WRITE) -eq $WRITE) { return 'Collaboration' }
    if (($droits -band 0x80000000) -ne 0 -or ($droits -band 0x20000000) -ne 0 -or ($droits -band $READ) -eq $READ) { return 'Lecture' }
    return $null
}

function Test-Permissions($fsi, $element, $stat) {
    $acl = Get-AclSafe $fsi
    if ($null -eq $acl) { return }
    $regles = @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]))
    $protege = $acl.AreAccessRulesProtected
    if ($regles.Count -eq 0 -and -not $protege) { return }
    $stat.PermissionsUniques++
    Add-Constat 'PERM_UNIQUE' $element ("{0} entrée(s) explicite(s){1}" -f $regles.Count, $(if ($protege) { ', héritage coupé' } else { '' }))
    $deny = 0; $orph = @(); $locaux = @(); $avancees = @()
    foreach ($r in $regles) {
        $sid = $r.IdentityReference.Value
        $nom = $sid; $orpheline = $false
        try { $nom = $r.IdentityReference.Translate([System.Security.Principal.NTAccount]).Value } catch { $orpheline = $true }
        $local = (-not $orpheline) -and ($nom -like "$env:COMPUTERNAME\*" -or $sid -match $SID_INTEGRES)
        $acces = [string]$r.AccessControlType
        $droits = [int64]$r.FileSystemRights
        $corresp = if ($acces -eq 'Deny') { 'Non migré (refus)' } else { Convert-Droits $droits }
        if ($null -eq $corresp) { $corresp = 'Droit avancé non converti'; $avancees += $nom }
        if ($acces -eq 'Deny') { $deny++ }
        if ($orpheline) { $orph += $sid }
        if ($local) { $locaux += $nom }
        $permissions.Add([pscustomobject][ordered]@{
            Bibliotheque = $element.Bibliotheque; Chemin = $element.Chemin; CheminRelatif = $element.Relatif; TypeElement = $element.Type
            HeritageCoupe = $protege; Identite = $nom; SID = $sid; TypeAcces = $acces; Droits = [string]$r.FileSystemRights
            CorrespondanceSharePoint = $corresp; Orpheline = $orpheline; CompteLocalOuIntegre = $local
        })
    }
    if ($deny) { $stat.Deny += $deny; Add-Constat 'PERM_DENY' $element "$deny refus explicite(s)" }
    if ($orph.Count) { $stat.Orphelins += $orph.Count; Add-Constat 'PERM_SID_ORPHELIN' $element ($orph -join ', ') }
    if ($locaux.Count) { Add-Constat 'PERM_COMPTE_LOCAL' $element (($locaux | Select-Object -Unique) -join ', ') }
    if ($avancees.Count) { Add-Constat 'PERM_AVANCEE' $element (($avancees | Select-Object -Unique) -join ', ') }
}

function Test-Nom($element, [bool]$estDossier, [bool]$racineBiblio) {
    $n = $element.Nom
    if ($reCaracteres.IsMatch($n)) { Add-Constat 'CARACTERES' $element ("Caractère(s) : " + (($reCaracteres.Matches($n) | ForEach-Object { $_.Value } | Select-Object -Unique) -join ' ')) }
    if ($n -ne $n.Trim()) { Add-Constat 'ESPACES' $element "Nom : '$n'" }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($n)
    if ($n.StartsWith('~$') -or $n.IndexOf('_vti_', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or ($nomsReserves -contains $n) -or ($nomsReserves -contains $base)) {
        Add-Constat 'NOM_RESERVE' $element "Nom : $n"
    }
    if ($estDossier -and $racineBiblio -and $n -eq 'forms') { Add-Constat 'FORMS_RACINE' $element 'Dossier forms à la racine de la bibliothèque' }
    if ($estDossier -and $n.Length -gt 0 -and ($n[0] -eq [char]0x309B -or $n[0] -eq [char]0x1027)) { Add-Constat 'DEBUT_DOSSIER' $element "Nom : $n" }
    if ($n.IndexOfAny([char[]]'#%') -ge 0) { Add-Constat 'DIESE_POURCENT' $element "Nom : $n" }
    if ($n.Contains(';')) { Add-Constat 'POINT_VIRGULE' $element "Nom : $n" }
    if ($element.LongueurCible -gt 400) { Add-Constat 'CHEMIN_400' $element ("{0} caractères" -f $element.LongueurCible) }
    elseif ($element.LongueurCible -gt 350) { Add-Constat 'CHEMIN_350' $element ("{0} caractères" -f $element.LongueurCible) }
}

# ------------------------------------------------------------------ parcours
Write-Host "Audit de $Source ..." -ForegroundColor Cyan
$stats = @{}
$enfantsDirects = @{}
$sousElements = @{}
$fichiersSous = @{}
$hashs = @{}
$pile = New-Object System.Collections.Stack
$pile.Push((New-Object System.IO.DirectoryInfo (Get-LongPath $Source)))
$nFichiers = 0

while ($pile.Count -gt 0) {
    $dir = $pile.Pop()
    $dirPlein = (Remove-LongPrefix $dir.FullName).TrimEnd($sep)
    try { $entrees = @($dir.EnumerateFileSystemInfos()) }
    catch {
        $rel = if ($dirPlein.Length -gt $Source.Length) { $dirPlein.Substring($Source.Length + 1) } else { '' }
        $b = Get-Bibliotheque $rel
        Add-Constat 'INACCESSIBLE' ([pscustomobject]@{ Bibliotheque = $b[1]; DossierRacine = $b[0]; Type = 'Dossier'; Chemin = $dirPlein; Relatif = $rel; Nom = $dir.Name; Extension = ''; Taille = $null; DateModif = ''; LongueurCible = $null }) $_.Exception.Message
        continue
    }
    $enfantsDirects[$dirPlein] = $entrees.Count

    foreach ($e in $entrees) {
        $plein = Remove-LongPrefix $e.FullName
        $relatif = $plein.Substring($Source.Length + 1)
        $b = Get-Bibliotheque $relatif
        if (-not $stats.ContainsKey($b[1])) {
            $stats[$b[1]] = @{ DossierRacine = $b[0]; Fichiers = 0; Dossiers = 0; Volume = [int64]0; PermissionsUniques = 0; Deny = 0; Orphelins = 0; MaxEnfants = 0; DossierMax = '' }
        }
        $stat = $stats[$b[1]]
        $estDossier = $e -is [System.IO.DirectoryInfo]
        $sous = if ($relatif.Contains([string]$sep)) { $relatif.Substring($b[0].Length + 1) } else { '' }
        $cible = if ($sous) { "$sitePath/$($b[1])/" + ($sous -replace [regex]::Escape([string]$sep), '/') } else { "$sitePath/$($b[1])" }
        $attr = [int64]$e.Attributes
        $element = [pscustomobject]@{
            Bibliotheque = $b[1]; DossierRacine = $b[0]; Type = $(if ($estDossier) { 'Dossier' } else { 'Fichier' })
            Chemin = $plein; Relatif = $relatif; Nom = $e.Name
            Extension = $(if ($estDossier) { '' } else { $e.Extension.ToLowerInvariant() })
            Taille = $(if ($estDossier) { $null } else { $e.Length })
            DateModif = $e.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss'); LongueurCible = $cible.Length
        }

        # incrémente les sous-éléments de tous les dossiers parents (limite de partage 50 000)
        $p = [System.IO.Path]::GetDirectoryName($plein)
        while ($p -and $p.Length -gt $Source.Length) {
            if ($sousElements.ContainsKey($p)) { $sousElements[$p]++ } else { $sousElements[$p] = 1 }
            if (-not $estDossier) { if ($fichiersSous.ContainsKey($p)) { $fichiersSous[$p]++ } else { $fichiersSous[$p] = 1 } }
            $p = [System.IO.Path]::GetDirectoryName($p)
        }

        $estRacineBiblio = $relatif -eq $b[0]
        if (-not $estRacineBiblio) { Test-Nom $element $estDossier ($estDossier -and ([System.IO.Path]::GetDirectoryName($relatif) -eq $b[0])) }
        if (($attr -band [int64][System.IO.FileAttributes]::Hidden) -or ($attr -band [int64][System.IO.FileAttributes]::System)) {
            Add-Constat 'CACHE_SYSTEME' $element ([string]$e.Attributes)
        }
        if ($attr -band [int64][System.IO.FileAttributes]::ReparsePoint) {
            Add-Constat 'LIEN_SYMBOLIQUE' $element ([string]$e.Attributes)
            if ($estDossier) { $stat.Dossiers++; continue }
        }

        if ($estDossier) {
            if (-not $estRacineBiblio) { $stat.Dossiers++ }
            Test-Permissions $e $element $stat
            $pile.Push($e)
            continue
        }

        # ---- fichier
        $nFichiers++
        if ($nFichiers % 2000 -eq 0) { Write-Host ("  {0:N0} fichiers analysés..." -f $nFichiers) }
        $stat.Fichiers++
        $stat.Volume += $e.Length
        $ext = $element.Extension
        $go = $e.Length / 1GB
        if ($go -gt 250) { Add-Constat 'TAILLE_250GO' $element ("{0:N1} Go" -f $go) }
        elseif ($go -gt 15) { Add-Constat 'TAILLE_15GO' $element ("{0:N1} Go" -f $go) }
        if ($typesScript -contains $ext) { Add-Constat 'TYPE_SCRIPT' $element "Extension $ext" }
        if ($ext -eq '.pst') { Add-Constat 'PST' $element ("{0:N0} Mo" -f ($e.Length / 1MB)) }
        if ($ExtensionsExclues -contains $ext) { Add-Constat 'EXT_EXCLUE' $element "Extension $ext" }
        if (($typesApercu -contains $ext) -and $e.Length -gt 100MB) { Add-Constat 'APERCU_100MO' $element ("{0:N0} Mo" -f ($e.Length / 1MB)) }
        if (($attr -band [int64][System.IO.FileAttributes]::Offline) -or ($attr -band $ATTR_RECALL)) { Add-Constat 'HORS_LIGNE' $element ([string]$e.Attributes) }
        if ($ext -eq '.lnk' -or $ext -eq '.url') { Add-Constat 'RACCOURCI' $element "Raccourci $ext" }
        if ($ext -eq '.one' -or $ext -eq '.onetoc2') { Add-Constat 'ONENOTE' $element ("{0:N1} Mo" -f ($e.Length / 1MB)) }
        if (-not $SansROT) {
            if ($e.Length -eq 0) { Add-Constat 'FICHIER_VIDE' $element '0 octet' }
            if ($ext -eq '.tmp' -or $ext -eq '.temp') { Add-Constat 'TEMPORAIRE' $element "Extension $ext" }
            if ($typesExec -contains $ext) { Add-Constat 'EXECUTABLE' $element "Extension $ext" }
            if ($e.LastWriteTime -lt $dateObsolete) { Add-Constat 'ANCIEN' $element ("Modifié le {0:dd/MM/yyyy}" -f $e.LastWriteTime) }
        }

        $horsLigne = ($attr -band [int64][System.IO.FileAttributes]::Offline) -or ($attr -band $ATTR_RECALL)
        if (-not $SansTestOuverture -and -not $horsLigne) {
            try {
                $fs = [System.IO.File]::Open($e.FullName, 'Open', 'Read', 'ReadWrite')
                try {
                    if ($AvecHash -and -not $SansROT -and $e.Length -gt 0) {
                        $h = [System.BitConverter]::ToString($sha.ComputeHash($fs)).Replace('-', '')
                        if (-not $hashs.ContainsKey($h)) { $hashs[$h] = New-Object System.Collections.Generic.List[object] }
                        $hashs[$h].Add($element)
                    }
                } finally { $fs.Dispose() }
            }
            catch [System.UnauthorizedAccessException] { Add-Constat 'INACCESSIBLE' $element 'Accès refusé en lecture' }
            catch [System.IO.IOException] { Add-Constat 'VERROUILLE' $element $_.Exception.Message }
            catch { Add-Constat 'INACCESSIBLE' $element $_.Exception.Message }
        }

        if ($zipDisponible -and ($typesOffice -contains $ext) -and $e.Length -gt 0 -and $e.Length -lt 50MB -and -not $horsLigne) {
            try {
                $zip = [System.IO.Compression.ZipFile]::OpenRead($e.FullName)
                try {
                    $liens = @()
                    foreach ($entree in $zip.Entries) {
                        if ($entree.FullName.EndsWith('.rels')) {
                            $sr = New-Object System.IO.StreamReader($entree.Open())
                            try { $xml = $sr.ReadToEnd() } finally { $sr.Dispose() }
                            foreach ($m in $reLienExterne.Matches($xml)) { $liens += ($m.Groups[1].Value + $m.Groups[2].Value) }
                        }
                    }
                    if ($liens.Count) { Add-Constat 'LIENS_UNC' $element ("{0} lien(s) : {1}" -f $liens.Count, (($liens | Select-Object -Unique -First 3) -join ' ; ')) }
                } finally { $zip.Dispose() }
            } catch { }
        }

        if ($PermissionsFichiers) { Test-Permissions $e $element $stat }
    }
}

# ------------------------------------------------------------------ contrôles globaux
foreach ($d in $enfantsDirects.Keys) {
    if ($d.Length -le $Source.Length) { continue }
    $rel = $d.Substring($Source.Length + 1)
    $b = Get-Bibliotheque $rel
    $el = [pscustomobject]@{ Bibliotheque = $b[1]; DossierRacine = $b[0]; Type = 'Dossier'; Chemin = $d; Relatif = $rel; Nom = [System.IO.Path]::GetFileName($d); Extension = ''; Taille = $null; DateModif = ''; LongueurCible = $null }
    $n = $enfantsDirects[$d]
    if ($stats.ContainsKey($b[1]) -and $n -gt $stats[$b[1]].MaxEnfants) { $stats[$b[1]].MaxEnfants = $n; $stats[$b[1]].DossierMax = $rel }
    if ($n -gt 5000) { Add-Constat 'DOSSIER_5000' $el ("{0:N0} éléments directs" -f $n) }
    if ($sousElements.ContainsKey($d) -and $sousElements[$d] -gt 50000) { Add-Constat 'PARTAGE_50000' $el ("{0:N0} sous-éléments" -f $sousElements[$d]) }
}
foreach ($h in $hashs.Keys) {
    $groupe = $hashs[$h]
    if ($groupe.Count -gt 1) {
        $tries = @($groupe | Sort-Object Chemin)
        $detail = "{0} copies, original : {1}" -f $tries.Count, $tries[0].Relatif
        for ($k = 1; $k -lt $tries.Count; $k++) { Add-Constat 'DOUBLON' $tries[$k] $detail }
    }
}
foreach ($bib in $stats.Keys) {
    $s = $stats[$bib]
    $el = [pscustomobject]@{ Bibliotheque = $bib; DossierRacine = $s.DossierRacine; Type = 'Bibliothèque'; Chemin = (Join-Path $Source $s.DossierRacine); Relatif = $s.DossierRacine; Nom = $bib; Extension = ''; Taille = $s.Volume; DateModif = ''; LongueurCible = $null }
    $items = $s.Fichiers + $s.Dossiers
    if ($items -gt 300000) { Add-Constat 'BIBLIO_300K' $el ("{0:N0} éléments" -f $items) }
    elseif ($items -gt 100000) { Add-Constat 'BIBLIO_100K' $el ("{0:N0} éléments" -f $items) }
    if ($s.PermissionsUniques -gt 50000) { Add-Constat 'PERM_UNIQUES_50000' $el ("{0:N0} permissions uniques" -f $s.PermissionsUniques) }
    elseif ($s.PermissionsUniques -gt 5000) { Add-Constat 'PERM_UNIQUES_5000' $el ("{0:N0} permissions uniques" -f $s.PermissionsUniques) }
}
$volumeTotal = ($stats.Values | ForEach-Object { $_.Volume } | Measure-Object -Sum).Sum

# ---- contrôles propres à Migration Manager
if (-not $CheminUNC) { $CheminUNC = Resolve-CheminUNC $Source }
if (-not $CheminUNC) {
    $detailUnc = if ($isWin) { "Aucun partage SMB ne contient $Source" } else { "Partage SMB non vérifiable sur ce système : indiquer -CheminUNC" }
    foreach ($bib in $stats.Keys) {
        $s = $stats[$bib]
        Add-Constat 'MM_CHEMIN_UNC' ([pscustomobject]@{ Bibliotheque = $bib; DossierRacine = $s.DossierRacine; Type = 'Bibliothèque'; Chemin = (Join-Path $Source $s.DossierRacine); Relatif = $s.DossierRacine; Nom = $bib; Extension = ''; Taille = $s.Volume; DateModif = ''; LongueurCible = $null }) $detailUnc
    }
}
if ($isWin -and -not $SansControlePosteAgent -and $env:APPDATA) {
    try {
        $lecteur = New-Object System.IO.DriveInfo ([System.IO.Path]::GetPathRoot($env:APPDATA))
        $libreGo = $lecteur.AvailableFreeSpace / 1GB
        if ($libreGo -lt $EspaceAgentMinGo) {
            Add-Constat 'MM_ESPACE_AGENT' ([pscustomobject]@{ Bibliotheque = '(toutes)'; DossierRacine = ''; Type = 'Poste'; Chemin = $lecteur.Name; Relatif = ''; Nom = [System.Environment]::MachineName; Extension = ''; Taille = $null; DateModif = ''; LongueurCible = $null }) ("{0:N1} Go libres sur {1}" -f $libreGo, $lecteur.Name)
        }
    } catch { }
}
if ($QuotaDisponibleGo -gt 0 -and ($volumeTotal / 1GB) -gt $QuotaDisponibleGo) {
    Add-Constat 'QUOTA' ([pscustomobject]@{ Bibliotheque = '(toutes)'; DossierRacine = ''; Type = 'Source'; Chemin = $Source; Relatif = ''; Nom = 'Volume total'; Extension = ''; Taille = $volumeTotal; DateModif = ''; LongueurCible = $null }) ("{0:N1} Go pour {1:N1} Go disponibles" -f ($volumeTotal / 1GB), $QuotaDisponibleGo)
}

# ------------------------------------------------------------------ synthèse
$dateAudit = $debut.ToString('yyyy-MM-ddTHH:mm:ss')
$constatsParBiblio = @{}
foreach ($grp in ($constats | Group-Object Bibliotheque)) { $constatsParBiblio[$grp.Name] = $grp.Group }
$synthese = foreach ($bib in ($stats.Keys | Sort-Object { $stats[$_].DossierRacine })) {
    $s = $stats[$bib]
    $c = if ($constatsParBiblio.ContainsKey($bib)) { $constatsParBiblio[$bib] } else { @() }
    $bloq = @($c | Where-Object { $_.Gravite -eq $R })
    # fichiers bloqués = fichiers avec un constat bloquant + fichiers situés sous un dossier bloquant
    $dossiersBloques = @($bloq | Where-Object { $_.TypeElement -eq 'Dossier' } | Select-Object -ExpandProperty Chemin -Unique | Sort-Object)
    $hauts = New-Object System.Collections.Generic.List[string]
    foreach ($d in $dossiersBloques) { if (-not ($hauts | Where-Object { $d.StartsWith($_ + $sep) })) { $hauts.Add($d) } }
    $fichiersBloques = 0
    foreach ($d in $hauts) { if ($fichiersSous.ContainsKey($d)) { $fichiersBloques += $fichiersSous[$d] } }
    $fichiersBloques += @($bloq | Where-Object { $_.TypeElement -eq 'Fichier' } | Where-Object { $f = $_.Chemin; -not ($hauts | Where-Object { $f.StartsWith($_ + $sep) }) } | Select-Object -ExpandProperty Chemin -Unique).Count
    $migrables = [Math]::Max(0, $s.Fichiers - $fichiersBloques)
    $nbR = $bloq.Count; $nbO = @($c | Where-Object { $_.Gravite -eq $O }).Count; $nbI = @($c | Where-Object { $_.Gravite -eq $I }).Count
    [pscustomobject][ordered]@{
        Bibliotheque = $bib; DossierRacine = $s.DossierRacine; Fichiers = $s.Fichiers; Dossiers = $s.Dossiers
        Elements = $s.Fichiers + $s.Dossiers; VolumeOctets = $s.Volume; MaxElementsDossier = $s.MaxEnfants; DossierLePlusCharge = $s.DossierMax
        PermissionsUniques = $s.PermissionsUniques; ACEDeny = $s.Deny; SIDOrphelins = $s.Orphelins
        ConstatsBloquants = $nbR; ConstatsATraiter = $nbO; ConstatsInformation = $nbI
        FichiersBloques = $fichiersBloques; FichiersMigrables = $migrables
        PctMigrable = $(if ($s.Fichiers) { [Math]::Round($migrables / $s.Fichiers, 4) } else { 1 })
        Feu = $(if ($nbR) { 'Rouge' } elseif ($nbO) { 'Orange' } else { 'Vert' })
        DateAudit = $dateAudit
    }
}

$tableControles | Export-Csv -LiteralPath (Join-Path $Sortie 'Audit_Source_Controles.csv') -NoTypeInformation -Encoding UTF8
$constats | Export-Csv -LiteralPath (Join-Path $Sortie 'Audit_Source_Constats.csv') -NoTypeInformation -Encoding UTF8
@($synthese) | Export-Csv -LiteralPath (Join-Path $Sortie 'Audit_Source_Synthese.csv') -NoTypeInformation -Encoding UTF8
if ($permissions.Count) {
    $permissions | Export-Csv -LiteralPath (Join-Path $Sortie 'Audit_Source_Permissions.csv') -NoTypeInformation -Encoding UTF8
} else {
    'Bibliotheque,Chemin,CheminRelatif,TypeElement,HeritageCoupe,Identite,SID,TypeAcces,Droits,CorrespondanceSharePoint,Orpheline,CompteLocalOuIntegre' |
        Set-Content -LiteralPath (Join-Path $Sortie 'Audit_Source_Permissions.csv') -Encoding UTF8
}

# ------------------------------------------------------------------ bilan console
Write-Host ''
Write-Host '=== Audit de migrabilité terminé ===' -ForegroundColor Green
Write-Host ("Chemin UNC pour Migration Manager : {0}" -f $(if ($CheminUNC) { $CheminUNC } else { 'NON TROUVÉ (bloquant)' }))
Write-Host ("Source : {0}   Fichiers : {1:N0}   Volume : {2:N1} Go   Durée : {3:N0} s" -f $Source, $nFichiers, ($volumeTotal / 1GB), ((Get-Date) - $debut).TotalSeconds)
foreach ($l in $synthese) {
    $couleur = switch ($l.Feu) { 'Rouge' { 'Red' } 'Orange' { 'Yellow' } default { 'Green' } }
    Write-Host ("  [{0,-6}] {1,-12} {2,6:N0} fichiers  {3,5:P1} migrables  bloquants {4,4}  à traiter {5,4}  infos {6,5}" -f $l.Feu, $l.Bibliotheque, $l.Fichiers, $l.PctMigrable, $l.ConstatsBloquants, $l.ConstatsATraiter, $l.ConstatsInformation) -ForegroundColor $couleur
}
$constats | Group-Object Code | Sort-Object Count -Descending | Select-Object -First 12 | ForEach-Object {
    Write-Host ("    {0,-20} {1,5}  {2}" -f $_.Name, $_.Count, $refControles[$_.Name].Gravite)
}
Write-Host "Sorties : $Sortie (Audit_Source_*.csv)"
