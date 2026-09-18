KIT DE DÉMONSTRATION - MIGRATION SERVEUR DE FICHIERS -> SHAREPOINT ONLINE (MIGRATION MANAGER) + POWER BI
Tenant cible : m365x71797824 (Contoso) - site : /sites/MigrationFileServer
Toutes les données sont fictives.

Le kit est livré en 2 archives, à extraire TOUTES LES DEUX dans le même dossier : C:\DemoMigration
  Kit_Migration_Manager.zip            scripts, Power BI, rapports d'exemple, données (partie 1)
  Kit_Migration_Manager_Donnees_2.zip  données (partie 2) : ajoute Kit_Migration_Manager\Donnees\FileServer_Contoso_2.zip
Renommer ensuite le dossier Kit_Migration_Manager en Kit (C:\DemoMigration\Kit).
(avant extraction : clic droit sur chaque .zip > Propriétés > cocher « Débloquer »)

Contenu
  Scripts\00_Configurer-Parametres.ps1                paramètres de la mission (client, centre d'administration, site cible,
                                                     chemin local/UNC, dossier des rapports) -> Scripts\Parametres_Client.psd1
  Scripts\Parametres_Client.psd1                      fichier de paramètres lu par les scripts 01 à 05
  Scripts\01_Preparer-ServeurFichiers.ps1            crée C:\DemoMigration\FileServer (+ 25 cas de test) et le partage \\<PC>\FileServer
  Scripts\05_Audit_Source.ps1                        audit de migrabilité avant migration (42 contrôles, feux par bibliothèque)
  Scripts\02_Inventorier-Source.ps1                  inventaire source + MigrationManager_Taches.csv (tâches en masse, 6 colonnes)
  Scripts\03_Collecter-RapportsMigrationManager.ps1  rassemble les rapports (module PowerShell, téléchargements .zip/.csv, agent)
  Scripts\04_Inventorier-SharePoint.ps1              inventaire post-migration via Microsoft Graph
  Outils\MigrationManager\                           (facultatif) y extraire le module PowerShell https://aka.ms/MMPowerShellModule
  Donnees\FileServer_Contoso_1.zip et _2.zip         jeu de données en 2 parties (≈1 360 fichiers Office/PDF/images)
  PowerBI\Migration Manager.pbip                     projet Power BI (7 pages, charte Devoteam)
  PowerBI\Secours_Script_TMDL.tmdl                   modèle complet à coller dans la vue TMDL si le .pbip ne s'ouvre pas
  Rapports_Exemple\                                  rapports Migration Manager SIMULÉS pour prévisualiser le dashboard

Prérequis Migration Manager
  - Compte administrateur SharePoint (centre d'administration > Migration center)
  - Poste de l'agent : Windows 10 / Server 2016 ou plus, .NET 4.6.2+, 8 Go RAM minimum, 150 Go libres (dossier de travail)
  - Compte Windows de l'agent : lecture sur le partage \\<PC>\FileServer
  - Source en chemin UNC uniquement (les chemins locaux C:\... sont refusés)

Ordre d'exécution (PowerShell « en tant qu'administrateur » pour le script 01, depuis C:\DemoMigration\Kit\Scripts)
  Set-ExecutionPolicy -Scope Process Bypass
  .\00_Configurer-Parametres.ps1                     -> paramètres de la mission (Entrée = valeur proposée)
  .\01_Preparer-ServeurFichiers.ps1                  -> dossier + partage SMB
  .\05_Audit_Source.ps1 -AvecHash                    -> feux de migrabilité
  .\02_Inventorier-Source.ps1                        -> MigrationManager_Taches.csv + Parametres_Migration_Manager.txt
  -> créer le site + 6 bibliothèques
  -> Migration center > Agents > installer l'agent ; (facultatif) Scans > ajouter les chemins UNC
  -> Migrations > Add task > Bulk migration > charger MigrationManager_Taches.csv, régler les paramètres, Run now
  -> une fois les tâches terminées : Download task report (chaque tâche) + Summary report, dans Téléchargements
  .\03_Collecter-RapportsMigrationManager.ps1 -Zip
  .\04_Inventorier-SharePoint.ps1
  -> ouvrir PowerBI\Migration Manager.pbip, paramètre DossierRapports = C:\DemoMigration\Rapports, Actualiser

Collecte des rapports (script 03)
  1. Module PowerShell Migration Manager s'il est trouvé (Outils\MigrationManager ou -ModuleMigrationManager) :
     Connect-MigrationService, Get-MigrationReport, Get-ScanReport. -SansModule pour l'ignorer.
  2. .zip et .csv récents du dossier Téléchargements (-DepuisJours 14) ou d'un dossier -DossierImport.
  3. Dossier de travail de l'agent %appdata%\Microsoft\SPMigration s'il existe sur ce poste.
  Les CSV sont reconnus d'après leurs colonnes, dédoublonnés et rangés dans Rapports\MigrationManager\.
  Relancer le script après chaque nouveau téléchargement : seuls les nouveaux rapports sont ajoutés.
  Microsoft conserve les rapports 90 jours.

Paramètres de la mission (Scripts\Parametres_Client.psd1)
  Cinq valeurs servent de valeurs par défaut aux scripts 01 à 05 :
    Client                 client / environnement, affiché en tête d'exécution et dans Parametres_Migration_Manager.txt
    CentreAdministration   URL du centre d'administration SharePoint (rappel, aucune connexion)
    SiteUrl                site cible                      -> scripts 02, 04, 05
    CheminLocal            source vue de ce poste          -> scripts 02, 05 ; son dossier parent sert de racine au script 01
    CheminUNC              source vue par l'agent          -> scripts 02, 05 ; vide = détection du partage SMB
    DossierRapports        dossier des rapports            -> scripts 02 à 05 et paramètre Power BI DossierRapports
  .\00_Configurer-Parametres.ps1              questions une par une (Entrée conserve la valeur)
  .\00_Configurer-Parametres.ps1 -Afficher    affiche les valeurs en cours et leur usage
  .\00_Configurer-Parametres.ps1 -Demo        revient aux valeurs de la démonstration Contoso
  .\00_Configurer-Parametres.ps1 -Client 'Fabrikam' -SiteUrl '...' -CheminLocal 'D:\Partages\Commun' ...   sans question
  Un paramètre passé en ligne de commande reste prioritaire sur le fichier.

Choix du dossier de destination
  Chaque script ouvre une fenêtre « Rechercher un dossier » au lancement, présélectionnée sur le dossier
  par défaut (C:\DemoMigration pour le script 01, C:\DemoMigration\Rapports pour les autres).
  Annuler = le script s'arrête sans rien produire.
  -Racine / -Sortie : dossier proposé dans la fenêtre ; -SansFenetre : utilise ce dossier sans fenêtre
  (exécution planifiée). Sans interface graphique, le dossier est demandé dans la console.
  Si le serveur de démonstration est créé ailleurs que C:\DemoMigration, passer -Source '<dossier>\FileServer'
  aux scripts 02 et 05. Conseil : choisir le même dossier de sortie pour 02 à 05, puis le renseigner
  dans le paramètre Power BI DossierRapports. Utiliser un dossier Rapports neuf (pas celui d'une démo SPMT).

Nettoyage : .\01_Preparer-ServeurFichiers.ps1 -Supprimer   (supprime aussi le partage SMB, en administrateur)
