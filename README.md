# Sharepoint_Migration — kit Migration Manager + Power BI

Kit de démonstration et d'outillage pour migrer un **serveur de fichiers vers SharePoint Online** avec **Migration Manager**, puis piloter la migration dans **Power BI**.

Toutes les données du jeu de démonstration sont **fictives** (société Contoso).

## Contenu

| Dossier | Rôle |
|---|---|
| `Scripts/00_Configurer-Parametres.ps1` | Paramètres de la mission (client, centre d'administration, site cible, chemin local/UNC, dossier des rapports) → `Scripts/Parametres_Client.psd1` |
| `Scripts/01_Preparer-ServeurFichiers.ps1` | Crée le serveur de fichiers de démonstration (≈ 1 354 fichiers, 25 cas de test) et le partage SMB |
| `Scripts/02_Inventorier-Source.ps1` | Inventaire source en chemins UNC + `MigrationManager_Taches.csv` (tâches en masse, 6 colonnes) |
| `Scripts/03_Collecter-RapportsMigrationManager.ps1` | Rassemble les rapports (module PowerShell, téléchargements `.zip`/`.csv`, dossier de l'agent) |
| `Scripts/04_Inventorier-SharePoint.ps1` | Inventaire post-migration via Microsoft Graph |
| `Scripts/05_Audit_Source.ps1` | Audit de migrabilité : 42 contrôles, feu vert/orange/rouge par bibliothèque |
| `PowerBI/Migration Manager.pbip` | Projet Power BI, 7 pages |
| `Donnees/` | Jeu de données fictif (serveur de fichiers de démonstration) |
| `Outils/` | Outils complémentaires (module Migration Manager) |
| `Rapports_Exemple/` | Rapports Migration Manager **simulés**, pour prévisualiser le dashboard |

`README.txt` (à la racine) donne le détail complet : prérequis, ordre d'exécution, collecte des rapports, nettoyage.

## Paramètres de la mission

Cinq valeurs décrivent la mission et servent de valeurs par défaut aux scripts 01 à 05 :

| Clé de `Scripts/Parametres_Client.psd1` | Sert à |
|---|---|
| `Client` | En-tête des exécutions et du fichier `Parametres_Migration_Manager.txt` |
| `CentreAdministration` | Rappel de l'URL du centre d'administration SharePoint |
| `SiteUrl` | Site cible (scripts 02, 04, 05 ; colonne *SharePointSite* du fichier de tâches) |
| `CheminLocal` | Source vue de ce poste (scripts 02, 05) ; son dossier parent est la racine du script 01 |
| `CheminUNC` | Source vue par l'agent ; vide = détection du partage SMB |
| `DossierRapports` | Sortie des scripts 02 à 05 et paramètre Power BI `DossierRapports` |

```powershell
.\00_Configurer-Parametres.ps1            # questions une par une, Entrée conserve la valeur
.\00_Configurer-Parametres.ps1 -Afficher  # valeurs en cours et scripts concernés
.\00_Configurer-Parametres.ps1 -Demo      # valeurs de la démonstration Contoso
```

Un paramètre passé en ligne de commande reste prioritaire sur le fichier.

## Ordre d'exécution

```powershell
cd Kit\Scripts
Set-ExecutionPolicy -Scope Process Bypass
.\00_Configurer-Parametres.ps1
.\01_Preparer-ServeurFichiers.ps1     # console administrateur (création du partage SMB)
.\05_Audit_Source.ps1 -AvecHash
.\02_Inventorier-Source.ps1
# -> créer le site et les 6 bibliothèques, installer l'agent,
#    charger MigrationManager_Taches.csv dans Migration center > Add task > Bulk migration
.\03_Collecter-RapportsMigrationManager.ps1 -Zip
.\04_Inventorier-SharePoint.ps1
# -> ouvrir PowerBI\Migration Manager.pbip, DossierRapports = votre dossier de rapports, Actualiser
```

## Prérequis

- Compte administrateur SharePoint (centre d'administration > Migration center)
- Poste de l'agent : Windows 10 / Server 2016 ou plus, .NET 4.6.2+, 8 Go de RAM minimum, 150 Go libres
- Compte Windows de l'agent : lecture sur le partage source
- Source en **chemin UNC uniquement** (`\\serveur\partage`) : Migration Manager refuse les chemins locaux
- Windows PowerShell 5.1 ou PowerShell 7 ; Power BI Desktop récent pour le `.pbip`

## Avertissements

- Le script 01 **crée un partage SMB** et écrit dans le dossier choisi : à réserver à un poste de démonstration.
- Les scripts 02 à 05 sont en lecture seule sur la source ; ils n'écrivent que dans le dossier des rapports.
- Les rapports Migration Manager ne sont conservés que **90 jours** côté Microsoft : collectez-les avant expiration.
- Les documents de mise en œuvre Devoteam (mode opératoire, note de contraintes) ne sont pas publiés ici.
