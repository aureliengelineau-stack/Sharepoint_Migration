#
#  Paramètres de la mission - kit Devoteam « Migration Manager »
#  --------------------------------------------------------------
#  Ce fichier est lu automatiquement par les scripts 01 à 05 : ses valeurs deviennent
#  les valeurs par défaut de leurs paramètres. Un paramètre passé en ligne de commande
#  reste prioritaire sur ce fichier.
#
#  Pour le remplir sans éditer ce fichier :  .\00_Configurer-Parametres.ps1
#  Pour revenir aux valeurs de la démo     :  .\00_Configurer-Parametres.ps1 -Demo
#
#  Laisser une valeur vide ('') = garder la valeur par défaut du script.
#
@{

    # Client / environnement : sert d'en-tête aux exécutions et aux rapports.
    Client = 'Contoso · tenant de démonstration M365x71797824'

    # URL du centre d'administration SharePoint (centre de migration, agents, tâches).
    # Utilisée pour l'affichage et les rappels : les scripts ne s'y connectent pas.
    CentreAdministration = 'https://m365x71797824-admin.sharepoint.com'

    # Site cible : site SharePoint qui reçoit les bibliothèques (scripts 02, 04, 05).
    SiteUrl = 'https://m365x71797824.sharepoint.com/sites/MigrationFileServer'

    # Chemin local de la source, vu depuis ce poste (scripts 02 et 05).
    # Son dossier parent sert de racine au script 01 (données + dossier Rapports).
    CheminLocal = 'C:\DemoMigration\FileServer'

    # Chemin UNC de la source, vu par l'agent Migration Manager (obligatoire pour les tâches).
    # Vide = les scripts cherchent seuls le partage SMB qui contient CheminLocal.
    CheminUNC = ''

    # Dossier des rapports : sortie des scripts 02 à 05, et paramètre DossierRapports de Power BI.
    DossierRapports = 'C:\DemoMigration\Rapports'

}
