# UserInfoForegroundJob - Interface graphique pour recherche AD

## Description
Ce projet fournit **deux applications distinctes** basées sur PowerShell avec interface graphique Windows Forms :

- **infos_user.exe** :  
  ➔ Version standard qui permet de récupérer **uniquement les informations d'attributs Active Directory** d'un utilisateur.  
  ➔ **Ne nécessite pas de droits administrateur**.

- **infos_user_connexion.exe** :  
  ➔ Version avancée qui récupère **les informations d'attributs Active Directory** **et** **les historiques de connexion** (workstation et IP) en consultant les **logs de sécurité** des contrôleurs de domaine.  
  ➔ **Nécessite des droits d'administrateur du domaine** pour interroger les journaux de sécurité.

## Tableau comparatif

| Fonctionnalité | infos_user.exe | infos_user_connexion.exe |
|:---------------|:--------------:|:------------------------:|
| Lecture attributs AD | ✅ | ✅ |
| Lecture logs de sécurité (4624) | ❌ | ✅ |
| Droits admin locaux | ❌ | ✅ |
| Droits admin domaine | ❌ | ✅ |
| Export CSV | ✅ | ✅ |

## Prérequis
- Windows avec **PowerShell 5.1** ou ultérieur
- **Module ActiveDirectory** installé (RSAT)
- **infos_user_connexion.exe** :
  - ⭐ Membre du groupe **Domain Admins** ou équivalent
  - ⭐ Permission de lecture sur le journal de sécurité des DCs

## Compilation
Utilisation de [ps2exe](https://github.com/MScholtes/PS2EXE) pour convertir les scripts `.ps1` en `.exe`.

- **Version standard** (infos utilisateurs uniquement) :
```powershell
ps2exe -InputFile "infos_user.ps1" -OutputFile "infos_user.exe" -noConsole
```

- **Version avancée** (infos utilisateurs + connexions) :
```powershell
ps2exe -InputFile "infos_user_connexion.ps1" -OutputFile "infos_user_connexion.exe" -noConsole -requireAdmin
```

## Fonctionnalités

### 1. Recherche d'informations Active Directory
- Saisie du **domaine**, du **nom d'utilisateur**, et d'une **période en jours**.
- Mise à jour dynamique des **Domain Controllers (DCs)**.
- Affichage des **attributs AD** sélectionnés.

### 2. Historique de connexions (version connexion seulement)
- Récupération des événements 4624 depuis les logs de sécurité du DC.
- Affichage des **postes** et **IPs** utilisés.

### 3. Interface utilisateur intuitive
- Windows Forms ergonomique.
- Logs d'activité en temps réel.
- ProgressBar pendant les recherches.

### 4. Fonctions avancées
- **Export CSV** des résultats.
- **Copie rapide** à partir du DataGridView.
- **Sélection rapide** des attributs AD.

## Utilisation
1. Lancez l'exécutable adapté à votre besoin.
2. Renseignez le domaine, l'utilisateur, et la période.
3. Sélectionnez un DC.
4. Sélectionnez les attributs AD.
5. Cliquez sur **Rechercher**.

## Notes importantes
- **infos_user.exe** : fonctionnement **sans élévation de privilèges**.
- **infos_user_connexion.exe** : fonctionnement **avec élévation UAC** **et droits d'administration domaine**.
- Les erreurs sont capturées et affichées dans la zone de logs.
- Utilisation de **Runspace asynchrone** pour conserver une interface fluide.

