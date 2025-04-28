# UserInfoForegroundJob - Interface graphique pour recherche AD

## Description
Ce script PowerShell fournit une interface graphique permettant de :
- Saisir un domaine (pré-rempli), un utilisateur et un nombre de jours
- Mettre à jour dynamiquement la liste des Domain Controllers (DC) et les attributs AD disponibles
- Utiliser des foreground jobs pour récupérer les logs de connexion
- Afficher les informations de connexion (WorkstationName, IPAddress)

## Prérequis
- PowerShell 5.1 ou ultérieur
- Module ActiveDirectory installé
- Droits d'administration pour interroger les logs de sécurité

## Compilation
Pour compiler le script en exécutable :
```powershell
ps2exe -InputFile "infos_user.ps1" -OutputFile "infos_user.exe" -noConsole -requireAdmin
```

## Fonctionnalités
1. **Recherche d'informations AD** :
   - Récupère les attributs utilisateur sélectionnables
   - Affiche les logs de connexion (événement 4624)

2. **Interface intuitive** :
   - Sélection dynamique des Domain Controllers
   - Liste des champs AD disponibles
   - Barre de progression et zone de logs

3. **Options** :
   - Période de recherche personnalisable (1-30 jours)
   - Sélection/Désélection rapide des champs AD

## Utilisation
1. Exécutez le script ou l'exécutable compilé
2. Entrez le domaine et le nom d'utilisateur
3. Sélectionnez un Domain Controller
4. Choisissez les champs AD à afficher
5. Cliquez sur "Rechercher"

## Notes
- Le script utilise des foreground jobs pour une meilleure réactivité de l'interface
- Les logs sont affichés en temps réel dans la zone dédiée
- Les erreurs sont capturées et affichées dans la zone de logs
