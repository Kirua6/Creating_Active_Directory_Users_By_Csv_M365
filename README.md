# Script de Création d'Utilisateurs Active Directory

Ce script PowerShell facilite la création en masse de comptes utilisateurs dans Active Directory à partir d'un fichier CSV, notamment à partir d'un export M365.
Une fiche [tutoriel](https://github.com/Kirua6/Creating_Active_Directory_Users_By_Csv_M365/blob/main/Fiche_Migration_Donn%C3%A9es_Compte_AD_Profil_Wizard_Git.pdf) est présente pour faire suite à l'utilisation du script, elle permet la migration des applis/ documents de votre ancien compte vers le nouveau avec Profile Wizard.
En cas de problèmes de synchronisation de compte, j'ai aussi créé [un tutoriel](https://github.com/Kirua6/Creating_Active_Directory_Users_By_Csv_M365/blob/main/Fiche_R%C3%A9solution_Doublon_Compte_AD_%26_M365_Git.pdf) pour résoudre ça.
## Fonctionnalités

- Interface graphique interactive pour sélectionner un fichier CSV et une Unité Organisationnelle (OU) dans Active Directory.
- Permet de définir un mot de passe personnalisé pour tous les utilisateurs créés.
- Traite un fichier CSV pour créer des utilisateurs dans l'OU spécifiée avec le mot de passe fourni.

## Prérequis

- Module PowerShell Active Directory installé sur le système exécutant le script.
- Permissions suffisantes pour créer des comptes utilisateurs dans l'Active Directory cible.

## Utilisation

1. Exportez le type d'utilisateur souhaité depuis Microsoft 365 vers un fichier CSV.
2. Assurez-vous que le CSV est bien formaté : ouvrez Excel -> ouvrez le CSV -> choisissez 'Délimité' -> cochez uniquement 'Virgule' -> choisissez 'Standard' -> terminez.
3. Modifiez le UserPrincipalName et l'EmailAddress aux lignes 120 et 125, respectivement, pour correspondre à votre domaine (actuellement réglé sur "@groupe-test.com").
4. Exécutez le script dans PowerShell, sélectionnez votre fichier CSV, et suivez les invites de l'interface graphique.

## Personnalisation

Vous devez éditer le script pour inclure votre UserPrincipalName et EmailAddress.

## Licence

[Licence MIT](https://github.com/Kirua6/Creating_Active_Directory_Users_By_Csv_M365/blob/main/LICENSE)

## Avertissement

Ce script est fourni "tel quel", sans garantie d'aucune sorte de réussit. Utilisez-le avec précaution.
