# Outil de Gestion Active Directory 🚀

[![PowerShell](https://img.shields.io/badge/Made%20with-PowerShell-0078D4?style=for-the-badge&logo=powershell)](https://docs.microsoft.com/en-us/powershell/)
[![GitHub Stars](https://img.shields.io/github/stars/KiralyGeddon/Active-Directory-Administration-Tools?style=for-the-badge&color=brightgreen)](https://github.com/KiralyGeddon/Active-Directory-Administration-Tools/stargazers)
[![GitHub Forks](https://img.shields.io/github/forks/KiralyGeddon/Active-Directory-Administration-Tools?style=for-the-badge&color=blue)](https://github.com/KiralyGeddon/Active-Directory-Administration-Tools/network/members)

## Table des matières

* [🌟 À propos du projet](#-à-propos-du-projet)
* [✨ Fonctionnalités](#-fonctionnalités)
* [🚀 Démarrage rapide](#-démarrage-rapide)
    * [Prérequis](#prérequis)
    * [Installation](#installation)
    * [Utilisation](#utilisation)
* [📸 Aperçu](#-aperçu)
* [🤝 Contribuer](#-contribuer)
* [📧 Contact](#-contact)
* [🙏 Remerciements](#-remerciements)

## 🌟 À propos du projet

Bienvenue sur le **Outil de Gestion Active Directory** ! Ce projet est une interface graphique (GUI) développée en PowerShell, conçue pour simplifier et automatiser les tâches de gestion courantes dans un environnement Active Directory. Fini la ligne de commande complexe pour les opérations quotidiennes ! Cet outil est idéal pour les administrateurs systèmes, les techniciens support, ou toute personne ayant besoin de gérer des utilisateurs, des groupes ou des unités d'organisation de manière efficace et intuitive.

**Pourquoi cet outil ?**
* **Simplicité** : Une interface utilisateur claire et facile à prendre en main.
* **Automatisation** : Exécute des tâches complexes en quelques clics.
* **Gain de temps** : Idéal pour les opérations répétitives (importation en masse, création/suppression rapide).
* **Fiabilité** : Basé sur les cmdlets PowerShell d'Active Directory, garantissant robustesse et conformité.
* **Internationalisé** : Disponible en plusieurs langues (français et anglais pour le moment).

## ✨ Fonctionnalités

Cet outil vous permet de :

* **Importation d'utilisateurs en masse via CSV** : Créez des dizaines ou des centaines d'utilisateurs en un seul fichier.
* **Gestion des utilisateurs individuelle** :
    * Ajouter de nouveaux utilisateurs avec leurs informations de base.
    * Supprimer des utilisateurs existants.
* **Gestion des groupes Active Directory** :
    * Créer de nouveaux groupes.
    * Supprimer des groupes existants.
    * Ajouter des utilisateurs à des groupes.
* **Gestion des Unités d'Organisation (OU)** :
    * Créer de nouvelles OUs.
    * Supprimer des OUs (doivent être vides).
* **Journalisation des opérations** : Suivez toutes les actions effectuées directement dans l'interface.
* **Support multilingue** : Basculez entre le français et l'anglais à la volée.
* **Vérification des prérequis** : Le script vérifie automatiquement les droits d'administrateur et l'installation des outils RSAT Active Directory, proposant même l'installation si nécessaire.

## 🚀 Démarrage rapide

Suivez ces étapes pour démarrer l'outil.

### Prérequis

* Un ordinateur sous Windows (client ou serveur) ayant accès à un contrôleur de domaine Active Directory.
* Des droits d'administrateur local sur la machine exécutant le script (le script tentera de s'élever si nécessaire).
* Le module PowerShell `ActiveDirectory` (partie des outils RSAT pour AD DS et AD LDS). Le script propose une installation automatique si ce module est manquant.
* Une connexion internet pour l'installation automatique des RSAT si besoin.

### Installation

1.  **Clonez le dépôt** :
    ```bash
    git clone [https://github.com/KiralyGeddon/Active-Directory-Administration-Tools.git](https://github.com/KiralyGeddon/Active-Directory-Administration-Tools.git)
    ```
2.  **Naviguez vers le dossier du projet** :
    ```bash
    cd Active-Directory-Administration-Tools
    ```
3.  **Vérifiez la politique d'exécution de PowerShell** :
    Assurez-vous que votre politique d'exécution PowerShell permet l'exécution de scripts locaux. Si ce n'est pas le cas, vous pouvez la définir temporairement :
    ```powershell
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
    ```

### Utilisation

Pour lancer l'application, exécutez le script `main.ps1` en tant qu'administrateur PowerShell :

```powershell
.\main.ps1
```

**Note** : Le script s'assurera que PowerShell est lancé avec les privilèges administrateur.

## 📸 Aperçu

Voici quelques captures d'écran de l'application en action :

**Écran Principal - Importation d'Utilisateurs via CSV**

![Ecran Principal](images\1.png)]

**Gestion des Utilisateurs**

![Gestion des Utilisateurs](images\2.png)]

**Gestion des Groupes**

![Gestion des Groupes](images\3.png)]

**Gestion des Unités d'Organisation (OU)**

![Gestion des Unités d'Organisation (OU)](images\4.png)]

**Informations de mise à jour**

![Informations de mise à jour](images\5.png)]

## 🤝 Contribuer

Les contributions sont les bienvenues ! Si vous avez des idées d'amélioration, des rapports de bugs ou des propositions de nouvelles fonctionnalités, n'hésitez pas à :
*	Faire un "Fork" du projet.
*	Créer une branche pour votre fonctionnalité (git checkout -b feature/nouvelle-fonctionnalite).
*	Commiter vos changements (git commit -m 'Ajout d'une nouvelle fonctionnalité').
*	Pousser vers la branche (git push origin feature/nouvelle-fonctionnalite).
*	Ouvrir une "Pull Request".

## 📧 Contact

Pour toute question ou suggestion, n'hésitez pas à contacter Kiraly Geddon.
*	GitHub : [https://github.com/KiralyGeddon](Kiraly Geddon)

## 🙏 Remerciements

*	À la communauté PowerShell pour ses ressources et son support.
*	À Google Gemini pour l'aide à la refactorisation et l'ajout de fonctionnalités.

