# 📂 Archivage intelligent Gmail vers Drive

![Version](https://img.shields.io/badge/version-5.0.0-blue.svg)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

## 📝 Description

Ce projet est une solution d'automatisation robuste ("Set and Forget") pour Google Workspace. Il surveille une boîte Gmail, détecte les factures ou documents importants, et assure leur archivage pérenne dans Google Drive.

Contrairement aux scripts basiques, cette solution gère la **conversion de formats** (Images vers PDF), le **classement dynamique** (Année/Mois) et la **normalisation des noms de fichiers** pour garantir un archivage propre et consultable.

## ✨ Fonctionnalités clés

* **🔍 Filtrage Précis :** Cible uniquement les e-mails non lus portant un libellé spécifique (ex: "Facture").
* **🔄 Conversion à la volée :** Transforme automatiquement les images (JPEG, PNG) en PDF. Les PDF natifs sont conservés tels quels.
* **🗂️ Classement Chronologique :** Crée et organise automatiquement les dossiers dans Drive selon la date de réception de l'e-mail (`Racine > Année > Mois`).
* **🏷️ Renommage Intelligent :** Standardise les noms de fichiers pour une lisibilité parfaite :
    * *Format :* `AAAA-MM-JJ_Expéditeur_NomOriginal.pdf`
    * *Exemple :* `2024-11-30_Amazon_Facture_X99.pdf`
* **🚀 Performance & Quotas :** Utilise des opérations par lots (Batch Operations) pour marquer les e-mails comme lus, minimisant les appels API.
* **📧 Reporting d'Incidents :** Envoie automatiquement un rapport HTML à l'administrateur en cas d'erreur (fichier corrompu, échec Drive, etc.).

## ⚙️ Prérequis

* Un compte Google Workspace ou Gmail.
* Un libellé créé dans Gmail pour identifier les e-mails à traiter (par défaut : "Facture").

## 🚀 Installation

1.  Ouvrez [Google Apps Script](https://script.google.com/).
2.  Créez un nouveau projet.
3.  Copiez le contenu du fichier `Code.js` de ce dépôt dans l'éditeur.
4.  Modifiez la constante de configuration au début du script :

```javascript
const CONFIGURATION = {
  LIBELLE_GMAIL: "Facture",          // Le libellé Gmail à surveiller
  NOM_DOSSIER_RACINE: "Factures",    // Le dossier racine dans Drive
  EMAIL_ADMIN: "votre@email.com",    // Pour recevoir les rapports d'erreurs
  TYPES_ACCEPTES: ["application/pdf", "image/jpeg", "image/png"]
};
