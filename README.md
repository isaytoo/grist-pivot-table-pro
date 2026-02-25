# Grist Pivot Table Pro Widget

Un widget de tableau croisé dynamique **sans limite de données** pour Grist.

## ✨ Fonctionnalités

- **Pas de limite de données** - Contrairement à WebDataRocks (1 MB max)
- **Glisser-déposer** - Interface intuitive pour configurer le pivot
- **Agrégations multiples** :
  - Somme, Nombre, Moyenne, Min, Max
  - Distinct, Médiane, Écart-type, Variance
- **Export CSV** - Exportez vos tableaux croisés
- **Bilingue** - Français / Anglais
- **Sauvegarde automatique** - La configuration est sauvegardée dans Grist

## 🚀 Installation

### Option 1 : URL directe
Ajoutez un widget personnalisé dans Grist avec l'URL :
```
https://isaytoo.github.io/grist-pivot-table-pro/
```

### Option 2 : Auto-hébergé
1. Clonez ce repo
2. Servez les fichiers via un serveur web
3. Ajoutez l'URL dans Grist

## 📖 Utilisation

1. Sélectionnez une table source
2. Glissez des champs dans les zones :
   - **Lignes** : Champs pour les lignes du tableau
   - **Colonnes** : Champs pour les colonnes du tableau
   - **Valeurs** : Champs à agréger (avec choix de la fonction)
3. Le tableau croisé se génère automatiquement
4. Exportez en CSV si besoin

## 🔧 Permissions requises

- `read table` - Pour lire les données des tables

## 📝 Licence

Apache License 2.0

## 👤 Auteur

**Said Hamadou (isaytoo)**
- GitHub: [@isaytoo](https://github.com/isaytoo)
- Site: [gristup.fr](https://gristup.fr)
