# TD Print Variables - Add-in Word

Add-in Word pour afficher et insérer facilement les variables TD Print dans vos documents.

## Fonctionnalités

- **Panneau latéral** : Liste complète des variables TD Print organisées par catégories
- **Recherche globale** : Filtrage en temps réel sur le nom, la description et la catégorie
- **Favoris** : Marquez vos variables préférées pour un accès rapide
- **Insertion facile** : Glisser-déposer ou clic pour insérer une variable
- **14 catégories** : Dossier, Client, Commercial, Fournisseur, Produit, etc.

## Types de variables

| Badge | Type | Description |
|-------|------|-------------|
| **C** | Champ | Variable texte standard |
| **B** | Booléen | Variable Oui/Non |
| **I** | Image | Variable image |
| **T** | Tableau | Variable tableau |

## Installation

### 1. Prérequis

- Node.js (version 14 ou supérieure)
- Microsoft Word (desktop ou web)

### 2. Installer les dépendances

```bash
cd tdPrintAddIn
npm install
```

### 3. Générer les icônes (optionnel)

Si vous souhaitez générer les fichiers PNG à partir des SVG :

```bash
npm install sharp --save-dev
npm run generate-icons
```

### 4. Lancer le serveur de développement

```bash
npm start
```

Le serveur démarre sur `https://localhost:3000`.

### 5. Charger l'add-in dans Word

#### Word Desktop (Windows/Mac)

1. Ouvrir Word
2. Aller dans **Insertion** > **Mes compléments** > **Charger mon complément**
3. Parcourir et sélectionner le fichier `manifest.xml`
4. Cliquer sur "Variables" dans l'onglet Accueil pour ouvrir le panneau

#### Word Online

1. Ouvrir Word Online
2. Aller dans **Insertion** > **Compléments Office** > **Charger mon complément**
3. Sélectionner "Manifeste" et parcourir jusqu'au fichier `manifest.xml`

## Structure du projet

```
tdPrintAddIn/
├── manifest.xml           # Manifest de l'add-in Office
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html  # Interface principale
│   │   ├── taskpane.css   # Styles (Fluent UI)
│   │   └── taskpane.js    # Logique (Office.js)
│   └── data/
│       └── variables.js   # Données des variables
├── assets/                # Icônes
├── scripts/               # Scripts utilitaires
├── package.json
└── webpack.config.js
```

## Utilisation

### Recherche

Tapez dans la barre de recherche pour filtrer les variables par :
- Nom du placeholder (ex: `dossier_id`)
- Description (ex: `montant`)
- Catégorie (ex: `Client`)

Raccourci : `Ctrl+F` ou `Cmd+F` pour activer la recherche.

### Favoris

Cliquez sur l'étoile pour ajouter/retirer une variable des favoris.
Les favoris sont sauvegardés localement et persistent entre les sessions.

### Insertion

**Option 1 : Glisser-déposer**
Glissez une variable directement dans votre document Word.

**Option 2 : Bouton Insérer**
Cliquez sur le bouton "Insérer" pour insérer la variable à la position du curseur.

## Build de production

```bash
npm run build
```

Les fichiers de production sont générés dans le dossier `dist/`.

## Catégories de variables

1. **Dossier** - Variables du dossier de financement
2. **Dossier Fournisseur** - Informations côté fournisseur
3. **Client** - Données du client
4. **Commercial** - Informations du commercial
5. **Fournisseur** - Données du fournisseur
6. **Société de Portage (SPV)** - Informations de la SPV
7. **Produit** - Variables des produits financés
8. **Produit Assurance** - Assurance des produits
9. **Caution** - Informations des cautions
10. **Mandat Caution** - Mandats de caution
11. **Crédit Fournisseur** - Crédits fournisseur
12. **Attestation Prix Net** - Attestations de prix
13. **Société (Assurance)** - Société d'assurance
14. **Simulation** - Variables de simulation

## Configuration HTTPS

Le serveur de développement utilise HTTPS (requis par Office Add-ins).
Si vous rencontrez des erreurs de certificat, acceptez le certificat auto-signé dans votre navigateur.

## Dépannage

### L'add-in ne se charge pas

1. Vérifiez que le serveur est démarré (`npm start`)
2. Vérifiez que l'URL `https://localhost:3000/taskpane.html` est accessible
3. Acceptez le certificat SSL auto-signé

### Les variables ne s'insèrent pas

1. Vérifiez que le curseur est positionné dans le document
2. Assurez-vous que le document n'est pas en lecture seule

## Licence

MIT
