# Installation de l'Add-in VBA TD Print Variables

## Fichiers fournis

- `ModVariables.bas` - Module contenant toutes les variables TD Print
- `ModMain.bas` - Module principal avec la macro de lancement
- `FrmTDPrintVariables.frm` - Code du UserForm (interface)

## Étapes d'installation

### 1. Ouvrir l'éditeur VBA

1. Ouvrir Word
2. Appuyer sur `Alt + F11` pour ouvrir l'éditeur VBA
3. Dans le menu, aller dans `Fichier > Importer fichier...`

### 2. Importer les modules

1. Importer `ModVariables.bas`
2. Importer `ModMain.bas`

### 3. Créer le UserForm

1. Dans l'éditeur VBA, clic droit sur le projet > `Insertion > UserForm`
2. Renommer le UserForm en `FrmTDPrintVariables` (dans les propriétés, changer `(Name)`)
3. Configurer les propriétés du formulaire :
   - Caption: `TD Print Variables`
   - Width: `360`
   - Height: `450`
   - StartUpPosition: `1 - CenterOwner`

### 4. Ajouter les contrôles au UserForm

Créez les contrôles suivants dans l'ordre (utilisez la boîte à outils) :

#### Zone de recherche (en haut)

| Contrôle | Nom | Propriétés |
|----------|-----|------------|
| Label | `lblSearch` | Caption: `Rechercher :`, Top: 10, Left: 10 |
| TextBox | `txtSearch` | Top: 10, Left: 75, Width: 200, Height: 20 |
| CommandButton | `btnClearSearch` | Caption: `X`, Top: 10, Left: 280, Width: 25, Height: 20 |

#### Section Favoris

| Contrôle | Nom | Propriétés |
|----------|-----|------------|
| Label | `lblFavorites` | Caption: `★ Favoris`, Top: 40, Left: 10, Font: Bold |
| Label | `lblFavoritesCount` | Caption: `0 favori(s)`, Top: 40, Left: 250 |
| ListBox | `lstFavorites` | Top: 58, Left: 10, Width: 335, Height: 60 |
| CommandButton | `btnRemoveFavorite` | Caption: `Retirer`, Top: 120, Left: 260, Width: 85 |

#### Section Catégories

| Contrôle | Nom | Propriétés |
|----------|-----|------------|
| Label | `lblCategories` | Caption: `Catégories`, Top: 150, Left: 10, Font: Bold |
| ListBox | `lstCategories` | Top: 168, Left: 10, Width: 335, Height: 80 |

#### Section Variables

| Contrôle | Nom | Propriétés |
|----------|-----|------------|
| Label | `lblVariablesTitle` | Caption: `Variables`, Top: 255, Left: 10, Font: Bold |
| Label | `lblVariablesCount` | Caption: `0 variable(s)`, Top: 255, Left: 250 |
| ListBox | `lstVariables` | Top: 273, Left: 10, Width: 335, Height: 100 |
| CommandButton | `btnAddFavorite` | Caption: `★ Ajouter aux favoris`, Top: 378, Left: 10, Width: 130 |

#### Boutons d'action (en bas)

| Contrôle | Nom | Propriétés |
|----------|-----|------------|
| CommandButton | `btnInsert` | Caption: `Insérer`, Top: 410, Left: 10, Width: 100, Height: 28, Default: True |
| CommandButton | `btnClose` | Caption: `Fermer`, Top: 410, Left: 245, Width: 100, Height: 28, Cancel: True |

### 5. Copier le code du UserForm

1. Double-cliquer sur le UserForm pour ouvrir la fenêtre de code
2. Copier tout le contenu du fichier `FrmTDPrintVariables.frm` (à partir de `Option Explicit`)
3. Coller dans la fenêtre de code du UserForm

### 6. Sauvegarder comme template

1. Dans Word, aller dans `Fichier > Enregistrer sous`
2. Choisir le type `Modèle Word prenant en charge les macros (*.dotm)`
3. Nommer le fichier `TDPrintVariables.dotm`
4. Sauvegarder dans le dossier des modèles Word :
   - Windows: `%appdata%\Microsoft\Word\STARTUP`
   - Mac: `~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word`

### 7. Ajouter un bouton dans le ruban (optionnel)

1. Dans Word, clic droit sur le ruban > `Personnaliser le ruban`
2. Dans la colonne de droite, créer un nouveau groupe
3. Dans la colonne de gauche, choisir `Macros` dans le menu déroulant
4. Sélectionner `TDPrint` et cliquer sur `Ajouter`
5. Renommer le bouton en "TD Print Variables"

## Utilisation

### Lancer l'add-in

- **Via la macro** : `Alt + F8` > Sélectionner `TDPrint` > Exécuter
- **Via le ruban** : Cliquer sur le bouton "TD Print Variables" (si configuré)

### Interface

1. **Recherche** : Tapez pour filtrer les variables par nom, description ou catégorie
2. **Favoris** : Vos variables préférées, accessibles rapidement
3. **Catégories** : Cliquez sur une catégorie pour voir ses variables
4. **Variables** : Double-cliquez ou cliquez sur "Insérer" pour ajouter au document
5. **Ajouter aux favoris** : Sélectionnez une variable et cliquez sur le bouton étoile

### Raccourcis

- Double-clic sur une variable = Insertion immédiate
- Entrée = Insérer la variable sélectionnée
- Échap = Fermer le panneau

## Dépannage

### "Les macros ont été désactivées"

1. `Fichier > Options > Centre de gestion de la confidentialité`
2. `Paramètres du Centre de gestion de la confidentialité`
3. `Paramètres des macros`
4. Sélectionner "Activer toutes les macros" ou "Désactiver avec notification"

### Le formulaire ne s'affiche pas correctement

Vérifiez que tous les contrôles ont les bons noms (Name) dans leurs propriétés.

### Les favoris ne sont pas sauvegardés

Les favoris sont stockés dans le registre Windows. Vérifiez que vous avez les droits d'écriture.
