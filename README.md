# Détecteur Intelligent de Doublons Étudiants

## Description

Ce script Python détecte et regroupe les **doublons** dans une liste d'étudiants à partir d'un fichier CSV. Il identifie les situations suivantes :

1. **Matricules dupliqués** : Même matricule utilisé plusieurs fois (après nettoyage)
2. **Étudiants multi-matricule** : Même étudiant (nom + prénom) avec différents matricules

Le script génère un **fichier Excel coloré et structuré** avec :

- Groupes de doublons regroupés et coloriés
- Statistiques détaillées
- Légende des groupes

## Fonctionnalités principales

### 🔍 Détection intelligente

- **Insensible à la casse** : "Martin" = "martin"
- **Insensible aux accents** : "Château" = "CHATEAU"
- **Nettoyage des matricules** : Supprime les suffixes `-ANNULE` et `--ANNULE`
- **Regroupement automatique** : Fusionne les groupes qui partagent des indices

### 📊 Résultats

- **Groupes colorés** : Chaque groupe reçoit une couleur unique
- **Bordures épaisses** : Identifient les limites entre groupes
- **Statistiques** : Nombre total de doublons, groupes détectés, types
- **Légende** : Tableau de correspondance groupe/couleur

### 📋 Colonnes de sortie

- `matricule` : Matricule original
- `nom` : Nom de l'étudiant
- `prenom` : Prénom de l'étudiant
- `sexe` : Sexe
- `Groupe` : Identifiant du groupe (G1, G2, ...) ou vide si unique
- `Type_doublon` : Type détecté ("MATRICULE DUPLIQUÉ", "ÉTUDIANT MULTI-MATRICULE", etc.)

## Installation

### Dépendances

```bash
pip install pandas openpyxl unidecode
```

## Utilisation

### Syntaxe

```bash
python detecteur_groupes.py <input.csv> <output.xlsx>
```

### Exemple

```bash
python detecteur_groupes.py etudiants_vak.csv resultat_groupes.xlsx
```

## Format d'entrée (CSV)

Le fichier CSV doit être **délimité par des points-virgules (`;`)** et contenir au minimum les colonnes :

- `matricule` : Identifiant unique (peut contenir -ANNULE)
- `nom` : Nom de l'étudiant
- `prenom` : Prénom de l'étudiant
- `sexe` : (optionnel mais recommandé)

### Exemple

```csv
matricule;nom;prenom;sexe
M001;Dupont;Jean;M
M001;DUPONT;jean;M
M002;Château;Marie-Claire;F
M003;chateau;MARIE CLAIRE;F
M004-ANNULE;Durand;Pierre;M
```

**Résultat** : Les 4 premières lignes seront groupées (même étudiant, même matri), le 5e sera ignoré (annulé)

## Résultats de sortie

### Fichier Excel généré

Un fichier `.xlsx` contenant :

1. **Feuille "Doublons_Intelligents"**

   - Tous les étudiants (doublons ET uniques)
   - Doublons regroupés et coloriés
   - Uniques en bas sans groupe

2. **Légende**

   - Liste des groupes (G1, G2, ...) avec leurs couleurs

3. **Mise en forme**
   - En-tête bleu foncé avec texte blanc
   - Bordures fines pour les cellules individuelles
   - Bordures épaisses marquant les transitions de groupe
   - Colonnes automatiquement ajustées

### Statistiques affichées en console

```
==============================================================================
                         RÉSULTATS FINAUX
==============================================================================
Total lignes lues                  : 5503
Étudiants en doublon               : 192
Groupes de doublons détectés       : 74
   → Matricules répétés            : 45
   → Étudiants avec plusieurs matricules : 29
==============================================================================
```

## Logique de groupement

Le script regroupe les doublons en deux étapes :

1. **Matricules identiques** : Toutes les lignes partageant un même `matricule_clean` sont groupées
2. **Étudiants multi-matricule** : Si un même étudiant (nom+prénom normalisé) a plusieurs matricules, tous les indices associés sont ajoutés au groupe

Quand deux groupes partagent au moins un indice, ils fusionnent automatiquement.

## Cas d'usage

✅ Détection de **doublets de saisie** (même étudiant enregistré deux fois)  
✅ Détection d'**erreurs de matricules** (même matricule pour deux personnes)  
✅ Détection d'**erreurs de normalisation** (accents, casse, espaces)  
✅ Identification d'**étudiants avec plusieurs matricules** (changements administratifs)

## Améliorations apportées

- ✅ Insensible à la casse et aux accents (utilise `unidecode`)
- ✅ Nettoyage intelligent des matricules (-ANNULE supprimé)
- ✅ Regroupement hiérarchique par matricule ET étudiant
- ✅ Couleurs progressives + générations aléatoires si trop de groupes
- ✅ Statistiques détaillées en sortie
- ✅ Excel formaté avec bordures et couleurs

## Dépannage

### "ERREUR : Fichier 'xxx' introuvable."

Vérifiez que le chemin du fichier CSV est correct et que le fichier existe.

### "ERREUR : can only concatenate list..."

Assurez-vous que toutes les lignes du CSV contiennent les colonnes obligatoires.

### Caractères mal affichés

Le script utilise `unidecode` pour normaliser les accents. Vérifiez que votre terminal supporte UTF-8.

## Auteur

Créé pour automatiser la détection de doublons dans les listes d'étudiants ESGC VAK.

## Licence

Usage libre et gratuit.
