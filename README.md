# 📋 RFP Respond - Analyseur Intelligent d'Appels d'Offres (RFP)

`rfp-respond` est un outil d'automatisation conçu pour aider les équipes de vente et techniques à analyser rapidement des dossiers d'appels d'offres (RFP - Request for Proposal) au format PDF.

Il utilise l'intelligence artificielle (Claude d'Anthropic) pour extraire les exigences, identifier les besoins implicites, et générer des propositions de réponse.

---

## 🚀 Fonctionnalités

1.  **Extraction Multidimensionnelle** : Lance 5 analyses en parallèle pour ne rien oublier :
    *   **Fonctionnel** : Ce que le système doit faire.
    *   **Contraintes** : Performance, sécurité, conformité, etc.
    *   **Parties prenantes** : Besoins spécifiques par rôle (admin, utilisateur, auditeur).
    *   **Implicite** : Ce que le client a oublié de demander mais qui est crucial (sauvegarde, formation, RGPD).
    *   **Signaux faibles** : Détection des "douleurs" passées du client à travers ses formulations.
2.  **Fusion & Déduplication** : Regroupe les résultats, élimine les doublons et trie par priorité MoSCoW.
3.  **Enrichissement IA** : Génère automatiquement une première ébauche de réponse pour chaque exigence.
4.  **Double Export** :
    *   **Excel** : Une matrice de conformité complète et filtrable pour le pilotage.
    *   **Word** : Un document de réponse structuré avec page de garde et synthèse exécutive.

---

## 🛠️ Installation (pour débutants)

### 1. Prérequis
*   Avoir **Python 3.10+** installé sur votre machine.
*   Une clé API **Anthropic** (Claude).

### 2. Cloner ou télécharger le projet
Placez-vous dans le dossier du projet.

### 3. Créer un environnement virtuel (recommandé)
C'est un "bac à sable" pour que les bibliothèques du projet n'interfèrent pas avec votre système.
```bash
# Windows
python -m venv venv
.\venv\Scripts\activate

# Linux / macOS
python3 -m venv venv
source venv/bin/activate
```

### 4. Installer les dépendances
```bash
pip install -r requirements.txt
```

### 5. Configurer les variables d'environnement
Copiez le fichier `.env.example` et renommez-le en `.env`.
Ouvrez le fichier `.env` et ajoutez votre clé API :
```env
ANTHROPIC_API_KEY=votre_cle_ici
CLAUDE_MODEL=claude-3-5-haiku-20241022
```

---

## 📖 Utilisation

Le script principal est `main.py`.

### Commande de base
Analyse un PDF et génère les exports dans le dossier `output/` :
```bash
python main.py --input mon_appel_offres.pdf
```

### Options disponibles
*   `--output-dir ./resultats` : Spécifier le dossier de sortie.
*   `--no-responses` : Ne génère pas les propositions de réponse (beaucoup plus rapide et moins coûteux en jetons API). Utile pour une première lecture rapide.

---

## 🏗️ Structure du Projet

*   `main.py` : Le chef d'orchestre. Il gère le flux de données entre les modules.
*   `readers.py` : Contient les "agents" de lecture. Chaque agent a un rôle spécifique et un prompt IA dédié.
*   `matrix.py` : La logique métier. Fusionne les lectures, nettoie les données et demande à l'IA de rédiger les réponses.
*   `exporters/` :
    *   `excel_export.py` : Mise en forme du fichier Excel (couleurs, filtres, onglets).
    *   `word_export.py` : Génération du document Word (styles, mise en page).

---

## 💡 Concepts clés (MoSCoW)

L'outil classe les exigences selon la méthode MoSCoW :
*   **Must** : Obligatoire (le projet échoue sans cela).
*   **Should** : Important mais non vital.
*   **Could** : Souhaitable (confort).
*   **Won't** : Hors périmètre pour cette fois.

---

## 🔒 Sécurité
*   Ne partagez jamais votre fichier `.env` ou votre clé API.
*   Le fichier `.env` est déjà listé dans `.gitignore` pour éviter tout envoi accidentel sur GitHub.
