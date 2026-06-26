# business_plan_app
 🚀 Application Streamlit de création de business plan guidée en 9 étapes. Répondez aux questions, obtenez automatiquement un Business Plan Word complet et un Plan Financier Excel avec projections 3 ans et graphiques. Basé sur le guide Legalstart 2026. 
# 🚀 Créateur de Business Plan — Streamlit App

> Application interactive qui transforme vos réponses en un **Business Plan Word professionnel** et un **Plan Financier Excel complet**, en quelques minutes.

Basée sur le **Guide Pratique de Création d'Entreprise — Édition 2026 (Legalstart)**, cette app guide n'importe quel entrepreneur, même sans connaissances juridiques ou financières, à travers 9 étapes claires pour structurer et formaliser son projet.

---

## 📸 Aperçu

```
Accueil → Votre Idée → Votre Projet → Étude de Marché → Produit/Service
       → Équipe → Financement → Projections → Statut Juridique → 📄 Téléchargement
```

---

## ✨ Fonctionnalités

| Fonctionnalité | Détail |
|---|---|
| 🧭 **9 étapes guidées** | Questionnaire structuré, barre de progression, navigation intuitive |
| 📄 **Business Plan Word** | Page de garde, sommaire, 10 sections mises en page avec tableaux et couleurs |
| 📊 **Plan Financier Excel** | 5 onglets, projections 3 ans, graphiques automatiques |
| ⚖️ **Statuts juridiques** | Micro-entreprise, SASU, EURL, SAS, SARL — tous couverts |
| 📅 **Réglementation 2026** | ACRE, ARE, ARCE, guichet unique INPI intégrés |
| 🎨 **Design moderne** | Interface CSS personnalisée, responsive, claire et agréable |

---

## 📄 Document Word généré

Le fichier `.docx` contient :

1. **Page de garde** — Nom du projet, porteur, statut, date
2. **Executive Summary** — Synthèse en un coup d'œil
3. **L'Idée & le Problème** — Problème, solution, différenciateur
4. **Étude de Marché** — Taille, concurrents, stratégie d'acquisition
5. **Produit / Service** — Description, modèle économique, marges
6. **L'Équipe** — Profils des fondateurs, compétences, recrutements
7. **Stratégie Commerciale** — Go-to-market, motivation
8. **Plan de Financement** — Apports, emprunts, aides
9. **Projections Financières** — Tableau CA / Charges / Résultat sur 3 ans
10. **Statut Juridique** — Structure, capital, domiciliation

---

## 📊 Fichier Excel généré

| Onglet | Contenu |
|---|---|
| 📊 Dashboard | KPIs clés + infos projet en un seul coup d'œil |
| 📈 P&L 3 Ans | Compte de résultat prévisionnel + graphique à barres |
| 💸 Charges Détail | Répartition mensuelle/annuelle + camembert |
| 💰 Financement | Sources de financement, solde disponible |
| 📅 CA Mensuel An1 | Projection mois par mois + courbe d'évolution |

---

## 🛠️ Installation & Lancement

### Prérequis

- Python 3.9+
- pip

### Étapes

```bash
# 1. Cloner le projet
git clone https://github.com/votre-utilisateur/business-plan-app.git
cd business-plan-app

# 2. Installer les dépendances
pip install streamlit python-docx openpyxl

# 3. Lancer l'application
streamlit run business_plan_app.py
```

L'app s'ouvre automatiquement sur `http://localhost:8501`

---

## 📦 Dépendances

```txt
streamlit
python-docx
openpyxl
```

---

## 🗂️ Structure du projet

```
business-plan-app/
│
├── business_plan_app.py   # Application principale (questionnaire + générateurs)
├── README.md              # Documentation
└── requirements.txt       # Dépendances Python
```

---

## ⚖️ Statuts juridiques couverts

L'application vous aide à choisir parmi :

- **Micro-entreprise** — Idéale pour démarrer seul, simplement
- **SASU** — Société unipersonnelle flexible, régime assimilé-salarié
- **EURL** — Société unipersonnelle, charges sociales optimisées (SSI)
- **SAS** — Structure collective souple, idéale pour lever des fonds
- **SARL** — Structure collective encadrée, idéale pour projets familiaux

> Inclut les mises à jour 2026 : ACRE plafonné à 25%, nouvelles règles ARE/ARCE depuis avril 2025, guichet unique INPI obligatoire.

---

## 💡 Cas d'usage

- 🧑‍💼 Porteur de projet qui formalise son idée pour la première fois
- 🎓 Étudiant en entrepreneuriat préparant un dossier
- 💼 Freelance souhaitant créer sa société
- 🏦 Entrepreneur cherchant à convaincre une banque ou un investisseur
- 👩‍🍳 Artisan, commerçant ou profession libérale qui se lance

---

## 📝 Licence

MIT — Libre d'utilisation, de modification et de distribution.

---

## 🙏 Sources

Ce projet s'appuie sur le contenu du **Guide Pratique Création d'Entreprise — Édition 2026**, rédigé par l'équipe juridique de [Legalstart](https://www.legalstart.fr).

---

<p align="center">
  Fait avec ❤️ pour simplifier l'entrepreneuriat
</p>
