# Mission Weaver

Application web pour générer automatiquement des ordres de mission à partir de fichiers Excel.

## Fonctionnalités

- 📊 Import de fichiers Excel avec les données des missions
- 📄 Import de modèles Word avec placeholders
- 🔄 Génération automatique de documents Word personnalisés
- 📦 Export en fichier ZIP contenant tous les ordres de mission
- 🛣️ Analyse intelligente des itinéraires (aller-retour, escales)
- 📅 Conversion automatique des dates Excel

## Technologies

- React 18 + TypeScript
- Vite
- Tailwind CSS + shadcn/ui
- XLSX (lecture Excel)
- Docxtemplater (génération Word)
- React Router

## Installation locale

```bash
# Cloner le projet
git clone https://github.com/phil444-lab/mission-weaver.git
cd mission-weaver

# Installer les dépendances
npm install

# Lancer en développement
npm run dev

# Builder pour production
npm run build
```

## Déploiement

Voir [DEPLOY_RENDER.md](./DEPLOY_RENDER.md) pour les instructions de déploiement sur Render.

## Utilisation

1. Préparez votre fichier Excel avec les colonnes: NOM, PRENOMS, Date, Intinéraire, etc.
2. Préparez votre modèle Word avec des placeholders entre accolades: `{Nom}`, `{Prenoms}`, `{Ville de départ}`, etc.
3. Uploadez les deux fichiers dans l'application
4. Cliquez sur "Générer les ordres de mission"
5. Téléchargez le fichier ZIP contenant tous les documents

## Placeholders disponibles

Voir [GUIDE_PLACEHOLDERS.md](./GUIDE_PLACEHOLDERS.md) pour la liste complète des placeholders disponibles dans les modèles Word.

## Licence

MIT
