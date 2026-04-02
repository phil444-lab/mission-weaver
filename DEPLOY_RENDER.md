# Guide de déploiement sur Render

## Étapes pour déployer manuellement sur Render

### 1. Créer un compte Render
- Allez sur [render.com](https://render.com)
- Créez un compte ou connectez-vous avec GitHub

### 2. Créer un nouveau Static Site (OPTION 1 - Gratuit)
1. Cliquez sur "New +" en haut à droite
2. Sélectionnez "Static Site"
3. Connectez votre dépôt GitHub `phil444-lab/mission-weaver`

### 3a. Configuration Static Site (GRATUIT)

**Name:** `mission-weaver`

**Branch:** `main`

**Build Command:**
```bash
npm install && npm run build
```

**Publish Directory:**
```
dist
```

**Start Command:** (laisser vide - pas nécessaire pour Static Site)

**Auto-Deploy:** `Yes`

---

### 2. Alternative: Créer un Web Service (OPTION 2 - Payant après 750h/mois)
1. Cliquez sur "New +" en haut à droite
2. Sélectionnez "Web Service"
3. Connectez votre dépôt GitHub `phil444-lab/mission-weaver`

### 3b. Configuration Web Service

**Name:** `mission-weaver`

**Branch:** `main`

**Runtime:** `Node`

**Build Command:**
```bash
npm install && npm run build
```

**Start Command:**
```bash
npm start
```

**Auto-Deploy:** `Yes`

### 4. Variables d'environnement (optionnel)
Aucune variable d'environnement n'est nécessaire pour ce projet.

### 5. Déployer
- Cliquez sur "Create Static Site"
- Render va automatiquement:
  - Cloner votre dépôt
  - Installer les dépendances
  - Construire le projet avec Vite
  - Déployer le contenu du dossier `dist`

### 6. Accéder à votre application
Une fois le déploiement terminé, Render vous donnera une URL du type:
```
https://mission-weaver.onrender.com
```

## Configuration des redirections (SPA)

Le fichier `render.yaml` à la racine du projet configure automatiquement les redirections pour que votre application React Router fonctionne correctement.

Si vous préférez configurer manuellement:
1. Allez dans les paramètres de votre Static Site
2. Section "Redirects/Rewrites"
3. Ajoutez une règle:
   - **Source:** `/*`
   - **Destination:** `/index.html`
   - **Action:** `Rewrite`

## Mises à jour automatiques

Avec Auto-Deploy activé, chaque fois que vous poussez sur la branche `main`, Render redéploiera automatiquement votre application.

## Commandes utiles

### Tester le build localement avant de déployer:
```bash
npm run build
npm run preview
```

### Vérifier que le build fonctionne:
```bash
npm run build
ls -la dist/
```

Le dossier `dist/` doit contenir:
- `index.html`
- `assets/` (avec les fichiers JS et CSS)
- Autres fichiers statiques

## Dépannage

### Le build échoue sur Render
- Vérifiez que `package.json` et `package-lock.json` sont à jour dans Git
- Vérifiez les logs de build sur Render pour voir l'erreur exacte

### Les routes ne fonctionnent pas (404)
- Vérifiez que le fichier `render.yaml` est présent à la racine
- Ou configurez manuellement les redirections dans les paramètres Render

### L'application ne charge pas
- Vérifiez que le "Publish Directory" est bien `dist`
- Vérifiez que le build s'est terminé avec succès

## Ressources
- [Documentation Render - Static Sites](https://render.com/docs/static-sites)
- [Documentation Vite - Deploying](https://vitejs.dev/guide/static-deploy.html)
