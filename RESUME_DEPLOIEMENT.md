# 🚀 Résumé du déploiement sur Render

## ✅ Deux options disponibles

### Option 1: Static Site (RECOMMANDÉ - GRATUIT)
- **Type:** Static Site
- **Build Command:** `npm install && npm run build`
- **Publish Directory:** `dist`
- **Start Command:** (laisser vide)
- **Coût:** GRATUIT
- **Avantages:** Plus rapide (CDN), gratuit, simple

### Option 2: Web Service (avec serveur Node.js)
- **Type:** Web Service
- **Runtime:** Node
- **Build Command:** `npm install && npm run build`
- **Start Command:** `npm start`
- **Coût:** Gratuit jusqu'à 750h/mois, puis payant
- **Avantages:** Plus de contrôle, peut ajouter des APIs

## 📝 Instructions rapides

1. **Pousser sur GitHub:**
   ```bash
   git add .
   git commit -m "Configuration Render avec serveur Express"
   git push origin main
   ```

2. **Sur Render.com:**
   - Créer un compte (avec GitHub)
   - New + → Static Site (ou Web Service)
   - Sélectionner le dépôt `phil444-lab/mission-weaver`
   - Configurer selon l'option choisie ci-dessus
   - Créer et attendre le déploiement

3. **Accéder à l'app:**
   - URL: `https://mission-weaver.onrender.com`

## 📚 Documentation complète

- **INSTRUCTIONS_RENDER.txt** - Instructions pas à pas détaillées
- **DEPLOY_RENDER.md** - Documentation complète avec dépannage
- **server.js** - Serveur Express (pour Option 2)
- **render.yaml** - Configuration automatique

## ⚡ Commandes utiles

```bash
# Tester le build localement
npm run build

# Tester le serveur localement (Option 2)
npm start

# Voir le contenu du build
ls -la dist/
```

## 🎯 Recommandation

**Utilisez l'Option 1 (Static Site)** sauf si vous avez besoin d'un backend Node.js.
C'est gratuit, plus rapide, et suffisant pour cette application.
