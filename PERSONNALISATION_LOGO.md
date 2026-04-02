# 🎨 Personnalisation du logo

## Logos actuels

L'application utilise maintenant:
- **Nom:** Mission Weaver
- **Logo:** Icône "MW" bleue avec un document
- **Favicon:** Version simplifiée du logo

## Fichiers de logo

### Logo principal
- **Fichier:** `public/logo.svg`
- **Taille:** 512x512px
- **Format:** SVG (vectoriel)
- **Utilisation:** Peut être affiché dans l'application

### Favicon
- **Fichier:** `public/favicon.svg`
- **Taille:** 32x32px
- **Format:** SVG (vectoriel)
- **Utilisation:** Icône dans l'onglet du navigateur

### Favicon alternatif
- **Fichier:** `public/favicon.ico`
- **Format:** ICO (bitmap)
- **Utilisation:** Fallback pour navigateurs anciens

## Comment personnaliser

### Option 1: Modifier les SVG existants
Éditez `public/logo.svg` ou `public/favicon.svg` avec:
- Un éditeur de texte (c'est du XML)
- Un logiciel comme Figma, Inkscape, ou Adobe Illustrator

### Option 2: Remplacer par vos propres fichiers
1. Créez votre logo (512x512px recommandé)
2. Exportez en SVG
3. Remplacez `public/logo.svg`
4. Créez une version 32x32px pour le favicon
5. Remplacez `public/favicon.svg`

### Option 3: Convertir en ICO
Pour créer un vrai favicon.ico:
1. Allez sur https://favicon.io/favicon-converter/
2. Uploadez votre image
3. Téléchargez le favicon.ico généré
4. Remplacez `public/favicon.ico`

## Couleurs actuelles

- **Bleu principal:** `#2563eb` (Tailwind blue-600)
- **Blanc:** `#ffffff`

Pour changer les couleurs, modifiez les valeurs `fill` dans les fichiers SVG.

## Tester localement

Après modification:
```bash
npm run dev
```

Puis ouvrez http://localhost:8080 et vérifiez:
- Le titre de l'onglet
- L'icône dans l'onglet (favicon)
- Le logo dans l'application (si affiché)

## Rebuild pour production

Après modification, n'oubliez pas de rebuilder:
```bash
npm run build
```

Les fichiers du dossier `public/` sont automatiquement copiés dans `dist/` lors du build.
