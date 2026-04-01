# Guide pour préparer votre fichier Word

## Placeholders disponibles

L'application lit automatiquement ces informations de votre fichier Excel :

### Données personnelles
- `{Nom}` - Nom de famille
- `{Prenoms}` - Prénoms

### Dates
- `{Date départ}` - Date de départ de la mission
- `{Date arrivée}` - Date d'arrivée/retour de la mission

### Itinéraire
- `{Itineraire}` - Itinéraire complet (ex: Paris → Cotonou → Paris)

## Comment modifier votre fichier Word

1. Ouvrez votre fichier Word dans Microsoft Word ou LibreOffice
2. Remplacez les valeurs statiques par les placeholders entre accolades `{}`

### Exemple de modification

**Avant :**
```
Nom: LICHEOU
Prénoms: Anthelme Silvère
Date de départ: 10/02/2024
Date de retour: 16/02/2024
Itinéraire: Cotonou - Paris - Casablanca
```

**Après :**
```
Nom: {Nom}
Prénoms: {Prenoms}
Date de départ: {Date départ}
Date de retour: {Date arrivée}
Itinéraire: {Itineraire}
```

## Aliases disponibles

Ces placeholders alternatifs fonctionnent aussi :
- `{Le (départ)}` = `{Date départ}`
- `{Le (retour)}` = `{Date arrivée}`
- `{Date de départ}` = `{Date départ}`
- `{Date de retour}` = `{Date arrivée}`
- `{Date d'arrivée}` = `{Date arrivée}`

## Structure du fichier Excel

Votre fichier Excel doit contenir ces colonnes :
- **NOM** - Nom de famille
- **PRENOMS** - Prénoms
- **Date** (colonne F) - Date de départ
- **Colonne G** - Date d'arrivée
- **Intinéraire** + colonnes suivantes - Villes de l'itinéraire

L'application construit automatiquement l'itinéraire à partir des villes listées.
