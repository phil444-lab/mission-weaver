# Guide des Placeholders - Mission Weaver

## 📋 Placeholders disponibles

L'application lit automatiquement ces informations de votre fichier Excel et les remplace dans votre modèle Word :

### 👤 Données personnelles
- `{Nom}` - Nom de famille
- `{Prenoms}` - Prénoms  
- `{Fonction}` - Fonction/poste
- `{Courriel}` ou `{Couriel}` ou `{Email}` - Adresse email
- `{Téléphone}` ou `{Telephone}` ou `{Tel}` - Numéro de téléphone

### 📅 Dates et heures
- `{Date départ}` - Date de départ de la mission
- `{Date arrivée}` ou `{Date retour}` - Date de retour de la mission
- `{Heure départ}` - Heure de départ (aller)
- `{Heure arrivée départ}` - Heure d'arrivée (aller)
- `{Heure départ retour}` ou `{Heure retour}` - Heure de départ (retour)
- `{Heure arrivée retour}` - Heure d'arrivée (retour)

### ✈️ Informations de vol
- `{Vols aller}` - Numéros de vol pour l'aller (ex: "(AF 804), (TK 1234)")
- `{Vols retour}` - Numéros de vol pour le retour (ex: "(AF 805), (TK 5678)")
- `{Tous les vols}` ou `{Vols}` ou `{Numéros de vol}` - Tous les numéros de vol (ex: "(AF 804), (AF 805), (TK 1234)")

### 🗺️ Itinéraire
- `{Itineraire}` - Itinéraire complet (ex: Paris, Cotonou, Dubai, Paris)
- `{Ville de départ}` ou `{Ville de depart}` ou `{De}` ou `{Départ}` - Ville de départ
- `{Ville d'arrivée}` ou `{Ville arrivée}` ou `{A}` ou `{Arrivée}` - Ville d'arrivée
- `{Escales aller}` - Itinéraire aller complet (ex: "Cotonou, Istanbul, Dubai")
- `{Escales retour}` - Itinéraire retour complet (ex: "Dubai, Istanbul, Cotonou")
- `{Itinéraire aller}` ou `{Itineraire aller}` - Trajet aller uniquement
- `{Itinéraire retour}` ou `{Itineraire retour}` - Trajet retour uniquement

## 📝 Comment modifier votre fichier Word

1. Ouvrez votre fichier Word dans Microsoft Word ou LibreOffice
2. Remplacez les valeurs statiques par les placeholders entre accolades `{}`

### Exemple de modification

**Avant :**
```
Nom: AGBO
Prénoms: Coffi Angelos
Fonction: Consultant
Email: agbo@example.com
Téléphone: +229 12 34 56 78

Date de départ: 24/09/2025 à 15h20
Date de retour: 01/10/2025 à 06h00

Itinéraire: Cotonou → Istanbul → Dubai → Istanbul → Cotonou
Vols: TK 1234, TK 5678, TK 9012, TK 3456
```

**Après :**
```
Nom: {Nom}
Prénoms: {Prenoms}
Fonction: {Fonction}
Email: {Courriel}
Téléphone: {Téléphone}

Date de départ: {Date départ} à {Heure départ}
Date de retour: {Date retour} à {Heure arrivée retour}

Itinéraire: {Itineraire}
Vols: {Tous les vols}
```

**Résultat généré :**
```
Nom: AGBO
Prénoms: Coffi Angelos
Fonction: Consultant
Email: agbo@example.com
Téléphone: +229 12 34 56 78

Date de départ: 24/09/2025 à 15h20
Date de retour: 01/10/2025 à 06h00

Itinéraire: Cotonou, Istanbul, Dubai, Istanbul, Cotonou
Vols: (TK 1234), (TK 5678), (TK 9012), (TK 3456)
```

## 📊 Structure du fichier Excel attendue

Votre fichier Excel doit contenir ces colonnes :

### Colonnes obligatoires
- **N°** - Numéro d'ordre
- **NOM** - Nom de famille
- **PRENOMS** - Prénoms

### Colonnes optionnelles (nouvelles fonctionnalités)
- **FONCTION** - Fonction/poste
- **COURIEL** - Adresse email
- **TELEPHONE** - Numéro de téléphone
- **Date Départ** - Date de départ
- **HEURE DEPART** - Heure de départ (aller)
- **HEURE ARRIVE** - Heure d'arrivée (aller)
- **Date Retour** - Date de retour
- **HEURE DEPART_1** - Heure de départ (retour)
- **HEURE ARRIVE_1** - Heure d'arrivée (retour)
- **N VOL 1** à **N VOL 6** - Numéros de vol (N VOL 6 peut être dans une colonne sans en-tête)
- **Intinéraire** + colonnes suivantes - Villes de l'itinéraire

## 🔄 Aliases disponibles

Ces placeholders alternatifs fonctionnent aussi :

### Dates
- `{Le (départ)}` = `{Date départ}`
- `{Le (retour)}` = `{Date arrivée}`
- `{Date de départ}` = `{Date départ}`
- `{Date de retour}` = `{Date arrivée}`
- `{Date d'arrivée}` = `{Date arrivée}`

### Lieux
- `{De}` = `{Ville de départ}`
- `{A}` = `{Ville d'arrivée}`
- `{Départ}` = `{Ville de départ}`
- `{Arrivée}` = `{Ville d'arrivée}`

### Contact
- `{Email}` = `{Courriel}`
- `{Tel}` = `{Téléphone}`

## ⚡ Fonctionnalités automatiques

L'application analyse automatiquement :

1. **L'itinéraire complet** à partir des colonnes après "Intinéraire"
2. **Les vols aller/retour** en séparant les numéros de vol
3. **Les dates Excel** (conversion automatique des numéros de série)
4. **Les escales** en analysant le trajet

## 🎯 Conseils d'utilisation

- Utilisez des placeholders simples comme `{Nom}` plutôt que des variantes
- Testez avec un petit échantillon avant de traiter tout le fichier
- Vérifiez que vos colonnes Excel correspondent aux placeholders utilisés
- Les placeholders sont sensibles à la casse : `{Nom}` ≠ `{nom}`

## 📁 Noms des fichiers générés

Les fichiers générés utilisent le numéro de la colonne "N°" :
- N° 001 → **OM001.docx**
- N° 025 → **OM025.docx**
- N° 28 → **OM28.docx**

Les fichiers conservent la numérotation exacte du fichier Excel.