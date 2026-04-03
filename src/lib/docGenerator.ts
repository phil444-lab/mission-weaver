import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { saveAs } from "file-saver";

export interface MissionData {
  Nom: string;
  Prenoms: string;
  Fonction?: string;
  Courriel?: string;
  Telephone?: string;
  "Date départ": string;
  "Heure départ"?: string;
  "Heure arrivée départ"?: string;
  "Date arrivée": string;
  "Heure départ retour"?: string;
  "Heure arrivée retour"?: string;
  Itineraire: string;
  "Vols aller"?: string;
  "Vols retour"?: string;
  "Tous les vols"?: string;
  [key: string]: string | undefined;
}

function formatDateToFrench(dateStr: string): string {
  if (!dateStr || dateStr.trim() === "") return "";
  
  // Si c'est déjà au format français (jj/mm/aaaa), le retourner tel quel
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateStr)) {
    return dateStr;
  }
  
  // Si c'est un nombre Excel (nombre de jours depuis 1900-01-01)
  const excelDateNumber = parseFloat(dateStr);
  if (!isNaN(excelDateNumber) && excelDateNumber > 1000) {
    // Convertir le nombre Excel en date (en UTC pour éviter les décalages)
    const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // Excel epoch en UTC
    const date = new Date(excelEpoch.getTime() + excelDateNumber * 86400000);
    
    if (!isNaN(date.getTime())) {
      const day = String(date.getUTCDate()).padStart(2, '0');
      const month = String(date.getUTCMonth() + 1).padStart(2, '0');
      const year = date.getUTCFullYear();
      return `${day}/${month}/${year}`;
    }
  }
  
  // Essayer de parser différents formats de date
  const date = new Date(dateStr);
  
  // Vérifier si la date est valide
  if (!isNaN(date.getTime())) {
    // Utiliser UTC pour éviter les problèmes de fuseau horaire
    const day = String(date.getUTCDate()).padStart(2, '0');
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const year = date.getUTCFullYear();
    return `${day}/${month}/${year}`;
  }
  
  // Si on ne peut pas parser, retourner la valeur originale
  return dateStr;
}

function formatDateToFrenchLong(dateStr: string): string {
  if (!dateStr || dateStr.trim() === "") return "";
  
  const monthNames = [
    "janvier", "février", "mars", "avril", "mai", "juin",
    "juillet", "août", "septembre", "octobre", "novembre", "décembre"
  ];
  
  // Si c'est déjà au format français long, le retourner tel quel
  if (/^\d{1,2}\s+\w+\s+\d{4}$/.test(dateStr)) {
    return dateStr;
  }
  
  let date: Date;
  
  // Si c'est un nombre Excel (nombre de jours depuis 1900-01-01)
  const excelDateNumber = parseFloat(dateStr);
  if (!isNaN(excelDateNumber) && excelDateNumber > 1000) {
    // Convertir le nombre Excel en date (en UTC pour éviter les décalages)
    const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // Excel epoch en UTC
    date = new Date(excelEpoch.getTime() + excelDateNumber * 86400000);
  } else {
    // Essayer de parser différents formats de date
    date = new Date(dateStr);
  }
  
  // Vérifier si la date est valide
  if (!isNaN(date.getTime())) {
    const day = String(date.getUTCDate()).padStart(2, '0');
    const monthIndex = date.getUTCMonth();
    const year = date.getUTCFullYear();
    return `${day} ${monthNames[monthIndex]} ${year}`;
  }
  
  // Si on ne peut pas parser, retourner la valeur originale
  return dateStr;
}

export function readExcelFile(file: File): Promise<MissionData[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const workbook = XLSX.read(data, { 
          type: "array",
          cellDates: false,  // Ne pas convertir automatiquement en dates
          cellText: false,
          raw: true  // Lire les valeurs brutes
        });
        
        // Lire la première feuille
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        // Obtenir la plage de cellules
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
        
        // Lire les en-têtes de colonnes (première ligne)
        const headers: string[] = [];
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: C });
          const cell = sheet[cellAddress];
          headers.push(cell ? String(cell.v || '') : '');
        }
        
        console.log("En-têtes détectés:", headers);
        
        // Colonnes qui ne doivent PAS être converties en dates
        const nonDateColumns = new Set<number>();
        headers.forEach((header, index) => {
          const normalizedHeader = header.toLowerCase().trim();
          if (
            normalizedHeader.includes('nbre') ||
            normalizedHeader.includes('nombre') ||
            normalizedHeader.includes('n°') ||
            normalizedHeader.includes('numero') ||
            normalizedHeader.includes('numéro') ||
            normalizedHeader.includes('billet') ||
            normalizedHeader.includes('jours') ||
            normalizedHeader.includes('piece') ||
            normalizedHeader.includes('pièce')
          ) {
            nonDateColumns.add(index);
            console.log(`Colonne "${header}" (index ${index}) ne sera pas convertie en date`);
          }
        });
        
        // Convertir les cellules pour gérer les dates et les grands nombres
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {  // +1 pour sauter les en-têtes
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = sheet[cellAddress];
            
            if (!cell) continue;
            
            // Gérer les nombres qui peuvent être des dates ou des numéros
            if (cell.t === 'n' && cell.v) {
              const numValue = cell.v;
              const numStr = String(numValue);
              const columnIndex = C - range.s.c;
              
              // Si c'est un grand nombre (probablement un numéro de billet)
              if (numStr.length > 10 || numStr.includes('e') || numStr.includes('E')) {
                cell.t = 's';
                cell.v = cell.w || numStr.replace(/\..*$/, '');
              }
              // Si c'est une colonne qui ne doit pas être convertie en date
              else if (nonDateColumns.has(columnIndex)) {
                // Garder comme nombre ou convertir en texte si nécessaire
                if (cell.w) {
                  cell.t = 's';
                  cell.v = cell.w;
                }
              }
              // Si c'est un nombre entre 40000 et 60000 (dates Excel entre 2009 et 2064)
              else if (numValue >= 40000 && numValue <= 60000) {
                // Convertir le nombre Excel en date
                const daysOffset = numValue > 59 ? numValue - 1 : numValue;
                const millisecondsPerDay = 86400000;
                const excelEpoch = new Date(1900, 0, 1).getTime();
                const dateTime = excelEpoch + (daysOffset - 1) * millisecondsPerDay;
                const date = new Date(dateTime);
                
                const day = String(date.getDate()).padStart(2, '0');
                const month = String(date.getMonth() + 1).padStart(2, '0');
                const year = date.getFullYear();
                
                cell.t = 's';
                cell.v = `${day}/${month}/${year}`;
                cell.w = `${day}/${month}/${year}`;
                
                console.log(`Conversion date (colonne ${headers[columnIndex]}): ${numValue} -> ${day}/${month}/${year}`);
              }
            }
          }
        }
        
        // Convertir en JSON - la ligne 1 contient les en-têtes principaux
        // La ligne 2 contient des sous-en-têtes qu'on va ignorer
        const jsonData = XLSX.utils.sheet_to_json<any>(sheet, { 
          raw: false,
          defval: "",
          blankrows: false
        });
        
        console.log(`Excel lu: ${jsonData.length} lignes brutes détectées`);
        console.log("Colonnes détectées:", Object.keys(jsonData[0] || {}));
        console.log("Premières lignes brutes:", jsonData.slice(0, 3));
        
        // Filtrer et restructurer les données
        const filteredData = jsonData.filter((row, index) => {
          // Ignorer la ligne 2 qui contient "Départ" et "Arrivée" (sous-en-têtes)
          if (row['Date '] === 'Départ' || row['Date '] === 'Arrivée') {
            console.log(`Ligne ${index + 1} ignorée (sous-en-têtes):`, row);
            return false;
          }
          
          const hasData = row.NOM || row.PRENOMS;
          if (!hasData && index < 5) {
            console.log(`Ligne ${index + 1} filtrée (vide):`, row);
          }
          return hasData;
        }).map(row => {
          // Construire l'itinéraire à partir de TOUTES les colonnes après "Intinéraire"
          const itineraireSteps: string[] = [];
          const datePattern = /^\d{2}\/\d{2}\/\d{4}$/;
          
          console.log(`Toutes les colonnes pour ${row.NOM} ${row.PRENOMS}:`, Object.keys(row));
          console.log(`Valeur colonne Intinéraire:`, row['Intinéraire']);
          console.log(`Valeur __EMPTY_7:`, row['__EMPTY_7']);
          console.log(`Valeur __EMPTY_8:`, row['__EMPTY_8']);
          console.log(`Valeur __EMPTY_9:`, row['__EMPTY_9']);
          
          // Ajouter la colonne "Intinéraire" elle-même (première ville)
          if (row['Intinéraire'] && String(row['Intinéraire']).trim() !== '') {
            const val = String(row['Intinéraire']).trim();
            if (!datePattern.test(val)) {
              console.log(`  Intinéraire: ${val} (ajouté)`);
              itineraireSteps.push(val);
            } else {
              console.log(`  Intinéraire: ${val} (date, ignorée)`);
            }
          } else {
            console.log(`  Intinéraire: vide ou undefined`);
          }
          
          // Lire TOUTES les colonnes après "Intinéraire"
          const allKeys = Object.keys(row);
          const itineraireIndex = allKeys.indexOf('Intinéraire');
          
          console.log(`TOUTES les colonnes de la ligne:`, allKeys);
          console.log(`Index de la colonne Intinéraire: ${itineraireIndex}`);
          
          if (itineraireIndex >= 0) {
            // Prendre toutes les colonnes après "Intinéraire" (sans inclure "Intinéraire" elle-même)
            const colonnesApresItineraire = allKeys.slice(itineraireIndex + 1);
            console.log(`Colonnes après Intinéraire (${colonnesApresItineraire.length}):`, colonnesApresItineraire);
            
            // Colonnes à ignorer (ne font pas partie de l'itinéraire)
            // Note: "Ordre de mission" peut contenir la dernière ville de l'itinéraire
            const colonnesAIgnorer = new Set([
              'Nom et Prénoms',
              'N°',
              'Date OM',
              'N° billet',
              'NOM',
              'PRENOMS',
              'Nbre de jours',
              'Frais de mission',
              'N° Pièce',
              'Retrouvé'
            ]);
            
            // Parcourir toutes les colonnes après "Intinéraire" dans l'ordre
            colonnesApresItineraire.forEach(key => {
              const val = row[key];
              const valStr = val ? String(val).trim() : '';
              const isEmpty = !val || valStr === '';
              const isSubHeader = valStr === 'Départ' || valStr === 'Arrivée' || valStr === 'Arrivée ';
              const isDate = datePattern.test(valStr);
              const isIgnored = colonnesAIgnorer.has(key);
              
              console.log(`  ${key}: "${valStr}" (vide: ${isEmpty}, sous-en-tête: ${isSubHeader}, date: ${isDate}, ignoré: ${isIgnored})`);
              
              // Ajouter la valeur si elle est valide
              if (val && valStr !== '' && !isSubHeader && !isDate && !isIgnored) {
                itineraireSteps.push(valStr);
              }
            });
          }
          
          console.log(`Itinéraire brut pour ${row.NOM} ${row.PRENOMS} (${itineraireSteps.length} villes):`, itineraireSteps);
          
          // Construire l'itinéraire complet avec des tirets
          const itineraireComplet = itineraireSteps.join(" - ");
          const analyse = analyzeItineraire(itineraireComplet);
          
          console.log(`Analyse itinéraire pour ${row.NOM} ${row.PRENOMS}:`, {
            itineraireComplet,
            ...analyse
          });
          
          // Trouver la date d'arrivée (première colonne __EMPTY qui est une date)
          let dateArrivee = "";
          const emptyKeys = allKeys.filter(key => key === '__EMPTY' || key.startsWith('__EMPTY_'));
          emptyKeys.sort((a, b) => {
            if (a === '__EMPTY') return -1;
            if (b === '__EMPTY') return 1;
            const numA = parseInt(a.replace('__EMPTY_', ''));
            const numB = parseInt(b.replace('__EMPTY_', ''));
            return numA - numB;
          }).forEach(key => {
            const val = row[key];
            if (!dateArrivee && val && datePattern.test(String(val))) {
              dateArrivee = String(val);
              console.log(`Date d'arrivée trouvée dans ${key}: ${dateArrivee}`);
            }
          });
          
          // Extraire les numéros de vol (N VOL 1 à N VOL 6)
          const vols = [];
          for (let i = 1; i <= 6; i++) {
            const volKey = `N VOL ${i}`;
            if (row[volKey] && String(row[volKey]).trim() !== '') {
              vols.push(String(row[volKey]).trim());
            }
          }
          
          // Vérifier aussi __EMPTY qui pourrait être N VOL 6 (colonne T sans en-tête)
          if (row['__EMPTY'] && String(row['__EMPTY']).trim() !== '') {
            const emptyValue = String(row['__EMPTY']).trim();
            // Vérifier si c'est un numéro de vol (pattern: 2 lettres + espace + chiffres)
            const volPattern = /^[A-Z]{2}\s*\d+$/;
            if (volPattern.test(emptyValue)) {
              vols.push(emptyValue);
              console.log(`N VOL 6 trouvé dans __EMPTY: ${emptyValue}`);
            }
          }
          
          // Séparer les vols aller et retour (approximation: première moitié = aller, deuxième = retour)
          const milieu = Math.ceil(vols.length / 2);
          const volsAller = vols.slice(0, milieu);
          const volsRetour = vols.slice(milieu);
          
          // Formater les vols avec des parenthèses
          const formatVols = (volsList: string[]) => {
            return volsList.map(vol => `(${vol})`).join(', ');
          };
          
          const formattedRow: MissionData = {
            Nom: row.NOM || "",
            Prenoms: row.PRENOMS || "",
            Fonction: row.FONCTION || "",
            Courriel: row.COURIEL || "",
            Telephone: row.TELEPHONE || "",
            "Date OM": formatDateToFrenchLong(row['Date OM'] || ""),
            "Date départ": row['Date Départ'] || row['Date '] || "",
            "Heure départ": row['HEURE DEPART'] || "",
            "Heure arrivée départ": row['HEURE ARRIVE'] || "",
            "Date arrivée": row['Date Retour '] || dateArrivee,
            "Heure départ retour": row['HEURE DEPART_1'] || "",
            "Heure arrivée retour": row['HEURE ARRIVE_1'] || "",
            Itineraire: itineraireComplet,
            "Vols aller": formatVols(volsAller),
            "Vols retour": formatVols(volsRetour),
            "Tous les vols": formatVols(vols),
            // Ajouter les informations d'analyse
            "Ville de départ": analyse.villeDepart,
            "Ville d'arrivée": analyse.villeArrivee,
            "Escales aller": analyse.itineraireAller,
            "Escales retour": analyse.itineraireRetour,
            "Itinéraire aller": analyse.itineraireAller,
            "Itineraire aller": analyse.itineraireAller,
            "Itinéraire retour": analyse.itineraireRetour,
            "Itineraire retour": analyse.itineraireRetour,
            // Garder aussi les colonnes originales
            ...row
          };
          
          return formattedRow;
        });
        
        console.log(`Résultat: ${filteredData.length} lignes valides après filtrage`);
        if (filteredData.length > 0) {
          console.log("Exemple première ligne valide:", filteredData[0]);
          console.log("Exemple dernière ligne valide:", filteredData[filteredData.length - 1]);
        }
        
        resolve(filteredData);
      } catch (err) {
        console.error("Erreur lecture Excel:", err);
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function sanitizeFilename(str: string): string {
  return str.replace(/[^a-zA-Z0-9àâäéèêëïîôùûüÿçÀÂÄÉÈÊËÏÎÔÙÛÜŸÇ_-]/g, "_");
}

function analyzeItineraire(itineraire: string) {
  const villes = itineraire.split(' - ').map(v => v.trim()).filter(v => v);
  
  if (villes.length === 0) {
    return {
      villeDepart: "",
      villeArrivee: "",
      escalesAller: "",
      escalesRetour: "",
      itineraireAller: "",
      itineraireRetour: ""
    };
  }
  
  // La première ville est le départ
  const villeDepart = villes[0];
  
  // Trouver le point de retournement (quand on revient à la ville de départ)
  // On cherche la ville la plus éloignée avant de revenir
  let indexArrivee = villes.length - 1;
  
  // Si la dernière ville est la même que la première, c'est un aller-retour
  if (villes[villes.length - 1] === villeDepart) {
    // Trouver le milieu du trajet (destination finale)
    const milieu = Math.floor(villes.length / 2);
    indexArrivee = milieu;
  } else {
    // Sinon, chercher où on commence à revenir vers le départ
    for (let i = 1; i < villes.length; i++) {
      if (villes[i] === villeDepart) {
        indexArrivee = i - 1;
        break;
      }
    }
  }
  
  const villeArrivee = villes[indexArrivee];
  
  // Extraire les escales aller (entre départ et arrivée, sans compter départ et arrivée)
  const escalesAllerBrut = villes.slice(1, indexArrivee).filter(v => v !== villeDepart && v !== villeArrivee);
  // Éliminer les doublons
  const escalesAller = [...new Set(escalesAllerBrut)];
  
  // Extraire les escales retour (après l'arrivée, sans compter arrivée et retour au départ)
  const escalesRetourBrut = villes.slice(indexArrivee + 1, -1).filter(v => v !== villeDepart && v !== villeArrivee);
  // Éliminer les doublons
  const escalesRetour = [...new Set(escalesRetourBrut)];
  
  // Construire les itinéraires avec des tirets
  const itineraireAller = [villeDepart, ...escalesAller, villeArrivee].join(' - ');
  const itineraireRetour = [villeArrivee, ...escalesRetour, villeDepart].join(' - ');
  
  return {
    villeDepart,
    villeArrivee,
    escalesAller: escalesAller.join(', ') || "Aucune",
    escalesRetour: escalesRetour.join(', ') || "Aucune",
    itineraireAller,
    itineraireRetour
  };
}

export function readWordTemplate(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target!.result as ArrayBuffer);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function normalizeKey(key: string): string {
  // Normaliser les clés pour le mapping (enlever espaces, accents, mettre en minuscules)
  return key
    .toLowerCase()
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, ""); // Enlever les accents
}

function createSmartMapping(excelData: Record<string, any>): Record<string, string> {
  const mapped: Record<string, string> = {};
  
  // Copier toutes les colonnes Excel
  for (const [key, value] of Object.entries(excelData)) {
    mapped[key] = value ?? "";
  }
  
  // Créer des alias standards pour les placeholders Word
  mapped["Nom"] = excelData.Nom || excelData.NOM || "";
  mapped["Prenoms"] = excelData.Prenoms || excelData.PRENOMS || "";
  mapped["Fonction"] = excelData.Fonction || excelData.FONCTION || "";
  mapped["Courriel"] = excelData.Courriel || excelData.COURIEL || "";
  mapped["Couriel"] = excelData.Courriel || excelData.COURIEL || "";
  mapped["Email"] = excelData.Courriel || excelData.COURIEL || "";
  mapped["Téléphone"] = excelData.Telephone || excelData.TELEPHONE || excelData.Téléphone || "";
  mapped["Telephone"] = excelData.Telephone || excelData.TELEPHONE || excelData.Téléphone || "";
  mapped["Tel"] = excelData.Telephone || excelData.TELEPHONE || excelData.Téléphone || "";
  
  mapped["Date OM"] = formatDateToFrenchLong(excelData["Date OM"] || "");
  mapped["Date départ"] = excelData["Date départ"] || excelData["Date Départ"] || excelData["Date "] || "";
  mapped["Date arrivée"] = excelData["Date arrivée"] || excelData["Date Retour "] || excelData["__EMPTY"] || "";
  mapped["Date retour"] = excelData["Date arrivée"] || excelData["Date Retour "] || excelData["__EMPTY"] || "";
  
  mapped["Heure départ"] = excelData["Heure départ"] || excelData["HEURE DEPART"] || "";
  mapped["Heure arrivée départ"] = excelData["Heure arrivée départ"] || excelData["HEURE ARRIVE"] || "";
  mapped["Heure départ retour"] = excelData["Heure départ retour"] || excelData["HEURE DEPART_1"] || "";
  mapped["Heure arrivée retour"] = excelData["Heure arrivée retour"] || excelData["HEURE ARRIVE_1"] || "";
  mapped["Heure retour"] = excelData["Heure départ retour"] || excelData["HEURE DEPART_1"] || "";
  
  mapped["Vols aller"] = excelData["Vols aller"] || "";
  mapped["Vols retour"] = excelData["Vols retour"] || "";
  mapped["Tous les vols"] = excelData["Tous les vols"] || "";
  mapped["Vols"] = excelData["Tous les vols"] || "";
  mapped["Numéros de vol"] = excelData["Tous les vols"] || "";
  
  mapped["Itineraire"] = excelData.Itineraire || "";
  
  // Analyser l'itinéraire pour extraire les détails
  const itineraireAnalyse = analyzeItineraire(mapped["Itineraire"]);
  
  // Ajouter les informations d'itinéraire détaillées avec plusieurs aliases
  mapped["Ville de départ"] = itineraireAnalyse.villeDepart;
  mapped["Ville de depart"] = itineraireAnalyse.villeDepart;
  mapped["Ville d'arrivée"] = itineraireAnalyse.villeArrivee;
  mapped["Ville d'arrivee"] = itineraireAnalyse.villeArrivee;
  mapped["Ville arrivée"] = itineraireAnalyse.villeArrivee;
  mapped["Ville arrivee"] = itineraireAnalyse.villeArrivee;
  mapped["Escales aller"] = itineraireAnalyse.itineraireAller;
  mapped["Escales retour"] = itineraireAnalyse.itineraireRetour;
  mapped["Itinéraire aller"] = itineraireAnalyse.itineraireAller;
  mapped["Itineraire aller"] = itineraireAnalyse.itineraireAller;
  mapped["Itinéraire retour"] = itineraireAnalyse.itineraireRetour;
  mapped["Itineraire retour"] = itineraireAnalyse.itineraireRetour;
  
  // Aliases simples pour compatibilité
  mapped["De"] = itineraireAnalyse.villeDepart;
  mapped["A"] = itineraireAnalyse.villeArrivee;
  mapped["Départ"] = itineraireAnalyse.villeDepart;
  mapped["Depart"] = itineraireAnalyse.villeDepart;
  mapped["Arrivée"] = itineraireAnalyse.villeArrivee;
  mapped["Arrivee"] = itineraireAnalyse.villeArrivee;
  
  // Aliases pour les dates
  mapped["Le (départ)"] = mapped["Date départ"];
  mapped["Le (depart)"] = mapped["Date départ"];
  mapped["Le (retour)"] = mapped["Date arrivée"];
  mapped["Date de départ"] = mapped["Date départ"];
  mapped["Date de depart"] = mapped["Date départ"];
  mapped["Date de retour"] = mapped["Date arrivée"];
  mapped["Date d'arrivée"] = mapped["Date arrivée"];
  mapped["Date d'arrivee"] = mapped["Date arrivée"];
  
  console.log("Mapping créé:", {
    Nom: mapped.Nom,
    Prenoms: mapped.Prenoms,
    "Date départ": mapped["Date départ"],
    "Date arrivée": mapped["Date arrivée"],
    Itineraire: mapped.Itineraire,
    "Ville de départ": mapped["Ville de départ"],
    "Ville d'arrivée": mapped["Ville d'arrivée"],
    "De": mapped["De"],
    "A": mapped["A"],
    "Escales aller": mapped["Escales aller"],
    "Escales retour": mapped["Escales retour"]
  });
  
  return mapped;
}

export async function generateMissionOrders(
  templateBuffer: ArrayBuffer,
  missions: MissionData[]
): Promise<void> {
  const zip = new JSZip();
  let successCount = 0;
  let errorCount = 0;
  const errors: string[] = [];

  // Vérifier les placeholders dans le template
  try {
    const pizzipTest = new PizZip(templateBuffer);
    const docTest = new Docxtemplater(pizzipTest, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "{", end: "}" },
    });
    
    // Extraire le contenu XML pour détecter les placeholders
    const xmlContent = pizzipTest.files["word/document.xml"]?.asText() || "";
    const placeholders = xmlContent.match(/\{[^}]+\}/g) || [];
    const uniquePlaceholders = [...new Set(placeholders.map(p => p.replace(/[{}]/g, '')))];
    
    console.log("Placeholders détectés dans le Word:", uniquePlaceholders);
    
    if (uniquePlaceholders.length === 0) {
      throw new Error(
        "Aucun placeholder trouvé dans le fichier Word. " +
        "Veuillez modifier votre document Word pour ajouter des placeholders entre accolades, " +
        "par exemple : {Nom}, {Prenoms}, {De}, {A}, {Le (départ)}, {Le (retour)}. " +
        "Consultez le fichier GUIDE_PLACEHOLDERS.md pour plus d'informations."
      );
    }
    
    // Vérifier que les colonnes Excel correspondent aux placeholders
    const excelColumns = Object.keys(missions[0] || {});
    const missingColumns = uniquePlaceholders.filter(p => !excelColumns.includes(p) && p !== 'Itineraire');
    
    if (missingColumns.length > 0) {
      console.warn("Placeholders Word sans colonne Excel correspondante:", missingColumns);
      console.log("Colonnes disponibles dans Excel:", excelColumns);
    }
    
  } catch (err) {
    console.error("Erreur lors de l'analyse du template:", err);
    throw err;
  }

  for (let i = 0; i < missions.length; i++) {
    try {
      const mission = missions[i];
      const pizzip = new PizZip(templateBuffer);
      const doc = new Docxtemplater(pizzip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: { start: "{", end: "}" },
      });

      // Préparer les données pour le template avec mapping simplifié
      const smartData = createSmartMapping(mission);
      
      console.log(`Mission ${i + 1}/${missions.length}:`, smartData);

      doc.render(smartData);

      const output = doc.getZip().generate({ type: "uint8array" });

      // Utiliser le numéro de la colonne "N°" du fichier Excel
      const numeroExcel = smartData["N°"] || smartData.N || String(i + 1);
      const filename = `OM${numeroExcel}.docx`;

      console.log(`Génération: ${filename} pour ${smartData.Nom} ${smartData.Prenoms}`);

      zip.file(filename, output);
      successCount++;
      
    } catch (err) {
      errorCount++;
      const errorMsg = err instanceof Error ? err.message : String(err);
      errors.push(`Mission ${i + 1} (${missions[i].Nom} ${missions[i].Prenoms}): ${errorMsg}`);
      console.error(`Erreur génération mission ${i + 1}:`, err);
      console.error("Données de la mission:", missions[i]);
    }
  }

  if (successCount === 0) {
    const errorDetails = errors.length > 0 ? "\n\nDétails:\n" + errors.slice(0, 3).join("\n") : "";
    throw new Error(
      "Aucun document n'a pu être généré. " +
      "Vérifiez que les placeholders dans le Word (entre accolades {}) correspondent aux colonnes Excel." +
      errorDetails
    );
  }

  const content = await zip.generateAsync({ type: "blob" });
  saveAs(content, "Ordres_de_Mission.zip");
  
  console.log(`Génération terminée: ${successCount} succès, ${errorCount} erreurs`);
  
  if (errorCount > 0) {
    console.warn("Erreurs rencontrées:", errors);
  }
}
