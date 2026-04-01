import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { saveAs } from "file-saver";

export interface MissionData {
  Nom: string;
  Prenoms: string;
  De: string;
  Etape: string;
  A: string;
  "Le (départ)": string;
  "Le (retour)": string;
  [key: string]: string;
}

export function readExcelFile(file: File): Promise<MissionData[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json<MissionData>(sheet, { raw: false });
        resolve(jsonData);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export function readWordTemplate(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target!.result as ArrayBuffer);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function sanitizeFilename(str: string): string {
  return str.replace(/[^a-zA-Z0-9àâäéèêëïîôùûüÿçÀÂÄÉÈÊËÏÎÔÙÛÜŸÇ_-]/g, "_");
}

export async function generateMissionOrders(
  templateBuffer: ArrayBuffer,
  missions: MissionData[]
): Promise<void> {
  const zip = new JSZip();

  for (const mission of missions) {
    const pizzip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(pizzip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "{", end: "}" },
    });

    // Map Excel columns to template placeholders
    const data: Record<string, string> = {};
    for (const [key, value] of Object.entries(mission)) {
      data[key] = value ?? "";
    }

    doc.render(data);

    const output = doc.getZip().generate({ type: "uint8array" });

    const nom = sanitizeFilename(mission.Nom || "Inconnu");
    const prenoms = sanitizeFilename(mission.Prenoms || "");
    const date = sanitizeFilename(mission["Le (départ)"] || "");
    const filename = `Ordre_Mission_${nom}_${prenoms}_${date}.docx`;

    zip.file(filename, output);
  }

  const content = await zip.generateAsync({ type: "blob" });
  saveAs(content, "Ordres_de_Mission.zip");
}
