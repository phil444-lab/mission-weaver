import { useState } from "react";
import { FileUploadZone } from "@/components/FileUploadZone";
import { readExcelFile, readWordTemplate, generateMissionOrders, MissionData } from "@/lib/docGenerator";
import { Button } from "@/components/ui/button";
import { Download, Loader2, FileStack, AlertCircle, Info } from "lucide-react";
import { toast } from "sonner";
import { Alert, AlertDescription } from "@/components/ui/alert";

const Index = () => {
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [missions, setMissions] = useState<MissionData[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [previewLoaded, setPreviewLoaded] = useState(false);

  const handleExcelSelect = async (file: File) => {
    setExcelFile(file);
    setPreviewLoaded(false);
    try {
      const data = await readExcelFile(file);
      console.log("Données Excel chargées:", data);
      setMissions(data);
      setPreviewLoaded(true);
      if (data.length === 0) {
        toast.warning("Aucune donnée trouvée dans le fichier Excel.");
      } else {
        toast.success(`${data.length} mission(s) détectée(s) dans le fichier.`);
      }
    } catch (error) {
      console.error("Erreur lecture Excel:", error);
      toast.error("Erreur lors de la lecture du fichier Excel.");
      setMissions([]);
      setPreviewLoaded(false);
    }
  };

  const handleGenerate = async () => {
    if (!wordFile || missions.length === 0) return;
    setIsGenerating(true);
    try {
      console.log("Début de génération avec les colonnes:", Object.keys(missions[0]));
      const templateBuffer = await readWordTemplate(wordFile);
      await generateMissionOrders(templateBuffer, missions);
      toast.success(`${missions.length} ordre(s) de mission généré(s) avec succès !`);
    } catch (err) {
      console.error("Erreur génération:", err);
      const errorMsg = err instanceof Error ? err.message : "Erreur lors de la génération des documents.";
      toast.error(errorMsg);
    } finally {
      setIsGenerating(false);
    }
  };

  const canGenerate = wordFile && missions.length > 0;

  return (
    <div className="min-h-screen flex flex-col">
      {/* Header */}
      <header className="border-b bg-card px-6 py-5">
        <div className="max-w-4xl mx-auto flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-primary flex items-center justify-center">
            <FileStack className="w-5 h-5 text-primary-foreground" />
          </div>
          <div>
            <h1 className="text-xl font-bold text-foreground">Générateur d'Ordres de Mission</h1>
            <p className="text-sm text-muted-foreground">Générez vos documents Word automatiquement</p>
          </div>
        </div>
      </header>

      {/* Main */}
      <main className="flex-1 px-6 py-10">
        <div className="max-w-4xl mx-auto space-y-8">
          {/* Step 1: Upload files */}
          <section>
            <h2 className="text-lg font-semibold text-foreground mb-1">1. Importez vos fichiers</h2>
            <p className="text-sm text-muted-foreground mb-4">Chargez le modèle Word et le tableau Excel contenant les données.</p>
            <div className="grid md:grid-cols-2 gap-4">
              <FileUploadZone
                accept=".docx"
                label="Modèle Word"
                description="Fichier .docx avec les champs {Nom}, {Prenoms}, etc."
                icon="word"
                onFileSelect={setWordFile}
                selectedFile={wordFile}
              />
              <FileUploadZone
                accept=".xlsx,.xls"
                label="Fichier Excel"
                description="Tableau avec colonnes : Nom, Prenoms, De, A, etc."
                icon="excel"
                onFileSelect={handleExcelSelect}
                selectedFile={excelFile}
              />
            </div>
          </section>

          {/* Step 2: Preview */}
          {previewLoaded && missions.length > 0 && (
            <section className="animate-in fade-in slide-in-from-bottom-4 duration-300">
              <h2 className="text-lg font-semibold text-foreground mb-1">2. Aperçu des données</h2>
              <p className="text-sm text-muted-foreground mb-4">
                {missions.length} ordre(s) de mission à générer
              </p>
              
              <Alert className="mb-4">
                <Info className="h-4 w-4" />
                <AlertDescription className="text-sm">
                  <strong>Placeholders disponibles pour votre fichier Word :</strong>
                  <div className="mt-2 font-mono text-xs bg-muted/50 p-2 rounded">
                    {"{Nom}"}, {"{Prenoms}"}, {"{Date départ}"}, {"{Date arrivée}"}, {"{Itineraire}"}
                  </div>
                  <div className="text-muted-foreground mt-2 text-xs">
                    💡 Vous pouvez aussi utiliser : {"{Le (départ)}"}, {"{Le (retour)}"}, {"{Date de départ}"}, {"{Date de retour}"}
                  </div>
                </AlertDescription>
              </Alert>

              <div className="border rounded-xl overflow-hidden bg-card">
                <div className="overflow-x-auto max-h-[500px] overflow-y-auto">
                  <table className="w-full text-sm">
                    <thead className="sticky top-0 bg-muted/80 backdrop-blur-sm z-10">
                      <tr>
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground border-b">#</th>
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground border-b">Nom</th>
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground border-b">Prénoms</th>
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground border-b">Date départ</th>
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground border-b">Date arrivée</th>
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground border-b">De</th>
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground border-b">À</th>
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground border-b">Escales aller</th>
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground border-b">Escales retour</th>
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground border-b">Itinéraire</th>
                      </tr>
                    </thead>
                    <tbody>
                      {missions.map((m, i) => (
                        <tr key={i} className="border-t hover:bg-muted/30 transition-colors">
                          <td className="px-4 py-2.5 text-muted-foreground font-medium">{i + 1}</td>
                          <td className="px-4 py-2.5">{m.Nom || <span className="text-muted-foreground italic">vide</span>}</td>
                          <td className="px-4 py-2.5">{m.Prenoms || <span className="text-muted-foreground italic">vide</span>}</td>
                          <td className="px-4 py-2.5 whitespace-nowrap">{m["Date départ"] || <span className="text-muted-foreground italic">vide</span>}</td>
                          <td className="px-4 py-2.5 whitespace-nowrap">{m["Date arrivée"] || <span className="text-muted-foreground italic">vide</span>}</td>
                          <td className="px-4 py-2.5">{m["Ville de départ"] || <span className="text-muted-foreground italic">-</span>}</td>
                          <td className="px-4 py-2.5">{m["Ville d'arrivée"] || <span className="text-muted-foreground italic">-</span>}</td>
                          <td className="px-4 py-2.5">{m["Escales aller"] || <span className="text-muted-foreground italic">-</span>}</td>
                          <td className="px-4 py-2.5">{m["Escales retour"] || <span className="text-muted-foreground italic">-</span>}</td>
                          <td className="px-4 py-2.5">{m.Itineraire || <span className="text-muted-foreground italic">vide</span>}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="px-4 py-2 text-xs text-muted-foreground border-t bg-muted/30 flex justify-between items-center">
                  <span>Total: {missions.length} ligne(s)</span>
                  <span>Colonnes affichées: 10 (avec analyse d'itinéraire)</span>
                </div>
              </div>
            </section>
          )}

          {/* Step 3: Generate */}
          <section>
            <h2 className="text-lg font-semibold text-foreground mb-1">{previewLoaded ? "3." : "2."} Générer les documents</h2>
            <p className="text-sm text-muted-foreground mb-4">Les fichiers seront téléchargés dans un ZIP.</p>

            {!canGenerate && (
              <div className="flex items-center gap-2 text-sm text-muted-foreground bg-muted/50 rounded-lg px-4 py-3 mb-4">
                <AlertCircle className="w-4 h-4" />
                <span>Veuillez importer les deux fichiers pour continuer.</span>
              </div>
            )}

            <Button
              size="lg"
              disabled={!canGenerate || isGenerating}
              onClick={handleGenerate}
              className="gap-2"
            >
              {isGenerating ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin" />
                  Génération en cours…
                </>
              ) : (
                <>
                  <Download className="w-4 h-4" />
                  Générer {missions.length > 0 ? `${missions.length} ordre(s)` : "les ordres"} de mission
                </>
              )}
            </Button>
          </section>
        </div>
      </main>
    </div>
  );
};

export default Index;
