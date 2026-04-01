import { useState } from "react";
import { FileUploadZone } from "@/components/FileUploadZone";
import { readExcelFile, readWordTemplate, generateMissionOrders, MissionData } from "@/lib/docGenerator";
import { Button } from "@/components/ui/button";
import { Download, Loader2, FileStack, AlertCircle } from "lucide-react";
import { toast } from "sonner";

const Index = () => {
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [missions, setMissions] = useState<MissionData[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [previewLoaded, setPreviewLoaded] = useState(false);

  const handleExcelSelect = async (file: File) => {
    setExcelFile(file);
    try {
      const data = await readExcelFile(file);
      setMissions(data);
      setPreviewLoaded(true);
      toast.success(`${data.length} mission(s) détectée(s) dans le fichier.`);
    } catch {
      toast.error("Erreur lors de la lecture du fichier Excel.");
      setMissions([]);
      setPreviewLoaded(false);
    }
  };

  const handleGenerate = async () => {
    if (!wordFile || missions.length === 0) return;
    setIsGenerating(true);
    try {
      const templateBuffer = await readWordTemplate(wordFile);
      await generateMissionOrders(templateBuffer, missions);
      toast.success(`${missions.length} ordre(s) de mission généré(s) avec succès !`);
    } catch (err) {
      console.error(err);
      toast.error("Erreur lors de la génération des documents.");
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
              <p className="text-sm text-muted-foreground mb-4">{missions.length} ordre(s) de mission à générer.</p>
              <div className="border rounded-xl overflow-hidden bg-card">
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="bg-muted/60">
                        <th className="text-left px-4 py-3 font-semibold text-muted-foreground">#</th>
                        {Object.keys(missions[0]).map((key) => (
                          <th key={key} className="text-left px-4 py-3 font-semibold text-muted-foreground whitespace-nowrap">{key}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {missions.slice(0, 10).map((m, i) => (
                        <tr key={i} className="border-t hover:bg-muted/30 transition-colors">
                          <td className="px-4 py-2.5 text-muted-foreground">{i + 1}</td>
                          {Object.values(m).map((v, j) => (
                            <td key={j} className="px-4 py-2.5 whitespace-nowrap">{v}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                {missions.length > 10 && (
                  <div className="px-4 py-2 text-sm text-muted-foreground border-t bg-muted/30">
                    … et {missions.length - 10} ligne(s) supplémentaire(s)
                  </div>
                )}
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
