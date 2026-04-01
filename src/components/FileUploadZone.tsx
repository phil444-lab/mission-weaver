import { useCallback, useState } from "react";
import { Upload, Check, FileText, Table } from "lucide-react";

interface FileUploadZoneProps {
  accept: string;
  label: string;
  description: string;
  icon: "word" | "excel";
  onFileSelect: (file: File) => void;
  selectedFile: File | null;
}

export function FileUploadZone({ accept, label, description, icon, onFileSelect, selectedFile }: FileUploadZoneProps) {
  const [isDragging, setIsDragging] = useState(false);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) onFileSelect(file);
  }, [onFileSelect]);

  const handleChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) onFileSelect(file);
  }, [onFileSelect]);

  const Icon = icon === "word" ? FileText : Table;

  return (
    <label
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={handleDrop}
      className={`
        relative flex flex-col items-center justify-center gap-3 p-8 rounded-xl border-2 border-dashed cursor-pointer transition-all duration-200
        ${isDragging ? "border-primary bg-primary/5 scale-[1.02]" : "border-border hover:border-primary/50 hover:bg-muted/50"}
        ${selectedFile ? "border-accent bg-accent/5" : ""}
      `}
    >
      <input type="file" accept={accept} onChange={handleChange} className="hidden" />
      
      <div className={`w-14 h-14 rounded-xl flex items-center justify-center transition-colors ${selectedFile ? "bg-accent/15 text-accent" : "bg-primary/10 text-primary"}`}>
        {selectedFile ? <Check className="w-7 h-7" /> : <Icon className="w-7 h-7" />}
      </div>
      
      <div className="text-center">
        <p className="font-semibold text-foreground">{label}</p>
        <p className="text-sm text-muted-foreground mt-1">{description}</p>
      </div>

      {selectedFile && (
        <div className="flex items-center gap-2 px-3 py-1.5 bg-accent/10 rounded-lg">
          <Upload className="w-3.5 h-3.5 text-accent" />
          <span className="text-sm font-medium text-accent">{selectedFile.name}</span>
        </div>
      )}
    </label>
  );
}
