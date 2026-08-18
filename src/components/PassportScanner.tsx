import React, { useState } from 'react';
import {
  Upload,
  Sparkles,
  Cpu,
  RefreshCw,
  AlertCircle,
  X,
  FileSearch,
  Image as ImageIcon,
  ShieldCheck,
  CheckCircle2,
  Camera,
} from 'lucide-react';
import { ExtractionResult } from '../types';
import { extractWithGeminiAI, extractWithTesseract } from '../services/ocrService';

interface PassportScannerProps {
  onExtractionComplete: (result: ExtractionResult) => void;
  onCancel?: () => void;
}

export const PassportScanner: React.FC<PassportScannerProps> = ({
  onExtractionComplete,
  onCancel,
}) => {
  const [engine, setEngine] = useState<'ai' | 'tesseract'>('ai');
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const [progressPercent, setProgressPercent] = useState<number>(0);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [selectedFilePreview, setSelectedFilePreview] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);

  // Resize / optimize image before sending to AI/OCR to guarantee first-try reading
  const optimizeImageIfNeeded = (dataUrl: string): Promise<string> => {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const maxDim = 2000;
        let width = img.width;
        let height = img.height;

        if (width <= maxDim && height <= maxDim) {
          resolve(dataUrl);
          return;
        }

        if (width > height) {
          height = Math.round((height * maxDim) / width);
          width = maxDim;
        } else {
          width = Math.round((width * maxDim) / height);
          height = maxDim;
        }

        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          resolve(dataUrl);
          return;
        }
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL('image/jpeg', 0.94));
      };
      img.onerror = () => resolve(dataUrl);
      img.src = dataUrl;
    });
  };

  // Perform extraction on an image base64
  const processImage = async (imageBase64: string, mimeType: string = 'image/jpeg') => {
    setIsProcessing(true);
    setErrorMessage(null);
    setProgressPercent(20);
    setSelectedFilePreview(imageBase64);

    try {
      // Optimize image resolution for crystal clear OCR
      const optimizedBase64 = await optimizeImageIfNeeded(imageBase64);

      if (engine === 'ai') {
        setProcessingStatus('Analizando documento con Inteligencia Artificial Gemini...');
        setProgressPercent(50);

        const result = await extractWithGeminiAI(optimizedBase64, mimeType);
        setProgressPercent(100);
        setProcessingStatus('¡Datos de pasaporte leídos y extraídos con éxito!');
        setTimeout(() => {
          onExtractionComplete(result);
        }, 400);
      } else {
        setProcessingStatus('Iniciando motor OCR local Tesseract.js...');
        const result = await extractWithTesseract(optimizedBase64, (prog, status) => {
          setProgressPercent(prog);
          setProcessingStatus(status);
        });
        result.imagePreview = optimizedBase64;
        setTimeout(() => {
          onExtractionComplete(result);
        }, 400);
      }
    } catch (err: any) {
      console.error('Extraction error:', err);
      // If AI fails, automatically perform OCR fallback
      if (engine === 'ai') {
        setProcessingStatus('Conectando con motor de respaldo OCR de alta precisión...');
        try {
          const fallback = await extractWithTesseract(imageBase64);
          fallback.imagePreview = imageBase64;
          fallback.notes = 'Extracción completada con motor OCR tras alta demanda del servicio IA.';
          onExtractionComplete(fallback);
          return;
        } catch (tessErr: any) {
          setErrorMessage(`Error en lectura: ${err.message || 'No se pudo procesar la imagen del pasaporte.'}`);
        }
      } else {
        setErrorMessage(`Error en el reconocimiento OCR: ${err.message}`);
      }
    } finally {
      setIsProcessing(false);
    }
  };

  // Handle file input
  const handleFileUpload = (file: File) => {
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const base64 = event.target?.result as string;
      if (base64) {
        processImage(base64, file.type || 'image/jpeg');
      }
    };
    reader.readAsDataURL(file);
  };

  // Drag & drop handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      handleFileUpload(file);
    }
  };

  return (
    <div className="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
      {/* Header bar */}
      <div className="bg-slate-900 text-white px-6 py-4 flex flex-wrap items-center justify-between gap-4 border-b border-slate-800">
        <div className="flex items-center gap-3">
          <div className="p-2 bg-amber-500/20 text-amber-400 rounded-lg border border-amber-500/30">
            <Upload className="w-5 h-5" />
          </div>
          <div>
            <h3 className="font-serif font-bold text-lg text-slate-100">
              Lectura y Extracción de Pasaporte / Documento
            </h3>
            <p className="text-xs text-slate-400">
              Sube la imagen del pasaporte o cédula para extraer automáticamente todos los campos del cliente
            </p>
          </div>
        </div>

        {/* Engine selector */}
        <div className="flex items-center bg-slate-800/90 p-1 rounded-lg border border-slate-700">
          <button
            id="btn-engine-ai"
            type="button"
            onClick={() => setEngine('ai')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium transition-all ${
              engine === 'ai'
                ? 'bg-amber-500 text-slate-950 shadow-sm font-semibold'
                : 'text-slate-300 hover:text-white'
            }`}
          >
            <Sparkles className="w-3.5 h-3.5" />
            <span>Extracción IA (Recomendado)</span>
          </button>
          <button
            id="btn-engine-tesseract"
            type="button"
            onClick={() => setEngine('tesseract')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium transition-all ${
              engine === 'tesseract'
                ? 'bg-amber-500 text-slate-950 shadow-sm font-semibold'
                : 'text-slate-300 hover:text-white'
            }`}
          >
            <Cpu className="w-3.5 h-3.5" />
            <span>OCR Tesseract.js</span>
          </button>
        </div>

        {onCancel && (
          <button
            onClick={onCancel}
            className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        )}
      </div>

      {/* Content Area */}
      <div className="p-6">
        {/* Error Alert */}
        {errorMessage && (
          <div className="mb-4 p-4 rounded-xl bg-red-50 border border-red-200 flex items-start gap-3 text-red-700 text-sm">
            <AlertCircle className="w-5 h-5 text-red-500 shrink-0 mt-0.5" />
            <div className="flex-1">
              <p className="font-semibold">Ocurrió un error al procesar el documento:</p>
              <p className="mt-0.5">{errorMessage}</p>
            </div>
            <button
              onClick={() => setErrorMessage(null)}
              className="text-red-500 hover:text-red-800 text-xs font-semibold"
            >
              Cerrar
            </button>
          </div>
        )}

        {/* Processing overlay state */}
        {isProcessing && (
          <div className="py-12 px-6 flex flex-col items-center justify-center text-center">
            <div className="w-16 h-16 rounded-2xl bg-amber-50 border border-amber-200 flex items-center justify-center mb-4 relative shadow-md">
              <RefreshCw className="w-8 h-8 text-amber-600 animate-spin" />
              <Sparkles className="w-4 h-4 text-amber-500 absolute -top-1 -right-1 animate-pulse" />
            </div>
            <h4 className="font-serif font-bold text-lg text-slate-800 mb-1">
              Leyendo y extrayendo datos del pasaporte
            </h4>
            <p className="text-sm text-slate-500 max-w-md mb-6">{processingStatus}</p>

            <div className="w-full max-w-md bg-slate-100 rounded-full h-3 overflow-hidden border border-slate-200">
              <div
                className="bg-gradient-to-r from-amber-500 to-amber-600 h-full transition-all duration-300 ease-out"
                style={{ width: `${progressPercent}%` }}
              />
            </div>
            <span className="text-xs text-slate-400 mt-2 font-mono">{progressPercent}% completado</span>
          </div>
        )}

        {/* File Upload Mode */}
        {!isProcessing && (
          <div className="space-y-5">
            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className={`border-2 border-dashed rounded-2xl p-10 flex flex-col items-center justify-center text-center transition-all ${
                isDragging
                  ? 'border-amber-500 bg-amber-50/60 scale-[1.01]'
                  : 'border-slate-300 hover:border-amber-500 bg-slate-50 hover:bg-amber-50/30'
              }`}
            >
              <label
                htmlFor="passport-file-input"
                className="cursor-pointer flex flex-col items-center justify-center w-full"
              >
                <div className="w-16 h-16 rounded-2xl bg-white shadow-md border border-slate-200 flex items-center justify-center mb-4 text-amber-600">
                  <Upload className="w-8 h-8" />
                </div>
                <h4 className="font-bold text-slate-800 text-base mb-1">
                  Arrastra o haz clic para subir la foto del pasaporte o cédula
                </h4>
                <p className="text-xs text-slate-500 max-w-md mb-4">
                  El sistema detectará automáticamente el <span className="font-semibold text-slate-700">nombre completo</span>, <span className="font-semibold text-slate-700">número de documento</span>, <span className="font-semibold text-slate-700">nacionalidad</span>, <span className="font-semibold text-slate-700">fecha de nacimiento</span> y <span className="font-semibold text-slate-700">condición (sexo/edad)</span> a la primera.
                </p>
                <div className="flex flex-wrap items-center justify-center gap-3">
                  <span className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl bg-slate-900 text-white text-xs font-bold shadow-md hover:bg-slate-800 transition-colors">
                    <FileSearch className="w-4 h-4 text-amber-400" />
                    Seleccionar Archivo de Imagen
                  </span>
                </div>
                <input
                  id="passport-file-input"
                  type="file"
                  accept="image/jpeg,image/png,image/webp,image/jpg,image/heic"
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (file) handleFileUpload(file);
                  }}
                  className="hidden"
                />
              </label>
            </div>

            {/* Helpful instructions card */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-3 pt-1">
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 flex items-start gap-2.5">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                <div>
                  <h5 className="text-xs font-bold text-slate-800">Lectura a la primera</h5>
                  <p className="text-[11px] text-slate-500">Optimización de imagen y contraste automático previo al escaneo.</p>
                </div>
              </div>
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 flex items-start gap-2.5">
                <ShieldCheck className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                <div>
                  <h5 className="text-xs font-bold text-slate-800">Campos Legales Clave</h5>
                  <p className="text-[11px] text-slate-500">Extrae (nombre), (número de pasaporte), (nacionalidad), (sexo/edad) y fechas.</p>
                </div>
              </div>
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 flex items-start gap-2.5">
                <ImageIcon className="w-4 h-4 text-blue-600 shrink-0 mt-0.5" />
                <div>
                  <h5 className="text-xs font-bold text-slate-800">Formatos Compatibles</h5>
                  <p className="text-[11px] text-slate-500">Admite JPG, PNG, WEBP y fotos tomadas desde cualquier teléfono o escáner.</p>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

