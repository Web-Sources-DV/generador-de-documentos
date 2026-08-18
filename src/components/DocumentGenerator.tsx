import React, { useState, useEffect } from 'react';
import {
  FileText,
  Download,
  CheckCircle2,
  Sparkles,
  Users,
  Calendar,
  AlertCircle,
  FileCheck,
  Tag,
  Eye,
  Sliders,
  RefreshCw,
  FolderOpen,
} from 'lucide-react';
import confetti from 'canvas-confetti';
import { Client, Template, GeneratedDocument, PlaceholderDef } from '../types';
import { generateAndDownloadDocx, determineSexAgeCategory } from '../services/docxService';
import { saveDocumentLog } from '../services/storageService';

interface DocumentGeneratorProps {
  clients: Client[];
  templates: Template[];
  selectedClient?: Client | null;
  onSelectClient?: (client: Client) => void;
  onDocumentGenerated?: (doc: GeneratedDocument) => void;
  onOpenTemplatesTab?: () => void;
}

export const DocumentGenerator: React.FC<DocumentGeneratorProps> = ({
  clients,
  templates,
  selectedClient: initialClient,
  onSelectClient,
  onDocumentGenerated,
  onOpenTemplatesTab,
}) => {
  const [selectedClientId, setSelectedClientId] = useState<string>(initialClient?.id || (clients[0]?.id ?? ''));
  const [selectedTemplateId, setSelectedTemplateId] = useState<string>(templates[0]?.id ?? '');
  const [templateCategoryFilter, setTemplateCategoryFilter] = useState<string>('all');
  const [formValues, setFormValues] = useState<Record<string, string>>({});
  const [customFileName, setCustomFileName] = useState<string>('');
  const [isGenerating, setIsGenerating] = useState<boolean>(false);
  const [lastGeneratedDoc, setLastGeneratedDoc] = useState<GeneratedDocument | null>(null);
  const [generationSuccess, setGenerationSuccess] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const activeClient = clients.find((c) => c.id === selectedClientId) || initialClient || null;
  const activeTemplate = templates.find((t) => t.id === selectedTemplateId) || templates[0] || null;

  // Sync selected client if prop changes
  useEffect(() => {
    if (initialClient?.id) {
      setSelectedClientId(initialClient.id);
    }
  }, [initialClient]);

  // Populate mapped form values whenever activeClient or activeTemplate changes
  useEffect(() => {
    if (!activeTemplate) return;

    const initialValues: Record<string, string> = {};

    activeTemplate.placeholders.forEach((placeholderKey) => {
      const cleanKey = placeholderKey.toLowerCase();

      // Find matching definition default if present
      const def = activeTemplate.placeholderDefs?.find((d) => d.key === placeholderKey);
      let val = def?.defaultValue || '';

      if (activeClient) {
        if (cleanKey.includes('cedula/pasaporte') || cleanKey.includes('cédula/pasaporte') || cleanKey.includes('cedula_pasaporte') || cleanKey.includes('tipo_documento') || cleanKey.includes('tipo de documento')) {
          val = activeClient.docType === 'cedula' ? 'Cédula' : activeClient.docType === 'dni' ? 'DNI' : 'Pasaporte';
        } else if (cleanKey.includes('nombre_completo') || cleanKey.includes('fullname') || cleanKey === 'cliente' || cleanKey === 'nombre' || cleanKey === '(nombre)' || cleanKey.includes('nombre')) {
          val = activeClient.fullName || `${activeClient.firstName} ${activeClient.lastName}`.trim();
        } else if (cleanKey.includes('nombres') || cleanKey.includes('firstname')) {
          val = activeClient.firstName || '';
        } else if (cleanKey.includes('apellidos') || cleanKey.includes('lastname')) {
          val = activeClient.lastName || '';
        } else if (cleanKey.includes('pasaporte') || cleanKey.includes('número de pasaporte') || cleanKey.includes('numero de pasaporte') || cleanKey.includes('documento') || cleanKey.includes('passport') || cleanKey.includes('dni') || cleanKey.includes('nie')) {
          val = activeClient.passportNumber || '';
        } else if (cleanKey.includes('nacionalidad') || cleanKey.includes('nationality')) {
          val = activeClient.nationality || '';
        } else if (cleanKey.includes('pais_emisor') || cleanKey.includes('pais') || cleanKey.includes('país')) {
          val = activeClient.issuingCountry || activeClient.nationality || '';
        } else if (cleanKey.includes('fecha_nacimiento') || cleanKey.includes('nacimiento') || cleanKey.includes('birth')) {
          val = activeClient.birthDate || '';
        } else if (cleanKey.includes('vencimiento') || cleanKey.includes('caducidad') || cleanKey.includes('expiry')) {
          val = activeClient.expiryDate || '';
        } else if (cleanKey.includes('sexo/edad') || cleanKey.includes('sexo_edad') || cleanKey.includes('sexo-edad') || cleanKey.includes('edad/sexo') || cleanKey.includes('condicion') || cleanKey.includes('condición')) {
          val = activeClient.sexAgeCategory || determineSexAgeCategory(activeClient.birthDate, activeClient.sex).category;
        } else if (cleanKey.includes('sexo') || cleanKey.includes('genero') || cleanKey.includes('gender')) {
          val = activeClient.sex === 'M' ? 'Masculino' : activeClient.sex === 'F' ? 'Femenino' : activeClient.sex;
        } else if (cleanKey.includes('telefono') || cleanKey.includes('phone') || cleanKey.includes('movil')) {
          val = activeClient.phone || '';
        } else if (cleanKey.includes('email') || cleanKey.includes('correo')) {
          val = activeClient.email || '';
        } else if (cleanKey.includes('direccion') || cleanKey.includes('dirección') || cleanKey.includes('domicilio') || cleanKey.includes('address')) {
          val = activeClient.address || '';
        } else if (cleanKey.includes('ciudad') || cleanKey.includes('city')) {
          val = activeClient.city || val || 'Madrid';
        }
      }

      // Date defaults
      if (!val && (cleanKey.includes('fecha_firma') || cleanKey.includes('fecha_solicitud') || cleanKey.includes('fecha_declaracion') || cleanKey === 'fecha')) {
        val = new Date().toLocaleDateString('es-ES', {
          day: 'numeric',
          month: 'long',
          year: 'numeric',
        });
      }

      if (!val && cleanKey.includes('ciudad')) {
        val = 'Madrid';
      }

      initialValues[placeholderKey] = val;
    });

    setFormValues(initialValues);

    // Default download filename
    if (activeTemplate && activeClient) {
      const cleanTplName = activeTemplate.name.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ]/g, '_');
      const cleanCliName = activeClient.fullName.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ]/g, '_');
      setCustomFileName(`${cleanTplName}_${cleanCliName}.docx`);
    } else if (activeTemplate) {
      setCustomFileName(`${activeTemplate.name.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ]/g, '_')}_Documento.docx`);
    }
  }, [activeTemplate, activeClient]);

  const handleFieldChange = (key: string, value: string) => {
    setFormValues((prev) => ({ ...prev, [key]: value }));
  };

  const handleGenerate = async () => {
    if (!activeTemplate) {
      setErrorMessage('Por favor selecciona una plantilla válida.');
      return;
    }

    setIsGenerating(true);
    setErrorMessage(null);
    setGenerationSuccess(false);

    try {
      const result = await generateAndDownloadDocx(activeTemplate, formValues, customFileName, activeClient);

      // Trigger celebration confetti
      try {
        confetti({
          particleCount: 80,
          spread: 70,
          origin: { y: 0.6 },
          colors: ['#f59e0b', '#d97706', '#10b981', '#3b82f6'],
        });
      } catch (cErr) {
        // Ignore confetti error
      }

      const newDoc: GeneratedDocument = {
        id: `doc-${Date.now()}`,
        title: `${activeTemplate.name} - ${activeClient ? activeClient.fullName : 'Cliente'}`,
        fileName: result.fileName,
        templateId: activeTemplate.id,
        templateName: activeTemplate.name,
        clientId: activeClient ? activeClient.id : 'sin-cliente',
        clientName: activeClient ? activeClient.fullName : 'CLIENTE DIRECTO',
        passportNumber: activeClient ? activeClient.passportNumber : '',
        generatedAt: new Date().toISOString(),
        fileSizeFormatted: result.sizeFormatted,
        dataSnapshot: { ...formValues },
      };

      saveDocumentLog(newDoc);
      setLastGeneratedDoc(newDoc);
      setGenerationSuccess(true);

      if (onDocumentGenerated) {
        onDocumentGenerated(newDoc);
      }
    } catch (err: any) {
      console.error('Error generating document:', err);
      setErrorMessage(`Error al generar el archivo Word: ${err.message || 'Comprueba los marcadores de la plantilla'}`);
    } finally {
      setIsGenerating(false);
    }
  };

  const filteredTemplates = templates.filter((t) => {
    if (templateCategoryFilter === 'all') return true;
    return t.category.toLowerCase() === templateCategoryFilter.toLowerCase();
  });

  const categories = Array.from(new Set(templates.map((t) => t.category)));

  return (
    <div className="space-y-6">
      {/* Top Banner / Selection bar */}
      <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* Client Selector */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="block text-xs font-bold uppercase tracking-wider text-slate-700 flex items-center gap-1.5">
                <Users className="w-4 h-4 text-amber-600" />
                1. Cliente Destinatario *
              </label>
              <span className="text-[11px] text-slate-500 font-medium">
                {clients.length} clientes en base de datos
              </span>
            </div>

            <select
              id="select-client-for-doc"
              value={selectedClientId}
              onChange={(e) => {
                setSelectedClientId(e.target.value);
                const cli = clients.find((c) => c.id === e.target.value);
                if (cli && onSelectClient) onSelectClient(cli);
              }}
              className="w-full text-sm font-semibold p-3 rounded-xl border border-slate-300 bg-white focus:ring-2 focus:ring-amber-500 focus:border-amber-500 shadow-sm"
            >
              {clients.map((cli) => (
                <option key={cli.id} value={cli.id}>
                  👤 {cli.fullName} | {cli.nationality} · Pasaporte: {cli.passportNumber || 'N/A'}
                </option>
              ))}
            </select>

            {activeClient && (
              <div className="mt-2.5 p-2.5 rounded-lg bg-amber-50/60 border border-amber-200/70 text-xs flex flex-wrap items-center justify-between gap-2 text-slate-700">
                <span>
                  <strong className="text-slate-900">Doc:</strong> {activeClient.passportNumber || 'S/N'} ({activeClient.nationality})
                </span>
                <span>
                  <strong className="text-slate-900">Nacimiento:</strong> {activeClient.birthDate || 'N/A'}
                </span>
                <span>
                  <strong className="text-slate-900">Caducidad:</strong> {activeClient.expiryDate || 'N/A'}
                </span>
              </div>
            )}
          </div>

            {/* Template Selector with Search & Filter */}
            <div>
              <div className="flex items-center justify-between mb-2">
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-700 flex items-center gap-1.5">
                  <FileText className="w-4 h-4 text-amber-600" />
                  2. Plantilla Word (.docx) *
                </label>
                {onOpenTemplatesTab && (
                  <button
                    type="button"
                    onClick={onOpenTemplatesTab}
                    className="text-xs text-amber-700 hover:text-amber-800 font-semibold underline"
                  >
                    + Subir / Gestionar Plantillas
                  </button>
                )}
              </div>

              <select
                id="select-template-for-doc"
                value={selectedTemplateId}
                onChange={(e) => setSelectedTemplateId(e.target.value)}
                className="w-full text-sm font-semibold p-3 rounded-xl border border-slate-300 bg-white focus:ring-2 focus:ring-amber-500 focus:border-amber-500 shadow-sm"
              >
                {templates.map((tpl) => (
                  <option key={tpl.id} value={tpl.id}>
                    📄 [{tpl.category}] {tpl.name} ({tpl.placeholders.length} marcadores)
                  </option>
                ))}
              </select>

              {activeTemplate && (
                <div className="mt-2.5 p-2.5 rounded-lg bg-slate-50 border border-slate-200 text-xs text-slate-600 flex items-center justify-between">
                  <span className="truncate max-w-xs">{activeTemplate.description}</span>
                  <span className="font-mono text-[11px] bg-slate-200 text-slate-800 px-2 py-0.5 rounded font-semibold shrink-0">
                    {activeTemplate.placeholders.length} tags
                  </span>
                </div>
              )}
            </div>
          </div>

          {/* Real-time Replacement Inspector Banner */}
          {activeClient && activeTemplate && (
            <div className="mt-6 pt-6 border-t border-slate-100">
              <div className="flex items-center justify-between mb-3">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-700 flex items-center gap-1.5">
                  <Sparkles className="w-3.5 h-3.5 text-amber-600" />
                  Vista Previa de Formato Notarial Automático
                </span>
                <span className="text-[11px] text-slate-500">
                  Valores que se insertarán en el archivo Word:
                </span>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-2 text-xs">
                <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-200">
                  <span className="text-[10px] text-slate-500 block font-mono">(nombre)</span>
                  <strong className="text-slate-900 font-bold block truncate">
                    {activeClient.fullName.toUpperCase()}
                  </strong>
                </div>
                <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-200">
                  <span className="text-[10px] text-slate-500 block font-mono">(número de pasaporte)</span>
                  <strong className="text-slate-900 font-bold block truncate">
                    {activeClient.passportNumber.toUpperCase()}
                  </strong>
                </div>
                <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-200">
                  <span className="text-[10px] text-slate-500 block font-mono">(nacionalidad) [minúsculas]</span>
                  <span className="text-slate-700 font-normal block truncate">
                    {activeClient.nationality.toLowerCase()}
                  </span>
                </div>
                <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-200">
                  <span className="text-[10px] text-slate-500 block font-mono">(cedula/pasaporte)</span>
                  <strong className="text-slate-900 font-bold block truncate">
                    {(activeClient.docType === 'cedula' ? 'CÉDULA' : 'PASAPORTE')}
                  </strong>
                </div>
                <div className="p-2.5 rounded-lg bg-amber-50/70 border border-amber-300">
                  <span className="text-[10px] text-amber-800 block font-mono font-bold">(sexo/edad) [notarial]</span>
                  <strong className="text-amber-950 font-extrabold block truncate">
                    {(activeClient.sexAgeCategory || determineSexAgeCategory(activeClient.birthDate, activeClient.sex).category).toUpperCase()}
                  </strong>
                </div>
              </div>
            </div>
          )}
        </div>

      {/* Success notification banner after generation */}
      {generationSuccess && lastGeneratedDoc && (
        <div className="bg-gradient-to-r from-emerald-600 to-teal-700 text-white rounded-2xl p-6 shadow-lg flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-white/20 flex items-center justify-center backdrop-blur-md">
              <FileCheck className="w-7 h-7 text-white" />
            </div>
            <div>
              <h4 className="font-serif font-bold text-lg text-white">
                ¡Documento Word (.docx) generado y descargado exitosamente!
              </h4>
              <p className="text-xs text-emerald-100 mt-0.5">
                Archivo: <span className="font-mono font-semibold">{lastGeneratedDoc.fileName}</span> ({lastGeneratedDoc.fileSizeFormatted})
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <button
              onClick={handleGenerate}
              className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl bg-white text-slate-900 font-bold text-xs shadow-md hover:bg-emerald-50 transition-colors"
            >
              <Download className="w-4 h-4 text-emerald-700" />
              <span>Descargar Nuevamente</span>
            </button>
          </div>
        </div>
      )}

      {/* Error alert */}
      {errorMessage && (
        <div className="p-4 rounded-xl bg-red-50 border border-red-200 flex items-center gap-3 text-red-700 text-sm">
          <AlertCircle className="w-5 h-5 text-red-500 shrink-0" />
          <p className="flex-1">{errorMessage}</p>
        </div>
      )}

      {/* Form Fields & Variables Mapping */}
      {activeTemplate && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 space-y-6">
          <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-100 pb-4">
            <div>
              <div className="flex items-center gap-2">
                <h3 className="font-serif font-bold text-slate-900 text-lg">
                  Variables y Marcadores Detectados en la Plantilla
                </h3>
                <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-amber-100 text-amber-800 border border-amber-200">
                  {activeTemplate.placeholders.length} campos
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Los datos del pasaporte y cliente se han mapeado automáticamente. Puedes ajustar cualquier valor antes de descargar.
              </p>
            </div>

            {/* Custom file name */}
            <div className="w-full sm:w-auto">
              <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">
                Nombre del Archivo Word de Salida:
              </label>
              <input
                type="text"
                value={customFileName}
                onChange={(e) => setCustomFileName(e.target.value)}
                className="text-xs font-mono p-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500 w-full sm:w-72 bg-slate-50"
                placeholder="nombre_archivo.docx"
              />
            </div>
          </div>

          {/* Placeholders Grid */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {activeTemplate.placeholders.map((key) => {
              const def = activeTemplate.placeholderDefs?.find((d) => d.key === key);
              const label = def?.label || key.replace(/_/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase());
              const isTextarea = def?.type === 'textarea' || key.includes('motivo') || key.includes('facultades') || key.includes('objeto');
              const isDate = def?.type === 'date' || key.includes('nacimiento') || key.includes('vencimiento');
              const isMappedFromClient = activeClient && (
                key.includes('nombre') ||
                key.includes('pasaporte') ||
                key.includes('nacionalidad') ||
                key.includes('telefono') ||
                key.includes('email') ||
                key.includes('direccion')
              );

              return (
                <div
                  key={key}
                  className={`p-3.5 rounded-xl border transition-all ${
                    isTextarea ? 'md:col-span-2 lg:col-span-3' : ''
                  } ${
                    isMappedFromClient
                      ? 'bg-amber-50/40 border-amber-200/80 shadow-xs'
                      : 'bg-slate-50/70 border-slate-200'
                  }`}
                >
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="block text-xs font-bold text-slate-800">
                      {label}
                    </label>
                    <span className="font-mono text-[10px] text-amber-800 bg-amber-100/80 px-1.5 py-0.2 rounded border border-amber-200">
                      {`{{${key}}}`}
                    </span>
                  </div>

                  {isTextarea ? (
                    <textarea
                      rows={3}
                      value={formValues[key] || ''}
                      onChange={(e) => handleFieldChange(key, e.target.value)}
                      className="w-full text-xs font-medium p-2.5 rounded-lg border border-slate-300 bg-white focus:ring-2 focus:ring-amber-500 focus:border-amber-500 leading-relaxed"
                      placeholder={`Ingresa el valor para {{${key}}}...`}
                    />
                  ) : isDate ? (
                    <input
                      type="date"
                      value={formValues[key] || ''}
                      onChange={(e) => handleFieldChange(key, e.target.value)}
                      className="w-full text-xs font-medium p-2 rounded-lg border border-slate-300 bg-white focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                    />
                  ) : (
                    <input
                      type="text"
                      value={formValues[key] || ''}
                      onChange={(e) => handleFieldChange(key, e.target.value)}
                      className="w-full text-xs font-medium p-2 rounded-lg border border-slate-300 bg-white focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                      placeholder={`Valor para {{${key}}}`}
                    />
                  )}
                </div>
              );
            })}
          </div>

          {/* Action Trigger Bar */}
          <div className="border-t border-slate-200 pt-6 flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="text-xs text-slate-500 flex items-center gap-2">
              <Sparkles className="w-4 h-4 text-amber-600 shrink-0" />
              <span>
                El motor compila un archivo DOCX nativo 100% compatible con Microsoft Word y Google Docs.
              </span>
            </div>

            <button
              id="btn-generate-docx"
              type="button"
              disabled={isGenerating}
              onClick={handleGenerate}
              className="w-full sm:w-auto inline-flex items-center justify-center gap-3 px-8 py-3.5 rounded-xl bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-400 hover:to-amber-500 text-slate-950 font-bold text-sm shadow-xl shadow-amber-500/25 active:scale-95 transition-all disabled:opacity-50"
            >
              {isGenerating ? (
                <>
                  <RefreshCw className="w-5 h-5 animate-spin text-slate-950" />
                  <span>Generando Documento...</span>
                </>
              ) : (
                <>
                  <Download className="w-5 h-5 text-slate-950" />
                  <span>Generar y Descargar Word (.docx)</span>
                </>
              )}
            </button>
          </div>
        </div>
      )}
    </div>
  );
};
