import React, { useState, useMemo } from 'react';
import {
  CheckCircle,
  UserCheck,
  ArrowRight,
  RotateCcw,
  Sparkles,
  ShieldCheck,
  AlertTriangle,
  UserPlus,
  Save,
  Users,
  Info,
  Calendar,
} from 'lucide-react';
import { ExtractionResult, Client } from '../types';
import { determineSexAgeCategory, calculateAgeFromBirthDate } from '../services/docxService';

interface DataReviewStepProps {
  extraction: ExtractionResult;
  existingClients: Client[];
  onConfirmClient: (client: Client, nextAction: 'generate' | 'save_only') => void;
  onRescan: () => void;
}

export const DataReviewStep: React.FC<DataReviewStepProps> = ({
  extraction,
  existingClients,
  onConfirmClient,
  onRescan,
}) => {
  const initialCategory = useMemo(() => {
    return determineSexAgeCategory(extraction.birthDate, extraction.sex).category;
  }, [extraction.birthDate, extraction.sex]);

  const [formData, setFormData] = useState<Partial<Client>>({
    id: `cli-${Date.now()}`,
    firstName: extraction.firstName || '',
    lastName: extraction.lastName || '',
    fullName: extraction.fullName || `${extraction.firstName || ''} ${extraction.lastName || ''}`.trim(),
    passportNumber: extraction.passportNumber || '',
    docType: 'pasaporte',
    nationality: extraction.nationality || 'ESPAÑOLA',
    issuingCountry: extraction.issuingCountry || extraction.nationality || 'ESPAÑA',
    birthDate: extraction.birthDate || '',
    expiryDate: extraction.expiryDate || '',
    sex: extraction.sex || 'M',
    sexAgeCategory: extraction.sexAgeCategory || initialCategory,
    email: '',
    phone: '',
    address: '',
    city: 'Madrid',
    notes: extraction.notes || 'Datos verificados mediante escaneo de pasaporte.',
    passportImageBase64: extraction.imagePreview,
  });

  const [associateMode, setAssociateMode] = useState<'new' | 'existing'>('new');
  const [selectedExistingId, setSelectedExistingId] = useState<string>('');
  const [showSavedNotification, setShowSavedNotification] = useState(false);

  // Update field
  const handleChange = (field: keyof Client, value: any) => {
    setFormData((prev) => {
      const updated = { ...prev, [field]: value };
      if (field === 'firstName' || field === 'lastName') {
        const fn = field === 'firstName' ? value : prev.firstName || '';
        const ln = field === 'lastName' ? value : prev.lastName || '';
        updated.fullName = `${fn} ${ln}`.trim().toUpperCase();
      }
      return updated;
    });
  };

  // If user picks an existing client, fill in their extra details (phone, email, address) while keeping scanned passport data
  const handleSelectExisting = (clientId: string) => {
    setSelectedExistingId(clientId);
    const existing = existingClients.find((c) => c.id === clientId);
    if (existing) {
      setFormData((prev) => ({
        ...prev,
        id: existing.id,
        phone: existing.phone || prev.phone,
        email: existing.email || prev.email,
        address: existing.address || prev.address,
        city: existing.city || prev.city,
      }));
    }
  };

  const handleContinue = (nextAction: 'generate' | 'save_only') => {
    const ageCalculated = calculateAgeFromBirthDate(formData.birthDate);
    const categoryInfo = determineSexAgeCategory(formData.birthDate, formData.sex, formData.sexAgeCategory);

    const finalClient: Client = {
      id: associateMode === 'existing' && selectedExistingId ? selectedExistingId : formData.id || `cli-${Date.now()}`,
      firstName: (formData.firstName || '').toUpperCase(),
      lastName: (formData.lastName || '').toUpperCase(),
      fullName: (formData.fullName || `${formData.firstName || ''} ${formData.lastName || ''}`).trim().toUpperCase(),
      passportNumber: (formData.passportNumber || '').toUpperCase(),
      docType: formData.docType || 'pasaporte',
      nationality: (formData.nationality || '').toUpperCase(),
      issuingCountry: (formData.issuingCountry || '').toUpperCase(),
      birthDate: formData.birthDate || '',
      expiryDate: formData.expiryDate || '',
      sex: formData.sex || 'M',
      sexAgeCategory: categoryInfo.category,
      age: ageCalculated ?? undefined,
      email: formData.email || '',
      phone: formData.phone || '',
      address: formData.address || '',
      city: formData.city || 'Madrid',
      notes: formData.notes || '',
      passportImageBase64: formData.passportImageBase64,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      documentCount: 0,
    };

    onConfirmClient(finalClient, nextAction);
  };

  return (
    <div className="space-y-6">
      {/* Top Banner */}
      <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4 flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-lg bg-emerald-500 text-white flex items-center justify-center shadow-sm">
            <CheckCircle className="w-6 h-6" />
          </div>
          <div>
            <h4 className="font-serif font-bold text-slate-900 text-base">
              ¡Datos del Pasaporte Extraídos con Éxito!
            </h4>
            <p className="text-xs text-slate-600">
              Revisa y edita los campos extraídos antes de guardarlos o generar los documentos legales.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1.5 px-3 py-1 bg-white rounded-lg border border-emerald-200 text-xs font-semibold text-emerald-800">
            <ShieldCheck className="w-4 h-4 text-emerald-600" />
            <span>Confianza: {extraction.confidenceScore || 95}%</span>
          </div>
          <button
            type="button"
            onClick={onRescan}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 text-xs font-medium transition-colors"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Volver a Escanear</span>
          </button>
        </div>
      </div>

      {/* Main Split Layout */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Left Column: Scanned Document & Extraction Info */}
        <div className="lg:col-span-4 space-y-4">
          <div className="bg-slate-900 text-white rounded-xl p-4 border border-slate-800 shadow-md">
            <h5 className="font-semibold text-xs tracking-wider uppercase text-slate-400 mb-3 flex items-center gap-2">
              <Sparkles className="w-3.5 h-3.5 text-amber-400" />
              Documento Escaneado
            </h5>

            {/* Scanned Image Preview */}
            {extraction.imagePreview ? (
              <div className="w-full aspect-[1.4] rounded-lg overflow-hidden border border-slate-700 bg-slate-950 mb-3">
                <img
                  src={extraction.imagePreview}
                  alt="Pasaporte escaneado"
                  className="w-full h-full object-contain"
                />
              </div>
            ) : (
              <div className="w-full aspect-[1.4] rounded-lg border border-dashed border-slate-700 flex items-center justify-center text-xs text-slate-500 mb-3">
                Sin vista previa de imagen
              </div>
            )}

            {/* Extraction Metadata */}
            <div className="space-y-2 text-xs">
              <div className="flex justify-between py-1 border-b border-slate-800">
                <span className="text-slate-400">Método de Lectura:</span>
                <span className="font-medium text-amber-400 uppercase">
                  {extraction.method === 'ai' ? 'IA Gemini 3.7 Flash' : 'OCR Tesseract.js'}
                </span>
              </div>
              {extraction.mrzLine1 && (
                <div className="pt-1">
                  <span className="text-slate-400 block mb-1">Zona MRZ Detectada:</span>
                  <div className="p-2 bg-slate-950 rounded font-mono text-[10px] text-amber-300 break-all leading-tight border border-slate-800">
                    <div>{extraction.mrzLine1}</div>
                    {extraction.mrzLine2 && <div>{extraction.mrzLine2}</div>}
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Client Linking Card */}
          <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-sm space-y-3">
            <h5 className="font-semibold text-slate-800 text-xs tracking-wider uppercase flex items-center gap-2">
              <Users className="w-4 h-4 text-amber-600" />
              Asociación con Clientes
            </h5>

            <div className="space-y-2">
              <label className="flex items-center gap-2 p-2 rounded-lg border border-slate-200 hover:bg-slate-50 cursor-pointer text-xs">
                <input
                  type="radio"
                  name="assocMode"
                  checked={associateMode === 'new'}
                  onChange={() => setAssociateMode('new')}
                  className="text-amber-600 focus:ring-amber-500"
                />
                <span className="font-medium text-slate-800">Registrar como Nuevo Cliente</span>
              </label>

              <label className="flex items-center gap-2 p-2 rounded-lg border border-slate-200 hover:bg-slate-50 cursor-pointer text-xs">
                <input
                  type="radio"
                  name="assocMode"
                  checked={associateMode === 'existing'}
                  onChange={() => setAssociateMode('existing')}
                  className="text-amber-600 focus:ring-amber-500"
                />
                <span className="font-medium text-slate-800">Actualizar Cliente Existente ({existingClients.length})</span>
              </label>
            </div>

            {associateMode === 'existing' && (
              <div className="pt-2">
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Seleccionar Cliente del Directorio:
                </label>
                <select
                  value={selectedExistingId}
                  onChange={(e) => handleSelectExisting(e.target.value)}
                  className="w-full text-xs p-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500 bg-white"
                >
                  <option value="">-- Elige un cliente existente --</option>
                  {existingClients.map((cli) => (
                    <option key={cli.id} value={cli.id}>
                      {cli.fullName} ({cli.passportNumber || 'Sin pasaporte'})
                    </option>
                  ))}
                </select>
              </div>
            )}
          </div>
        </div>

        {/* Right Column: Editable Form */}
        <div className="lg:col-span-8 bg-white rounded-xl p-6 border border-slate-200 shadow-sm space-y-6">
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <div>
              <h4 className="font-serif font-bold text-slate-900 text-lg">
                Ficha del Cliente & Datos del Documento
              </h4>
              <p className="text-xs text-slate-500">
                Puedes ajustar manualmente cualquier dato antes de incorporarlo al expediente
              </p>
            </div>
            <span className="px-2.5 py-1 rounded-full text-xs font-semibold bg-amber-100 text-amber-800 border border-amber-200">
              Modo Edición
            </span>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            {/* Nombres */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Nombres de Pila *
              </label>
              <input
                type="text"
                value={formData.firstName || ''}
                onChange={(e) => handleChange('firstName', e.target.value)}
                className="w-full text-sm font-medium p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                placeholder="Ej. CARLOS ANDRÉS"
              />
            </div>

            {/* Apellidos */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Apellidos *
              </label>
              <input
                type="text"
                value={formData.lastName || ''}
                onChange={(e) => handleChange('lastName', e.target.value)}
                className="w-full text-sm font-medium p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                placeholder="Ej. RESTREPO GÓMEZ"
              />
            </div>

            {/* Nombre Completo */}
            <div className="sm:col-span-2">
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Nombre Completo (para encabezados y firmas) *
              </label>
              <input
                type="text"
                value={formData.fullName || ''}
                onChange={(e) => handleChange('fullName', e.target.value)}
                className="w-full text-sm font-bold text-slate-900 bg-slate-50 p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
              />
            </div>

            {/* Tipo de Documento */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Tipo de Documento (cédula/pasaporte) *
              </label>
              <select
                value={formData.docType || 'pasaporte'}
                onChange={(e) => handleChange('docType', e.target.value)}
                className="w-full text-sm font-semibold p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500 bg-white"
              >
                <option value="pasaporte">Pasaporte</option>
                <option value="cedula">Cédula de Identidad</option>
                <option value="dni">DNI / NIE</option>
                <option value="otro">Documento Nacional</option>
              </select>
            </div>

            {/* Pasaporte / Cédula Nº */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Nº de Pasaporte / Documento *
              </label>
              <input
                type="text"
                value={formData.passportNumber || ''}
                onChange={(e) => handleChange('passportNumber', e.target.value)}
                className="w-full font-mono text-sm font-bold text-red-700 p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                placeholder="Ej. PA1234567"
              />
            </div>

            {/* Nacionalidad */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Nacionalidad *
              </label>
              <input
                type="text"
                value={formData.nationality || ''}
                onChange={(e) => handleChange('nationality', e.target.value)}
                className="w-full text-sm font-medium p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                placeholder="Ej. ESPAÑOLA, MEXICANA, COLOMBIANA"
              />
            </div>

            {/* País Emisor */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                País Emisor de Pasaporte
              </label>
              <input
                type="text"
                value={formData.issuingCountry || ''}
                onChange={(e) => handleChange('issuingCountry', e.target.value)}
                className="w-full text-sm font-medium p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                placeholder="Ej. ESPAÑA"
              />
            </div>

            {/* Sexo */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Sexo / Género
              </label>
              <select
                value={formData.sex || 'M'}
                onChange={(e) => handleChange('sex', e.target.value)}
                className="w-full text-sm font-medium p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500 bg-white"
              >
                <option value="M">M - Masculino</option>
                <option value="F">F - Femenino</option>
                <option value="X">X - No binario / Otro</option>
              </select>
            </div>

            {/* Fecha Nacimiento */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Fecha de Nacimiento
              </label>
              <input
                type="date"
                value={formData.birthDate || ''}
                onChange={(e) => handleChange('birthDate', e.target.value)}
                className="w-full text-sm font-medium p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
              />
            </div>

            {/* Condición Dinámica (sexo/edad) - Cálculo Automático de Edad */}
            {(() => {
              const info = determineSexAgeCategory(formData.birthDate, formData.sex, formData.sexAgeCategory);
              const isMinorOrYouth = info.age !== null && info.age <= 17;
              return (
                <div className="sm:col-span-2 bg-gradient-to-r from-amber-50/70 to-orange-50/50 border border-amber-200 rounded-xl p-4 shadow-sm">
                  <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
                    <div className="flex items-center gap-2">
                      <div className="w-6 h-6 rounded-md bg-amber-500 text-slate-950 flex items-center justify-center font-bold text-xs">
                        ⚖️
                      </div>
                      <span className="text-xs font-bold uppercase tracking-wider text-slate-900">
                        Campo Notarial: (sexo/edad)
                      </span>
                    </div>

                    <div className="flex items-center gap-2">
                      <span className="text-xs text-slate-600">
                        {info.age !== null ? (
                          <span className="font-semibold text-slate-800">
                            Edad calculada: <span className="text-amber-700 font-bold">{info.age} años</span>
                          </span>
                        ) : (
                          <span className="italic text-slate-400">Ingresa fecha para calcular edad</span>
                        )}
                      </span>
                      <span className="px-2 py-0.5 rounded text-[11px] font-bold bg-amber-500 text-slate-950 uppercase tracking-wide">
                        Valor Word: {info.category}
                      </span>
                    </div>
                  </div>

                  <p className="text-xs text-slate-600 mb-3">
                    {info.explanation} <br />
                    <span className="text-[11px] text-slate-500">
                      Regla: Hasta 17 años se clasifica como <strong>JOVEN</strong>. De 18 años en adelante se clasifica como <strong>VARÓN</strong> o <strong>MUJER</strong> según el sexo. Se inserta en <strong>MAYÚSCULAS Y NEGRITA</strong> en el Word.
                    </span>
                  </p>

                  <div className="flex flex-wrap items-center gap-2">
                    <span className="text-xs font-semibold text-slate-700 mr-1">Opciones rápidas:</span>
                    {(['VARÓN', 'MUJER', 'JOVEN', 'MENOR'] as const).map((cat) => {
                      const isSelected = (formData.sexAgeCategory || info.category).toUpperCase() === cat;
                      return (
                        <button
                          key={cat}
                          type="button"
                          onClick={() => handleChange('sexAgeCategory', cat)}
                          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                            isSelected
                              ? 'bg-slate-900 text-amber-400 shadow-sm border border-slate-900'
                              : 'bg-white text-slate-700 border border-slate-300 hover:bg-slate-50'
                          }`}
                        >
                          {isSelected && '✓ '}
                          {cat}
                        </button>
                      );
                    })}
                  </div>
                </div>
              );
            })()}

            {/* Fecha Vencimiento */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Fecha de Vencimiento
              </label>
              <input
                type="date"
                value={formData.expiryDate || ''}
                onChange={(e) => handleChange('expiryDate', e.target.value)}
                className="w-full text-sm font-medium p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
              />
            </div>

            {/* Teléfono */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Teléfono de Contacto
              </label>
              <input
                type="text"
                value={formData.phone || ''}
                onChange={(e) => handleChange('phone', e.target.value)}
                className="w-full text-sm p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                placeholder="+34 600 000 000"
              />
            </div>

            {/* Correo Electrónico */}
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Correo Electrónico
              </label>
              <input
                type="email"
                value={formData.email || ''}
                onChange={(e) => handleChange('email', e.target.value)}
                className="w-full text-sm p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                placeholder="cliente@ejemplo.com"
              />
            </div>

            {/* Domicilio */}
            <div className="sm:col-span-2">
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Domicilio / Dirección Legal
              </label>
              <input
                type="text"
                value={formData.address || ''}
                onChange={(e) => handleChange('address', e.target.value)}
                className="w-full text-sm p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                placeholder="Calle, Número, Piso, Código Postal, Ciudad"
              />
            </div>

            {/* Notas */}
            <div className="sm:col-span-2">
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                Notas / Observaciones del Expediente
              </label>
              <textarea
                rows={2}
                value={formData.notes || ''}
                onChange={(e) => handleChange('notes', e.target.value)}
                className="w-full text-sm p-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-amber-500 focus:border-amber-500"
                placeholder="Detalles sobre el caso legal, requerimientos especiales o estado del trámite..."
              />
            </div>
          </div>

          {/* Action Footer Buttons */}
          <div className="border-t border-slate-200 pt-6 flex flex-col sm:flex-row items-center justify-between gap-4">
            <button
              type="button"
              id="btn-save-client-only"
              onClick={() => handleContinue('save_only')}
              className="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-5 py-2.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 font-semibold text-sm transition-colors"
            >
              <Save className="w-4 h-4 text-slate-600" />
              <span>Guardar en Directorio de Clientes</span>
            </button>

            <button
              type="button"
              id="btn-confirm-and-generate"
              onClick={() => handleContinue('generate')}
              className="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-6 py-3 rounded-xl bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-400 hover:to-amber-500 text-slate-950 font-bold text-sm shadow-lg shadow-amber-500/20 active:scale-95 transition-all"
            >
              <span>Continuar a Generar Documento Word</span>
              <ArrowRight className="w-4 h-4 text-slate-950" />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
