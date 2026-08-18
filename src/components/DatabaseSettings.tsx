import React, { useState } from 'react';
import {
  Database,
  Download,
  Upload,
  RotateCcw,
  HardDrive,
  ShieldCheck,
  CheckCircle2,
  AlertCircle,
  FileCode,
  FileSpreadsheet,
  Layers,
  Save,
} from 'lucide-react';
import saveAs from 'file-saver';
import { DatabaseStats } from '../types';
import {
  exportFullDatabaseJson,
  importDatabaseJson,
  resetDatabaseToDefaults,
} from '../services/storageService';

interface DatabaseSettingsProps {
  stats?: DatabaseStats;
  onDatabaseReload: () => void;
}

export const DatabaseSettings: React.FC<DatabaseSettingsProps> = ({
  stats = {
    totalClients: 0,
    totalTemplates: 0,
    totalGeneratedDocs: 0,
    storageUsageEstimateKb: 0,
  },
  onDatabaseReload,
}) => {
  const [feedbackMessage, setFeedbackMessage] = useState<{
    type: 'success' | 'error';
    text: string;
  } | null>(null);

  const handleExportBackup = () => {
    try {
      const json = exportFullDatabaseJson();
      const blob = new Blob([json], { type: 'application/json' });
      const dateStr = new Date().toISOString().split('T')[0];
      saveAs(blob, `SQP_Legal_Backup_Completo_${dateStr}.json`);
      setFeedbackMessage({
        type: 'success',
        text: '¡Copia de seguridad descargada exitosamente en formato JSON!',
      });
      onDatabaseReload();
    } catch (err: any) {
      setFeedbackMessage({
        type: 'error',
        text: `Error al exportar respaldo: ${err.message}`,
      });
    }
  };

  const handleImportBackup = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const text = event.target?.result as string;
      if (text) {
        const res = importDatabaseJson(text);
        if (res.success) {
          setFeedbackMessage({ type: 'success', text: res.message });
          onDatabaseReload();
        } else {
          setFeedbackMessage({ type: 'error', text: res.message });
        }
      }
    };
    reader.readAsText(file);
  };

  const handleReset = () => {
    if (
      confirm(
        '¿Deseas restablecer la base de datos a los valores y plantillas iniciales por defecto? Se mantendrán las plantillas legales oficiales.'
      )
    ) {
      resetDatabaseToDefaults();
      onDatabaseReload();
      setFeedbackMessage({
        type: 'success',
        text: 'La base de datos se ha restablecido a los valores iniciales predeterminados.',
      });
    }
  };

  return (
    <div className="space-y-6 max-w-4xl mx-auto">
      {/* Top Card */}
      <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
        <div className="flex items-center gap-3 mb-2">
          <div className="p-3 rounded-xl bg-amber-50 text-amber-600 border border-amber-200/60">
            <Database className="w-6 h-6" />
          </div>
          <div>
            <h3 className="font-serif font-bold text-slate-900 text-xl">
              Almacenamiento y Base de Datos Local
            </h3>
            <p className="text-xs text-slate-500">
              Sistema de persistencia en navegador (localStorage / Web API) de alta disponibilidad para SQP Legal Consulting.
            </p>
          </div>
        </div>

        {/* Feedback alert */}
        {feedbackMessage && (
          <div
            className={`mt-4 p-4 rounded-xl flex items-center gap-3 text-xs font-semibold ${
              feedbackMessage.type === 'success'
                ? 'bg-emerald-50 border border-emerald-200 text-emerald-800'
                : 'bg-red-50 border border-red-200 text-red-700'
            }`}
          >
            {feedbackMessage.type === 'success' ? (
              <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0" />
            ) : (
              <AlertCircle className="w-5 h-5 text-red-500 shrink-0" />
            )}
            <span>{feedbackMessage.text}</span>
          </div>
        )}

        {/* Stats Grid */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mt-6">
          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 text-center">
            <span className="text-[10px] uppercase font-bold text-slate-400 block mb-1">
              Clientes en Base
            </span>
            <span className="font-serif font-bold text-2xl text-slate-900">
              {stats.totalClients}
            </span>
          </div>

          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 text-center">
            <span className="text-[10px] uppercase font-bold text-slate-400 block mb-1">
              Plantillas Word
            </span>
            <span className="font-serif font-bold text-2xl text-slate-900">
              {stats.totalTemplates}
            </span>
          </div>

          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 text-center">
            <span className="text-[10px] uppercase font-bold text-slate-400 block mb-1">
              Docs Generados
            </span>
            <span className="font-serif font-bold text-2xl text-slate-900">
              {stats.totalGeneratedDocs}
            </span>
          </div>

          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 text-center">
            <span className="text-[10px] uppercase font-bold text-slate-400 block mb-1">
              Espacio Ocupado
            </span>
            <span className="font-mono font-bold text-2xl text-amber-700">
              {stats.storageUsageEstimateKb} KB
            </span>
          </div>
        </div>
      </div>

      {/* Backup & Restore Operations */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {/* Export Backup Card */}
        <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm flex flex-col justify-between">
          <div>
            <div className="w-10 h-10 rounded-xl bg-blue-50 text-blue-700 flex items-center justify-center mb-3 border border-blue-200">
              <Download className="w-5 h-5" />
            </div>
            <h4 className="font-bold text-slate-900 text-sm mb-1">
              Exportar Copia de Seguridad Completa
            </h4>
            <p className="text-xs text-slate-500 leading-relaxed mb-4">
              Descarga un archivo JSON consolidado con todos tus clientes, plantillas personalizadas y registros de documentos generados para archivar o migrar a otro ordenador.
            </p>
          </div>

          <button
            id="btn-export-backup-json"
            onClick={handleExportBackup}
            className="w-full inline-flex items-center justify-center gap-2 py-2.5 px-4 rounded-xl bg-slate-900 hover:bg-slate-800 text-white font-semibold text-xs transition-colors shadow-sm"
          >
            <Download className="w-4 h-4 text-amber-400" />
            <span>Descargar Respaldo JSON</span>
          </button>
        </div>

        {/* Import Backup Card */}
        <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm flex flex-col justify-between">
          <div>
            <div className="w-10 h-10 rounded-xl bg-amber-50 text-amber-700 flex items-center justify-center mb-3 border border-amber-200">
              <Upload className="w-5 h-5" />
            </div>
            <h4 className="font-bold text-slate-900 text-sm mb-1">
              Restaurar Copia de Seguridad
            </h4>
            <p className="text-xs text-slate-500 leading-relaxed mb-4">
              Carga un archivo de respaldo JSON generado previamente para sincronizar y restaurar tus clientes y plantillas en este navegador.
            </p>
          </div>

          <label
            htmlFor="backup-file-input"
            className="w-full inline-flex items-center justify-center gap-2 py-2.5 px-4 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold text-xs cursor-pointer transition-colors shadow-sm"
          >
            <Upload className="w-4 h-4" />
            <span>Seleccionar Archivo JSON de Respaldo</span>
            <input
              id="backup-file-input"
              type="file"
              accept=".json,application/json"
              onChange={handleImportBackup}
              className="hidden"
            />
          </label>
        </div>
      </div>

      {/* Reset Section */}
      <div className="bg-slate-50 rounded-2xl p-6 border border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-4">
        <div>
          <h4 className="font-bold text-slate-900 text-xs uppercase tracking-wider">
            Preparar para Producción / Limpieza de Datos
          </h4>
          <p className="text-xs text-slate-500 mt-0.5">
            Deja la base de datos totalmente limpia (0 clientes, 0 documentos y plantillas oficiales en blanco listas para usar).
          </p>
        </div>

        <button
          id="btn-clean-database"
          onClick={handleReset}
          className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-red-600 hover:bg-red-700 text-white text-xs font-semibold transition-colors shrink-0 shadow-xs"
        >
          <RotateCcw className="w-4 h-4" />
          <span>Limpiar Todos los Datos (0 Clientes)</span>
        </button>
      </div>
    </div>
  );
};
