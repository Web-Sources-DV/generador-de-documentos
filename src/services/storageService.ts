import { Client, Template, GeneratedDocument, DatabaseStats } from '../types';
import { getInitialDefaultTemplates } from './docxService';

const STORAGE_KEYS = {
  CLIENTS: 'sqp_clients_v2',
  TEMPLATES: 'sqp_templates_v2',
  DOCUMENTS: 'sqp_generated_docs_v2',
  SETTINGS: 'sqp_settings_v2',
  LAST_BACKUP: 'sqp_last_backup_v2',
};

// Purge any legacy v1 test mock data from browser localStorage
try {
  if (typeof window !== 'undefined' && window.localStorage) {
    ['sqp_clients_v1', 'sqp_generated_docs_v1', 'sqp_templates_v1', 'sqp_settings_v1', 'sqp_last_backup_v1'].forEach((k) => {
      window.localStorage.removeItem(k);
    });
  }
} catch (e) {
  // ignore
}

// Clean production initial states (0 clients, 0 generated documents)
const INITIAL_CLIENTS: Client[] = [];
const INITIAL_DOCUMENTS: GeneratedDocument[] = [];

export function getStoredClients(): Client[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.CLIENTS);
    if (!raw) {
      localStorage.setItem(STORAGE_KEYS.CLIENTS, JSON.stringify(INITIAL_CLIENTS));
      return [];
    }
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    console.error('Error reading clients from storage:', e);
    return [];
  }
}

export function saveClients(clients: Client[]): void {
  localStorage.setItem(STORAGE_KEYS.CLIENTS, JSON.stringify(clients));
}

export function saveClient(client: Client): Client {
  const clients = getStoredClients();
  const index = clients.findIndex((c) => c.id === client.id);
  const now = new Date().toISOString();

  if (index >= 0) {
    clients[index] = { ...client, updatedAt: now };
  } else {
    clients.unshift({ ...client, createdAt: client.createdAt || now, updatedAt: now });
  }

  saveClients(clients);
  return client;
}

export function deleteClient(id: string): void {
  const clients = getStoredClients().filter((c) => c.id !== id);
  saveClients(clients);
}

export function clearAllClients(): void {
  saveClients([]);
}

export function clearAllDocuments(): void {
  localStorage.setItem(STORAGE_KEYS.DOCUMENTS, JSON.stringify([]));
}

export function getStoredTemplates(): Template[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.TEMPLATES);
    if (!raw) {
      const initial = getInitialDefaultTemplates();
      localStorage.setItem(STORAGE_KEYS.TEMPLATES, JSON.stringify(initial));
      return initial;
    }
    const parsed: Template[] = JSON.parse(raw);
    if (!parsed || parsed.length === 0) {
      const initial = getInitialDefaultTemplates();
      localStorage.setItem(STORAGE_KEYS.TEMPLATES, JSON.stringify(initial));
      return initial;
    }
    return parsed;
  } catch (e) {
    console.error('Error reading templates from storage:', e);
    return getInitialDefaultTemplates();
  }
}

export function saveTemplates(templates: Template[]): void {
  localStorage.setItem(STORAGE_KEYS.TEMPLATES, JSON.stringify(templates));
}

export function saveTemplate(template: Template): Template {
  const templates = getStoredTemplates();
  const index = templates.findIndex((t) => t.id === template.id);
  const now = new Date().toISOString();

  if (index >= 0) {
    templates[index] = { ...template, updatedAt: now };
  } else {
    templates.unshift({ ...template, createdAt: template.createdAt || now, updatedAt: now, usageCount: 0 });
  }

  saveTemplates(templates);
  return template;
}

export function deleteTemplate(id: string): void {
  const templates = getStoredTemplates().filter((t) => t.id !== id);
  saveTemplates(templates);
}

export function getStoredDocuments(): GeneratedDocument[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.DOCUMENTS);
    if (!raw) {
      localStorage.setItem(STORAGE_KEYS.DOCUMENTS, JSON.stringify(INITIAL_DOCUMENTS));
      return [];
    }
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    console.error('Error reading documents from storage:', e);
    return [];
  }
}

export function saveDocumentLog(doc: GeneratedDocument): void {
  const docs = getStoredDocuments();
  docs.unshift(doc);
  localStorage.setItem(STORAGE_KEYS.DOCUMENTS, JSON.stringify(docs));

  // Increment template usage
  const templates = getStoredTemplates();
  const tpl = templates.find((t) => t.id === doc.templateId);
  if (tpl) {
    tpl.usageCount = (tpl.usageCount || 0) + 1;
    saveTemplates(templates);
  }

  // Increment client document count
  const clients = getStoredClients();
  const cli = clients.find((c) => c.id === doc.clientId);
  if (cli) {
    cli.documentCount = (cli.documentCount || 0) + 1;
    saveClients(clients);
  }
}

export function deleteDocumentLog(id: string): void {
  const docs = getStoredDocuments().filter((d) => d.id !== id);
  localStorage.setItem(STORAGE_KEYS.DOCUMENTS, JSON.stringify(docs));
}

export function getDatabaseStats(): DatabaseStats {
  const clients = getStoredClients();
  const templates = getStoredTemplates();
  const docs = getStoredDocuments();
  const lastBackup = localStorage.getItem(STORAGE_KEYS.LAST_BACKUP) || undefined;

  let totalBytes = 0;
  for (const key in localStorage) {
    if (key.startsWith('sqp_')) {
      totalBytes += (localStorage.getItem(key) || '').length * 2;
    }
  }

  return {
    totalClients: clients.length,
    totalTemplates: templates.length,
    totalGeneratedDocs: docs.length,
    lastBackupDate: lastBackup,
    storageUsageEstimateKb: Math.round(totalBytes / 1024),
  };
}

export function exportFullDatabaseJson(): string {
  const clients = getStoredClients();
  const templates = getStoredTemplates();
  const documents = getStoredDocuments();

  const exportObj = {
    appName: 'SQP Legal Consulting · Gestión Documental',
    version: '2.0.0',
    exportedAt: new Date().toISOString(),
    clients,
    templates,
    documents,
  };

  localStorage.setItem(STORAGE_KEYS.LAST_BACKUP, new Date().toISOString());
  return JSON.stringify(exportObj, null, 2);
}

export function importDatabaseJson(jsonString: string): { success: boolean; message: string } {
  try {
    const data = JSON.parse(jsonString);
    if (!data.clients && !data.templates && !data.documents) {
      return { success: false, message: 'El archivo JSON no contiene una estructura válida de respaldo de SQP Legal.' };
    }

    if (Array.isArray(data.clients)) {
      saveClients(data.clients);
    }
    if (Array.isArray(data.templates)) {
      saveTemplates(data.templates);
    }
    if (Array.isArray(data.documents)) {
      localStorage.setItem(STORAGE_KEYS.DOCUMENTS, JSON.stringify(data.documents));
    }

    return {
      success: true,
      message: `Base de datos restaurada con éxito: ${data.clients?.length || 0} clientes, ${data.templates?.length || 0} plantillas y ${data.documents?.length || 0} documentos.`,
    };
  } catch (err: any) {
    return { success: false, message: `Error al procesar el archivo JSON: ${err.message}` };
  }
}

export function resetDatabaseToDefaults(): void {
  localStorage.setItem(STORAGE_KEYS.CLIENTS, JSON.stringify([]));
  localStorage.setItem(STORAGE_KEYS.TEMPLATES, JSON.stringify(getInitialDefaultTemplates()));
  localStorage.setItem(STORAGE_KEYS.DOCUMENTS, JSON.stringify([]));
  localStorage.removeItem(STORAGE_KEYS.LAST_BACKUP);
}

export function clearEntireDatabase(): void {
  localStorage.setItem(STORAGE_KEYS.CLIENTS, JSON.stringify([]));
  localStorage.setItem(STORAGE_KEYS.DOCUMENTS, JSON.stringify([]));
  localStorage.removeItem(STORAGE_KEYS.LAST_BACKUP);
}

export const getClients = getStoredClients;
export const getTemplates = getStoredTemplates;
export const getGeneratedDocuments = getStoredDocuments;
