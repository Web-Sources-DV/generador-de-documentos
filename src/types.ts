export interface Client {
  id: string;
  firstName: string;
  lastName: string;
  fullName: string;
  passportNumber: string;
  docType: 'pasaporte' | 'dni' | 'nie' | 'cedula' | 'otro';
  nationality: string;
  issuingCountry: string;
  birthDate: string;
  expiryDate: string;
  issueDate?: string;
  sex: 'M' | 'F' | 'X' | 'Otro' | string;
  sexAgeCategory?: 'VARÓN' | 'MUJER' | 'JOVEN' | 'MENOR' | string;
  age?: number;
  personalNumber?: string;
  placeOfBirth?: string;
  email?: string;
  phone?: string;
  address?: string;
  city?: string;
  notes?: string;
  passportImageBase64?: string;
  createdAt: string;
  updatedAt: string;
  documentCount?: number;
}

export interface PlaceholderDef {
  key: string;
  label: string;
  description?: string;
  defaultValue?: string;
  type: 'text' | 'date' | 'textarea' | 'select' | 'number';
  options?: string[];
  category?: 'cliente' | 'pasaporte' | 'legal' | 'fechas' | 'personalizado';
}

export interface Template {
  id: string;
  name: string;
  description: string;
  category: 'Notarial' | 'Contratos' | 'Migratorio' | 'Judicial' | 'Corporativo' | 'General';
  fileName: string;
  fileData?: string; // Base64 of .docx file
  placeholders: string[]; // List of {{tags}}
  placeholderDefs: PlaceholderDef[];
  isDefault?: boolean;
  samplePreviewText?: string;
  createdAt: string;
  updatedAt: string;
  usageCount: number;
}

export interface GeneratedDocument {
  id: string;
  title: string;
  fileName: string;
  templateId: string;
  templateName: string;
  clientId: string;
  clientName: string;
  passportNumber: string;
  generatedAt: string;
  fileSizeFormatted: string;
  dataSnapshot: Record<string, string>;
  fileBase64?: string; // Optional stored generated docx
}

export interface ExtractionResult {
  firstName: string;
  lastName: string;
  fullName: string;
  passportNumber: string;
  nationality: string;
  issuingCountry: string;
  birthDate: string;
  expiryDate: string;
  issueDate?: string;
  sex: string;
  sexAgeCategory?: 'VARÓN' | 'MUJER' | 'JOVEN' | 'MENOR' | string;
  age?: number;
  docType?: 'pasaporte' | 'dni' | 'nie' | 'cedula' | 'otro';
  documentType?: string; // e.g. "Pasaporte" or "Cédula"
  personalNumber?: string;
  placeOfBirth?: string;
  mrzLine1?: string;
  mrzLine2?: string;
  confidenceScore?: number;
  notes?: string;
  method: 'ai' | 'tesseract' | 'manual';
  rawOcrText?: string;
  imagePreview?: string;
}

export type ActiveTab = 'wizard' | 'generator' | 'clients' | 'templates' | 'history' | 'database';

export interface DatabaseStats {
  totalClients: number;
  totalTemplates: number;
  totalGeneratedDocs: number;
  lastBackupDate?: string;
  storageUsageEstimateKb: number;
}
