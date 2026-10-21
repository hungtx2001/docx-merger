/**
 * Options for configuring the docx merge process.
 */
export interface DocxMergerOptions {
  /**
   * Whether to insert a section page break between merged documents.
   * If false, documents flow continuously.
   * @default true
   */
  pageBreak?: boolean;

  /**
   * Style collision resolution strategy.
   * - 'source': Retains custom styles from source documents by prefixing them.
   * - 'target': Adopts the primary document's styles whenever conflicts arise.
   * @default 'source'
   */
  style?: 'source' | 'target';
}

/**
 * Supported output formats when saving the merged document.
 */
export type SaveOutputFormat =
  | 'nodebuffer'
  | 'blob'
  | 'arraybuffer'
  | 'uint8array'
  | 'base64'
  | 'binarystring';

/**
 * Accepted input types for each DOCX file.
 */
export type InputFileType =
  | Buffer
  | Uint8Array
  | ArrayBuffer
  | string
  | Blob;

/**
 * Statistics and details about the merged document.
 */
export interface MergeMetadata {
  documentCount: number;
  totalImages: number;
  totalSections: number;
  hasNumbering: boolean;
}
