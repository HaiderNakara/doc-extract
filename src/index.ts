import * as fs from 'fs/promises';
import * as path from 'path';
import * as pdfParse from 'pdf-parse';
import * as mammoth from 'mammoth';
import * as textract from 'textract';
import { promisify } from 'util';

// Types for the library
export interface DocumentContent {
  text: string;
  metadata?: {
    pages?: number;
    words?: number;
    characters?: number;
    fileSize?: number;
    fileName?: string;
  };
}

export interface PdfContent extends DocumentContent {
  metadata: DocumentContent['metadata'] & {
    pages: number;
    info?: any;
  };
}

export interface DocxContent extends DocumentContent {
  html?: string;
  messages?: any[];
}

export enum SupportedFormats {
  PDF = 'pdf',
  DOCX = 'docx',
  DOC = 'doc',
  PPTX = 'pptx',
  PPT = 'ppt',
  TXT = 'txt',
}

export class DocumentReaderError extends Error {
  constructor(message: string, public readonly code?: string) {
    super(message);
    this.name = 'DocumentReaderError';
  }
}

export class DocumentReader {
  private readonly textractFromFile = promisify(textract.fromFileWithPath);
  private debug = false;

  constructor(options?: { debug?: boolean }) {
    this.debug = options?.debug || false;
  }

  /**
   * Read any supported document format
   */
  async readDocument(filePath: string): Promise<DocumentContent> {
    try {
      await this.validateFile(filePath);
      const fileExtension = this.getFileExtension(filePath);
      const stats = await fs.stat(filePath);

      switch (fileExtension) {
        case SupportedFormats.PDF:
          return await this.readPdf(filePath, stats.size);
        case SupportedFormats.DOCX:
          return await this.readDocx(filePath, stats.size);
        case SupportedFormats.TXT:
          return await this.readTextFile(filePath, stats.size);
        case SupportedFormats.DOC:
        case SupportedFormats.PPTX:
        case SupportedFormats.PPT:
          return await this.readWithTextract(filePath, stats.size);
        default:
          throw new DocumentReaderError(
            `Unsupported file format: ${fileExtension}`,
            'UNSUPPORTED_FORMAT'
          );
      }
    } catch (error) {
      if (error instanceof DocumentReaderError) {
        throw error;
      }
      this.log(`Error reading document ${filePath}:`, error);
      throw new DocumentReaderError(
        `Failed to read document: ${error.message}`,
        'READ_ERROR'
      );
    }
  }

  /**
   * Read multiple documents from file paths
   */
  async readMultipleDocuments(filePaths: string[]): Promise<DocumentContent[]> {
    const results = await Promise.allSettled(
      filePaths.map(filePath => this.readDocument(filePath))
    );

    return results.map((result, index) => {
      if (result.status === 'rejected') {
        this.log(`Failed to read document ${filePaths[index]}:`, result.reason);
        throw new DocumentReaderError(
          `Failed to read document ${filePaths[index]}: ${result.reason.message}`,
          'MULTI_READ_ERROR'
        );
      }
      return result.value;
    });
  }

  /**
   * Read PDF file
   */
  async readPdf(filePath: string, fileSize?: number): Promise<PdfContent> {
    try {
      const buffer = await fs.readFile(filePath);
      const data = await pdfParse(buffer);

      return {
        text: data.text,
        metadata: {
          pages: data.numpages,
          words: this.countWords(data.text),
          characters: data.text.length,
          fileSize: fileSize || buffer.length,
          fileName: path.basename(filePath),
          info: data.info,
        },
      };
    } catch (error) {
      this.log(`Error reading PDF ${filePath}:`, error);
      throw new DocumentReaderError(`Failed to read PDF: ${error.message}`, 'PDF_READ_ERROR');
    }
  }

  /**
   * Read DOCX file
   */
  async readDocx(filePath: string, fileSize?: number): Promise<DocxContent> {
    try {
      // Extract raw text
      const textResult = await mammoth.extractRawText({ path: filePath });

      // Extract HTML (optional)
      const htmlResult = await mammoth.convertToHtml({ path: filePath });

      return {
        text: textResult.value,
        html: htmlResult.value,
        messages: [...textResult.messages, ...htmlResult.messages],
        metadata: {
          words: this.countWords(textResult.value),
          characters: textResult.value.length,
          fileSize,
          fileName: path.basename(filePath),
        },
      };
    } catch (error) {
      this.log(`Error reading DOCX ${filePath}:`, error);
      throw new DocumentReaderError(`Failed to read DOCX: ${error.message}`, 'DOCX_READ_ERROR');
    }
  }

  /**
   * Read PPT/PPTX files using textract
   */
  async readPowerPoint(filePath: string, fileSize?: number): Promise<DocumentContent> {
    return this.readWithTextract(filePath, fileSize);
  }

  /**
   * Read documents using textract (fallback for various formats)
   */
  private async readWithTextract(filePath: string, fileSize?: number): Promise<DocumentContent> {
    try {
      const text = await this.textractFromFile(filePath);

      return {
        text: text || '',
        metadata: {
          words: this.countWords(text || ''),
          characters: text?.length || 0,
          fileSize,
          fileName: path.basename(filePath),
        },
      };
    } catch (error) {
      this.log(`Error reading document with textract ${filePath}:`, error);
      throw new DocumentReaderError(
        `Failed to read document: ${error.message}`,
        'TEXTRACT_READ_ERROR'
      );
    }
  }

  /**
   * Read document from buffer
   */
  async readDocumentFromBuffer(
    buffer: Buffer,
    fileName: string,
    mimeType?: string,
  ): Promise<DocumentContent> {
    try {
      const fileExtension =
        this.getFileExtensionFromName(fileName) ||
        this.getExtensionFromMimeType(mimeType);

      switch (fileExtension) {
        case SupportedFormats.PDF:
          return await this.readPdfFromBuffer(buffer, fileName);
        case SupportedFormats.DOCX:
          return await this.readDocxFromBuffer(buffer, fileName);
        case SupportedFormats.DOC:
        case SupportedFormats.PPT:
        case SupportedFormats.PPTX:
          return await this.readWithTextractFromBuffer(buffer, fileName);
        case SupportedFormats.TXT:
          return await this.readTextFromBuffer(buffer, fileName);
        default:
          throw new DocumentReaderError(
            `Unsupported format for buffer reading: ${fileExtension}`,
            'UNSUPPORTED_BUFFER_FORMAT'
          );
      }
    } catch (error) {
      if (error instanceof DocumentReaderError) {
        throw error;
      }
      this.log(`Error reading document from buffer:`, error);
      throw new DocumentReaderError(
        `Failed to read document from buffer: ${error.message}`,
        'BUFFER_READ_ERROR'
      );
    }
  }

  /**
   * Read multiple documents from buffers
   */
  async readMultipleFromBuffers(
    buffers: Array<{ buffer: Buffer; fileName: string; mimeType?: string }>,
  ): Promise<DocumentContent[]> {
    const results = await Promise.allSettled(
      buffers.map(({ buffer, fileName, mimeType }) =>
        this.readDocumentFromBuffer(buffer, fileName, mimeType)
      )
    );

    return results.map((result, index) => {
      if (result.status === 'rejected') {
        this.log(`Failed to read buffer ${buffers[index].fileName}:`, result.reason);
        throw new DocumentReaderError(
          `Failed to read buffer ${buffers[index].fileName}: ${result.reason.message}`,
          'MULTI_BUFFER_READ_ERROR'
        );
      }
      return result.value;
    });
  }

  /**
   * Read PDF from buffer
   */
  private async readPdfFromBuffer(buffer: Buffer, fileName: string): Promise<PdfContent> {
    try {
      const data = await pdfParse(buffer);

      return {
        text: data.text,
        metadata: {
          pages: data.numpages,
          words: this.countWords(data.text),
          characters: data.text.length,
          fileSize: buffer.length,
          fileName,
          info: data.info,
        },
      };
    } catch (error) {
      this.log(`Error reading PDF from buffer:`, error);
      throw new DocumentReaderError(
        `Failed to read PDF from buffer: ${error.message}`,
        'PDF_BUFFER_READ_ERROR'
      );
    }
  }

  /**
   * Read DOCX from buffer
   */
  private async readDocxFromBuffer(buffer: Buffer, fileName: string): Promise<DocxContent> {
    try {
      const textResult = await mammoth.extractRawText({ buffer });
      const htmlResult = await mammoth.convertToHtml({ buffer });

      return {
        text: textResult.value,
        html: htmlResult.value,
        messages: [...textResult.messages, ...htmlResult.messages],
        metadata: {
          words: this.countWords(textResult.value),
          characters: textResult.value.length,
          fileSize: buffer.length,
          fileName,
        },
      };
    } catch (error) {
      this.log(`Error reading DOCX from buffer:`, error);
      throw new DocumentReaderError(
        `Failed to read DOCX from buffer: ${error.message}`,
        'DOCX_BUFFER_READ_ERROR'
      );
    }
  }

  /**
   * Read text from buffer
   */
  private async readTextFromBuffer(buffer: Buffer, fileName: string): Promise<DocumentContent> {
    try {
      const text = buffer.toString('utf-8');

      return {
        text,
        metadata: {
          words: this.countWords(text),
          characters: text.length,
          fileSize: buffer.length,
          fileName,
        },
      };
    } catch (error) {
      this.log(`Error reading text from buffer:`, error);
      throw new DocumentReaderError(
        `Failed to read text from buffer: ${error.message}`,
        'TEXT_BUFFER_READ_ERROR'
      );
    }
  }

  /**
   * Read documents from buffer using textract
   */
  private async readWithTextractFromBuffer(buffer: Buffer, fileName: string): Promise<DocumentContent> {
    try {
      // Create a temporary file to use with textract
      const tempDir = path.join(process.cwd(), 'temp');
      await fs.mkdir(tempDir, { recursive: true });
      const tempFilePath = path.join(tempDir, fileName);

      try {
        await fs.writeFile(tempFilePath, buffer);

        // Add specific configuration for PowerPoint files
        const options = {
          preserveLineBreaks: true,
          preserveOnlyMultipleLineBreaks: true,
          pdftotextOptions: {
            layout: 'raw'
          }
        };

        const text = await this.textractFromFile(tempFilePath, options);

        if (!text) {
          throw new Error('No text content could be extracted from the file');
        }

        return {
          text: text || '',
          metadata: {
            words: this.countWords(text || ''),
            characters: text?.length || 0,
            fileSize: buffer.length,
            fileName,
          },
        };
      } finally {
        // Clean up the temporary file
        try {
          await fs.unlink(tempFilePath);
        } catch (error) {
          this.log(`Failed to delete temporary file ${tempFilePath}:`, error);
        }
      }
    } catch (error) {
      this.log(`Error reading document from buffer with textract:`, error);
      throw new DocumentReaderError(
        `Failed to read PowerPoint file: ${error.message}. Please ensure the file is not corrupted and try again.`,
        'TEXTRACT_BUFFER_READ_ERROR'
      );
    }
  }

  /**
   * Check if file format is supported
   */
  isFormatSupported(filePath: string): boolean {
    const extension = this.getFileExtension(filePath);
    return Object.values(SupportedFormats).includes(extension as SupportedFormats);
  }

  /**
   * Check if file format is supported by filename
   */
  isFormatSupportedByName(fileName: string): boolean {
    const extension = this.getFileExtensionFromName(fileName);
    return Object.values(SupportedFormats).includes(extension as SupportedFormats);
  }

  /**
   * Get supported formats
   */
  getSupportedFormats(): string[] {
    return Object.values(SupportedFormats);
  }

  /**
   * Validate file exists and is readable
   */
  async validateFile(filePath: string): Promise<void> {
    try {
      await fs.access(filePath, fs.constants.R_OK);
      const stats = await fs.stat(filePath);

      if (!stats.isFile()) {
        throw new DocumentReaderError('Path is not a file', 'INVALID_FILE_PATH');
      }

      if (!this.isFormatSupported(filePath)) {
        throw new DocumentReaderError(
          `Unsupported file format. Supported formats: ${this.getSupportedFormats().join(', ')}`,
          'UNSUPPORTED_FORMAT'
        );
      }
    } catch (error) {
      if (error instanceof DocumentReaderError) {
        throw error;
      }
      throw new DocumentReaderError(`File validation failed: ${error.message}`, 'VALIDATION_ERROR');
    }
  }

  /**
   * Utility methods
   */
  private getFileExtension(filePath: string): string {
    return path.extname(filePath).toLowerCase().slice(1);
  }

  private getFileExtensionFromName(fileName: string): string {
    return path.extname(fileName).toLowerCase().slice(1);
  }

  private getExtensionFromMimeType(mimeType?: string): string {
    const mimeMap: Record<string, string> = {
      'application/pdf': 'pdf',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
      'application/msword': 'doc',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
      'application/vnd.ms-powerpoint': 'ppt',
      'text/plain': 'txt',
    };

    return mimeType ? mimeMap[mimeType] || '' : '';
  }

  private countWords(text: string): number {
    return text
      .trim()
      .split(/\s+/)
      .filter((word) => word.length > 0).length;
  }

  /**
   * Read text file
   */
  private async readTextFile(filePath: string, fileSize?: number): Promise<DocumentContent> {
    try {
      const text = await fs.readFile(filePath, 'utf-8');

      return {
        text,
        metadata: {
          words: this.countWords(text),
          characters: text.length,
          fileSize,
          fileName: path.basename(filePath),
        },
      };
    } catch (error) {
      this.log(`Error reading text file ${filePath}:`, error);
      throw new DocumentReaderError(
        `Failed to read text file: ${error.message}`,
        'TEXT_READ_ERROR'
      );
    }
  }

  private log(message: string, ...args: any[]): void {
    if (this.debug) {
      console.log(`[DocumentReader] ${message}`, ...args);
    }
  }
}

// Convenience function for quick usage
export async function readDocument(filePath: string): Promise<DocumentContent> {
  const reader = new DocumentReader();
  return reader.readDocument(filePath);
}

export async function readDocumentFromBuffer(
  buffer: Buffer,
  fileName: string,
  mimeType?: string,
): Promise<DocumentContent> {
  const reader = new DocumentReader();
  return reader.readDocumentFromBuffer(buffer, fileName, mimeType);
}

// Export the main class as default
export default DocumentReader;