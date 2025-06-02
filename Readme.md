# Doc Extract

A powerful Node.js library for reading and extracting text from various document formats including PDF, DOCX, DOC, PPT, PPTX, and TXT files.

## Features

- 📄 **Multiple Format Support**: PDF, DOCX, DOC, PPT, PPTX, TXT
- 🔍 **Text Extraction**: Extract clean text content from documents
- 📊 **Rich Metadata**: Get document statistics (word count, character count, pages, etc.)
- 💾 **Buffer Support**: Read documents from memory buffers
- 🔧 **TypeScript**: Full TypeScript support with type definitions
- 🚀 **Promise-based**: Modern async/await API
- 🛡️ **Error Handling**: Comprehensive error handling with custom error types

## Installation

```bash
npm install doc-extract
```

### System Dependencies

This library depends on some system packages for full functionality:

**For PDF support:**

- No additional dependencies required

**For PowerPoint and DOC support:**

```bash
# Ubuntu/Debian
sudo apt-get install antiword unrtf poppler-utils tesseract-ocr

# macOS
brew install antiword unrtf poppler tesseract

# Windows
# Install poppler and tesseract manually or use chocolatey:
choco install poppler tesseract
```

## Quick Start

```typescript
import DocumentReader, { readDocument } from "doc-extract";

// Simple usage
const content = await readDocument("./path/to/document.pdf");
console.log(content.text);
console.log(content.metadata);

// Using the class for more control
const reader = new DocumentReader({ debug: true });
const content = await reader.readDocument("./path/to/document.docx");
```

## API Reference

### Class: DocumentReader

#### Constructor

```typescript
new DocumentReader(options?: { debug?: boolean })
```

- `options.debug`: Enable debug logging (default: false)

#### Methods

##### readDocument(filePath: string): Promise<DocumentContent>

Read a document from file path.

```typescript
const reader = new DocumentReader();
const content = await reader.readDocument("./document.pdf");
```

##### readDocumentFromBuffer(buffer: Buffer, fileName: string, mimeType?: string): Promise<DocumentContent>

Read a document from a Buffer.

```typescript
const fs = require("fs");
const buffer = fs.readFileSync("./document.pdf");
const content = await reader.readDocumentFromBuffer(buffer, "document.pdf");
```

##### readMultipleDocuments(filePaths: string[]): Promise<DocumentContent[]>

Read multiple documents at once.

```typescript
const contents = await reader.readMultipleDocuments([
  "./doc1.pdf",
  "./doc2.docx",
  "./doc3.pptx",
]);
```

##### readMultipleFromBuffers(buffers: Array<{buffer: Buffer, fileName: string, mimeType?: string}>): Promise<DocumentContent[]>

Read multiple documents from buffers.

```typescript
const contents = await reader.readMultipleFromBuffers([
  { buffer: buffer1, fileName: "doc1.pdf" },
  { buffer: buffer2, fileName: "doc2.docx" },
]);
```

##### Specific Format Methods

```typescript
// PDF specific
const pdfContent = await reader.readPdf("./document.pdf");

// DOCX specific (includes HTML conversion)
const docxContent = await reader.readDocx("./document.docx");
console.log(docxContent.html); // HTML version of the document

// PowerPoint specific
const pptContent = await reader.readPowerPoint("./presentation.pptx");
```

##### Utility Methods

```typescript
// Check if format is supported
const isSupported = reader.isFormatSupported("./document.pdf"); // true

// Get all supported formats
const formats = reader.getSupportedFormats(); // ['pdf', 'docx', 'doc', 'pptx', 'ppt', 'txt']

// Validate file
await reader.validateFile("./document.pdf"); // throws error if invalid
```

### Convenience Functions

```typescript
import { readDocument, readDocumentFromBuffer } from "doc-extract";

// Quick read from file
const content = await readDocument("./document.pdf");

// Quick read from buffer
const content = await readDocumentFromBuffer(buffer, "document.pdf");
```

## Types

### DocumentContent

```typescript
interface DocumentContent {
  text: string;
  metadata?: {
    pages?: number;
    words?: number;
    characters?: number;
    fileSize?: number;
    fileName?: string;
  };
}
```

### PdfContent

```typescript
interface PdfContent extends DocumentContent {
  metadata: DocumentContent["metadata"] & {
    pages: number;
    info?: any; // PDF metadata from pdf-parse
  };
}
```

### DocxContent

```typescript
interface DocxContent extends DocumentContent {
  html?: string; // HTML version of the document
  messages?: any[]; // Conversion messages from mammoth
}
```

### SupportedFormats

```typescript
enum SupportedFormats {
  PDF = "pdf",
  DOCX = "docx",
  DOC = "doc",
  PPTX = "pptx",
  PPT = "ppt",
  TXT = "txt",
}
```

## Error Handling

The library uses custom error types for better error handling:

```typescript
import { DocumentReaderError } from "doc-extract";

try {
  const content = await readDocument("./nonexistent.pdf");
} catch (error) {
  if (error instanceof DocumentReaderError) {
    console.log("Error code:", error.code);
    console.log("Error message:", error.message);
  }
}
```

### Error Codes

- `UNSUPPORTED_FORMAT`: File format not supported
- `READ_ERROR`: General read error
- `PDF_READ_ERROR`: PDF-specific read error
- `DOCX_READ_ERROR`: DOCX-specific read error
- `TEXTRACT_READ_ERROR`: Textract-related error
- `BUFFER_READ_ERROR`: Buffer reading error
- `VALIDATION_ERROR`: File validation error
- `INVALID_FILE_PATH`: Invalid file path

## Examples

### Express.js Integration

```typescript
import express from "express";
import multer from "multer";
import { DocumentReader } from "doc-extract";

const app = express();
const upload = multer();
const reader = new DocumentReader();

app.post("/upload", upload.single("document"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const content = await reader.readDocumentFromBuffer(
      req.file.buffer,
      req.file.originalname,
      req.file.mimetype
    );

    res.json({
      text: content.text,
      metadata: content.metadata,
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});
```

### Batch Processing

```typescript
import { DocumentReader } from "doc-extract";
import { promises as fs } from "fs";
import path from "path";

async function processDocumentsInDirectory(dirPath: string) {
  const reader = new DocumentReader({ debug: true });

  const files = await fs.readdir(dirPath);
  const documentPaths = files
    .filter((file) => reader.isFormatSupportedByName(file))
    .map((file) => path.join(dirPath, file));

  const results = await reader.readMultipleDocuments(documentPaths);

  results.forEach((content, index) => {
    console.log(`Document ${documentPaths[index]}:`);
    console.log(`Words: ${content.metadata?.words}`);
    console.log(`Characters: ${content.metadata?.characters}`);
    console.log("---");
  });
}
```

### Search in Documents

```typescript
import { DocumentReader } from "doc-extract";

async function searchInDocument(filePath: string, searchTerm: string) {
  const reader = new DocumentReader();
  const content = await reader.readDocument(filePath);

  const lines = content.text.split("\n");
  const matchingLines = lines
    .map((line, index) => ({ line, lineNumber: index + 1 }))
    .filter(({ line }) =>
      line.toLowerCase().includes(searchTerm.toLowerCase())
    );

  return {
    totalMatches: matchingLines.length,
    matches: matchingLines,
    metadata: content.metadata,
  };
}
```

## Performance Tips

1. **Reuse DocumentReader instances** - The class can be reused for multiple operations
2. **Use batch methods** - `readMultipleDocuments()` is more efficient than individual calls
3. **Enable debug mode** only during development
4. **Clean up temporary files** - The library handles this automatically, but ensure your temp directory has sufficient space

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Development Setup

```bash
git clone https://github.com/HaiderNakara/doc-extract.git
cd doc-extract
npm install
npm run build
npm test
```

### Running Tests

```bash
npm test          # Run tests once
npm run test:watch # Run tests in watch mode
npm run test:coverage # Run tests with coverage
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
