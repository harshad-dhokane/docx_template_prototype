import { RequestHandler } from 'express';
import path from 'path';
import fs from 'fs';

const DIRS = {
  generatedDocx: path.resolve(__dirname, '..', 'output-generated', 'docx'),
  generatedPdf: path.resolve(__dirname, '..', 'output-generated', 'pdf'),
  generatedExcel: path.resolve(__dirname, '..', 'output-generated', 'excel')
};

// Validate that directories exist
Object.values(DIRS).forEach(dir => {
  if (!fs.existsSync(dir)) {
    console.error(`Directory does not exist: ${dir}`);
  }
});

interface DownloadParams {
  type: 'pdf' | 'excel' | 'docx';
  filename: string;
}

// MIME types for different file formats
const MIME_TYPES = {
  pdf: 'application/pdf',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  excel: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

// Output directories by file type
const OUTPUT_DIRS = {
  pdf: 'generatedPdf',
  excel: 'generatedExcel',
  docx: 'generatedDocx'
};

export const downloadHandler: RequestHandler<DownloadParams> = (req, res) => {
  try {
    const { type, filename } = req.params;
    console.log('Download request:', { type, filename });
    
    // Additional validation
    if (!type || !filename) {
      console.error('Missing parameters:', { type, filename });
      res.status(400).send('Missing parameters');
      return;
    }

    // For safety, clean the filename
    const cleanFilename = path.basename(filename);
    console.log('Cleaned filename:', cleanFilename);
    let filePath;
    
    // Validate file type and get the correct directory
    let targetDir;
    switch (type) {
      case 'pdf':
        targetDir = DIRS.generatedPdf;
        break;
      case 'excel':
        targetDir = DIRS.generatedExcel;
        break;
      case 'docx':
        targetDir = DIRS.generatedDocx;
        break;
      default:
        console.error('Invalid file type:', type);
        res.status(400).send('Invalid file type');
        return;
    }
    
    // Check if directory exists
    if (!fs.existsSync(targetDir)) {
      console.error('Directory not found:', targetDir);
      res.status(500).send('Server configuration error');
      return;
    }
    
    filePath = path.join(targetDir, filename);
    
    // Log the absolute path being checked
    console.log('Looking for file at:', filePath);

    console.log('Attempting to download file:', filePath);
    
    if (!fs.existsSync(filePath)) {
      console.error('File not found:', filePath);
      res.status(404).send('File not found');
      return;
    }

    console.log('File exists, sending download...');    // Set the appropriate headers with proper encoding for filenames
    const encodedFilename = encodeURIComponent(filename);
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodedFilename}`);
    res.setHeader('Content-Type', MIME_TYPES[type]);
    res.setHeader('Content-Length', fs.statSync(filePath).size);
    
    // Stream the file
    const fileStream = fs.createReadStream(filePath);
    fileStream.pipe(res);
    
    fileStream.on('error', (error) => {
      console.error('Error streaming file:', error);
      if (!res.headersSent) {
        res.status(500).send('Error downloading file');
      }
    });

    fileStream.on('end', () => {
      console.log('File download completed:', filePath);
    });
    
  } catch (error: any) {
    console.error('Error downloading file:', error);
    if (!res.headersSent) {
      res.status(500).send(error.message);
    }
  }
};
