import { RequestHandler } from 'express';
import path from 'path';
import fs from 'fs';

const DIRS = {
  generatedDocx: path.resolve(__dirname, '..', 'output-generated', 'docx'),
  generatedPdf: path.resolve(__dirname, '..', 'output-generated', 'pdf'),
  generatedExcel: path.resolve(__dirname, '..', 'output-generated', 'excel')
};

interface DownloadParams {
  type: 'pdf' | 'excel' | 'docx';
  filename: string;
}

export const downloadHandler: RequestHandler<DownloadParams> = (req, res) => {
  try {
    const { type, filename } = req.params;
    console.log('Download request:', { type, filename });
    let filePath;
    
    switch (type) {
      case 'pdf':
        filePath = path.join(DIRS.generatedPdf, filename);
        break;
      case 'excel':
        filePath = path.join(DIRS.generatedExcel, filename);
        break;
      case 'docx':
        filePath = path.join(DIRS.generatedDocx, filename);
        break;
      default:
        console.error('Invalid file type:', type);
        res.status(400).send('Invalid file type');
        return;
    }

    console.log('Attempting to download file:', filePath);
    
    if (!fs.existsSync(filePath)) {
      console.error('File not found:', filePath);
      res.status(404).send('File not found');
      return;
    }

    console.log('File exists, sending download...');
    
    // Set the appropriate headers
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', type === 'pdf' ? 'application/pdf' : 
                                type === 'docx' ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' :
                                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    
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
