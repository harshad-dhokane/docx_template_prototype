import { RequestHandler } from 'express';
import path from 'path';
import fs from 'fs';

export const downloadHandler: RequestHandler = (req, res) => {
  try {
    const { type, filename } = req.params;
    let filePath;
    
    switch (type) {
      case 'pdf':
        filePath = path.join(process.cwd(), 'output-generated', 'pdf', filename);
        break;
      case 'excel':
        filePath = path.join(process.cwd(), 'output-generated', 'excel', filename);
        break;
      case 'docx':
        filePath = path.join(process.cwd(), 'output-generated', 'docx', filename);
        break;
      default:
        res.status(400).send('Invalid file type');
        return;
    }
    
    if (!fs.existsSync(filePath)) {
      res.status(404).send('File not found');
      return;
    }
    
    res.download(filePath);
  } catch (error: any) {
    console.error('Error downloading file:', error);
    res.status(500).send(error.message);
  }
};
