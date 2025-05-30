import express, { Request, Response, RequestHandler } from 'express';
import multer from 'multer';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { TemplateHandler, MimeType } from 'easy-template-x';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as ExcelJS from 'exceljs';
import { downloadHandler } from './routes/download';

const execAsync = promisify(exec);
const app = express();
const port = 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.set('view engine', 'ejs');

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.join(__dirname, 'templates'));
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  }
});

const upload = multer({ storage });

// Directory paths
const DIRS = {
  assets: path.resolve(__dirname, 'assets'),
  templates: path.resolve(__dirname, 'templates'),
  generatedDocx: path.resolve(__dirname, 'output-generated', 'docx'),
  generatedPdf: path.resolve(__dirname, 'output-generated', 'pdf'),
  generatedExcel: path.resolve(__dirname, 'output-generated', 'excel')
};

// Ensure directories exist
Object.values(DIRS).forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// Function to find placeholders in Excel cells
async function findExcelPlaceholders(filePath: string): Promise<Set<string>> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const placeholders = new Set<string>();
  const placeholderRegex = /{{([^}]+)}}/g;

  console.log('Searching for placeholders in Excel file:', filePath);
  console.log('Number of worksheets:', workbook.worksheets.length);
  
  workbook.worksheets.forEach(worksheet => {
    console.log('Checking worksheet:', worksheet.name);
    let rowCount = 0;
    worksheet.eachRow((row) => {
      rowCount++;
      row.eachCell((cell) => {
        const rawValue = cell.value;
        console.log(`Row ${rowCount}, Column ${cell.col}, Raw value:`, rawValue);
        
        let cellValue = '';
        if (typeof rawValue === 'string') {
          cellValue = rawValue;
        } else if (rawValue && typeof rawValue === 'object' && 'result' in rawValue) {
          // Handle formula cells
          cellValue = rawValue.result?.toString() || '';
        } else if (rawValue && typeof rawValue === 'object' && 'text' in rawValue) {
          // Handle rich text cells
          cellValue = rawValue.text || '';
        }
        
        if (cellValue) {
          const matches = cellValue.match(placeholderRegex);
          if (matches) {
            matches.forEach(match => {
              // Remove {{ and }} to get the placeholder name
              const placeholder = match.slice(2, -2);
              console.log('Found placeholder:', placeholder, 'in cell value:', cellValue);
              placeholders.add(placeholder);
            });
          }
        }
      });
    });
    console.log(`Processed ${rowCount} rows in worksheet ${worksheet.name}`);
  });

  console.log('Total placeholders found:', placeholders.size);
  return placeholders;
}

// Function to replace placeholders in Excel file
async function processExcelTemplate(templatePath: string, outputPath: string, formData: any): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  const placeholderRegex = /{{([^}]+)}}/g;

  workbook.worksheets.forEach(worksheet => {
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (typeof cell.value === 'string') {
          let cellValue = cell.value as string;
          let hasPlaceholder = false;
          
          // Replace all placeholders in the cell
          cellValue = cellValue.replace(placeholderRegex, (match, placeholder) => {
            hasPlaceholder = true;
            return formData[placeholder] || match;
          });

          if (hasPlaceholder) {
            // Preserve the cell style while updating the value
            const style = { ...cell.style };
            cell.value = cellValue;
            cell.style = style;
          }
        }
      });
    });
  });

  await workbook.xlsx.writeFile(outputPath);
}

// Routes
app.get('/', (req, res) => {
  res.render('index');
});

// Upload template and extract placeholders
app.post('/upload', upload.single('template'), async (req, res) => {
  try {
    if (!req.file) {
      throw new Error('No file uploaded');
    }

    const templatePath = path.join(DIRS.templates, req.file.originalname);
    const fileExt = path.extname(req.file.originalname).toLowerCase();
    let uniquePlaceholders: Set<string>;    if (fileExt === '.xlsx') {
      // Handle Excel template
      console.log('Processing Excel template:', templatePath);
      uniquePlaceholders = await findExcelPlaceholders(templatePath);
      console.log('Placeholders found in Excel:', Array.from(uniquePlaceholders));
    } else if (fileExt === '.docx') {
      // Handle Word template
      const templateBuffer = await fs.promises.readFile(templatePath);
      const handler = new TemplateHandler();
      const tags = await handler.parseTags(templateBuffer);
      uniquePlaceholders = new Set<string>();
      for (const tag of tags) {
        console.log('Found placeholder:', tag.name);
        uniquePlaceholders.add(tag.name);
      }
    } else {
      throw new Error('Unsupported file format. Please upload a .docx or .xlsx file.');
    }

    if (uniquePlaceholders.size === 0) {
      throw new Error('No placeholders found in template. Make sure your template has placeholders in the format {{placeholder}}');
    }

    console.log(`Found ${uniquePlaceholders.size} unique placeholders`);
    
    return res.render('form', { 
      placeholders: Array.from(uniquePlaceholders),
      templateName: req.file.originalname
    });

  } catch (error: any) {
    console.error('Error processing template:', error);
    res.status(500).send(error.message);
  }
});

// Generate document from form data
app.post('/generate', async (req, res) => {
  try {
    const { templateName, formData } = req.body;
    const fileExt = path.extname(templateName).toLowerCase();
    const templatePath = path.join(DIRS.templates, templateName);
    const timestamp = new Date().getTime();
    
    if (fileExt === '.xlsx') {
      // Handle Excel template
      const outputExcel = path.join(DIRS.generatedExcel, `generated-${timestamp}.xlsx`);
      await processExcelTemplate(templatePath, outputExcel, formData);
      res.json({ 
        message: 'Excel file generated successfully',
        filename: `generated-${timestamp}.xlsx`,
        fileType: 'excel'
      });
    } else if (fileExt === '.docx') {
      // Handle Word template
      // Convert base64 image data to buffers
      Object.entries(formData).forEach(([key, value]: [string, any]) => {
        if (value && value._type === 'image') {
          console.log(`Processing image for ${key}`);
          // Convert base64 to Buffer
          const imageBuffer = Buffer.from(value.source, 'base64');
          // Update the image data to match the easy-template-x format
          formData[key] = {
            _type: 'image',
            source: imageBuffer,
            format: MimeType.Png,
            width: 150,
            height: 100,
            altText: value.altText || key,
            transparencyPercent: value.transparencyPercent || 0
          };

          console.log(`Image processed: ${key}`, {
            format: 'png',
            size: imageBuffer.length,
            width: value.width || 200,
            height: value.height || 200
          });
        }
      });

      const docxFilename = `generated-${timestamp}.docx`;
      const pdfFilename = `generated-${timestamp}.pdf`;
      const outputDocx = path.join(DIRS.generatedDocx, docxFilename);
      const outputPdf = path.join(DIRS.generatedPdf, pdfFilename);

      // Read template as buffer
      const templateContent = await fs.promises.readFile(templatePath);

      // Process template with form data
      const handler = new TemplateHandler();
      const doc = await handler.process(templateContent, formData);

      // Save generated DOCX
      await fs.promises.writeFile(outputDocx, doc);      
      // Convert to PDF
      await convertToPdf(outputDocx, outputPdf);
      
      // Debug log
      console.log('Generated files:', {
        docxFilename,
        pdfFilename,
        fileType: 'docx'
      });
      
      res.json({ 
        message: 'Files generated successfully',
        docxFilename,
        pdfFilename,
        fileType: 'docx'
      });
    } else {
      throw new Error('Unsupported file format');
    }
  } catch (error: any) {
    console.error('Error processing template:', error);
    res.status(500).send(error.message);
  }
});

// Download route handler
app.get('/download/:type/:filename', downloadHandler);

// Function to convert DOCX to PDF using LibreOffice
async function convertToPdf(inputPath: string, outputPath: string): Promise<void> {
  try {
    const absoluteInputPath = path.resolve(inputPath);
    const absoluteOutputDir = path.resolve(path.dirname(outputPath));
    
    if (!fs.existsSync(absoluteInputPath)) {
      throw new Error(`Input file not found: ${absoluteInputPath}`);
    }
    
    if (!fs.existsSync(absoluteOutputDir)) {
      fs.mkdirSync(absoluteOutputDir, { recursive: true });
    }

    const command = `soffice --headless --norestore --convert-to pdf:writer_pdf_Export --outdir "${absoluteOutputDir}" "${absoluteInputPath}"`;
    const { stdout, stderr } = await execAsync(command);
    
    const expectedPdfPath = path.join(absoluteOutputDir, path.basename(absoluteInputPath, '.docx') + '.pdf');
    if (!fs.existsSync(expectedPdfPath)) {
      throw new Error('PDF file was not created after conversion');
    }
  } catch (error: any) {
    throw new Error(`PDF conversion error: ${error.message}`);
  }
}

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
