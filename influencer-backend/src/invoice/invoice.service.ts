import { Injectable, InternalServerErrorException } from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';
import { exec } from 'child_process';
import { join } from 'path';
import { v4 as uuid } from 'uuid';
import * as archiver from 'archiver';

@Injectable()
export class InvoiceService {
  private readonly assetsDir: string;
  private readonly tmpDir: string;

  constructor() {
    this.assetsDir = join(process.cwd(), 'assets');
    this.tmpDir = join(process.cwd(), 'tmp');
    
    // Create directories if they don't exist
    if (!fs.existsSync(this.assetsDir)) {
      fs.mkdirSync(this.assetsDir, { recursive: true });
    }
    if (!fs.existsSync(this.tmpDir)) {
      fs.mkdirSync(this.tmpDir, { recursive: true });
    }
  }

  async generateInvoiceZip(invoicesData: any[]): Promise<string> {
    const today = new Date();
    const dateStr = `${String(today.getFullYear()).slice(2)}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
    const zipFileName = `${dateStr}.zip`;
    const zipFilePath = join(this.tmpDir, zipFileName);

    const generatedFiles: { path: string; name: string }[] = [];
    const errors: { invoiceNum: string; error: string }[] = [];
    
    try {
      // Process PDFs sequentially to avoid LibreOffice conflicts
      type PdfResult = 
        | { success: true; file: { path: string; name: string } }
        | { success: false; invoiceNum: string };
        
      const pdfFiles: PdfResult[] = [];
      
      for (const data of invoicesData) {
        try {
          // Add a small delay between conversions
          if (pdfFiles.length > 0) {
            await new Promise(resolve => setTimeout(resolve, 500));
          }
          
          const pdfPath = await this.fillTemplateAndExportPDF(data);
          const file = { path: pdfPath, name: `${data.invoiceNum}.pdf` };
          generatedFiles.push(file);
          pdfFiles.push({ success: true, file });
        } catch (error) {
          console.error(`Failed to generate PDF for invoice ${data.invoiceNum}:`, error);
          errors.push({ 
            invoiceNum: data.invoiceNum, 
            error: error.message || 'Unknown error'
          });
          pdfFiles.push({ success: false, invoiceNum: data.invoiceNum });
        }
      }

      // Filter out successful PDF generations
      const successfulPdfs = pdfFiles
        .filter((result): result is { success: true; file: { path: string; name: string } } => 
          result.success
        )
        .map(result => result.file);

      if (successfulPdfs.length === 0) {
        throw new Error('No PDFs were successfully generated. Errors: ' + 
          JSON.stringify(errors, null, 2));
      }

      // Create and return the zip file
      return await new Promise((resolve, reject) => {
        try {
          // Create a write stream for the zip file
          const output = fs.createWriteStream(zipFilePath);
          const archive = archiver('zip', {
            zlib: { level: 9 } // Maximum compression
          });

          // Listen for all archive data to be written
          output.on('close', () => {
            this.cleanupFiles(generatedFiles);
            
            // If there were some failures but some successes, resolve with a warning
            if (errors.length > 0) {
              console.warn('Some PDFs failed to generate:', errors);
            }
            
            resolve(zipFilePath);
          });

          output.on('end', () => {
            console.log('Data has been drained');
          });

          // Handle errors
          archive.on('error', (err) => {
            this.cleanupFiles(generatedFiles);
            if (fs.existsSync(zipFilePath)) {
              fs.unlinkSync(zipFilePath);
            }
            reject(new Error(`Failed to create zip file: ${err.message}`));
          });

          output.on('error', (err) => {
            this.cleanupFiles(generatedFiles);
            if (fs.existsSync(zipFilePath)) {
              fs.unlinkSync(zipFilePath);
            }
            reject(new Error(`Failed to write zip file: ${err.message}`));
          });

          // Pipe archive data to the file
          archive.pipe(output);

          // Add the PDF files to the archive
          successfulPdfs.forEach(file => {
            if (fs.existsSync(file.path)) {
              archive.file(file.path, { name: file.name });
            } else {
              console.error(`PDF file not found: ${file.path}`);
            }
          });

          // Add a report of failed PDFs if any
          if (errors.length > 0) {
            const errorReport = 'Failed to generate the following PDFs:\n' +
              errors.map(e => `${e.invoiceNum}: ${e.error}`).join('\n');
            archive.append(errorReport, { name: '_errors.txt' });
          }

          // Finalize the archive
          archive.finalize();
        } catch (error) {
          this.cleanupFiles(generatedFiles);
          if (fs.existsSync(zipFilePath)) {
            fs.unlinkSync(zipFilePath);
          }
          reject(error);
        }
      });
    } catch (error) {
      // Clean up any generated files if something goes wrong
      this.cleanupFiles(generatedFiles);
      if (fs.existsSync(zipFilePath)) {
        fs.unlinkSync(zipFilePath);
      }
      throw error;
    }
  }

  private cleanupFiles(files: { path: string }[]): void {
    files.forEach(file => {
      try {
        if (fs.existsSync(file.path)) {
          fs.unlinkSync(file.path);
        }
      } catch (err) {
        console.error(`Failed to clean up file ${file.path}:`, err);
      }
    });
  }

  async fillTemplateAndExportPDF(data: any): Promise<string> {
    try {
      const templatePath = join(this.assetsDir, 'Invoice.xlsx');
      
      // Check if template exists
      if (!fs.existsSync(templatePath)) {
        throw new Error('Invoice template not found. Please ensure Invoice.xlsx exists in the assets directory.');
      }

      const tempXlsx = join(this.tmpDir, `invoice_${uuid()}.xlsx`);
      const tempPdf = tempXlsx.replace('.xlsx', '.pdf');

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(templatePath);
      const sheet = workbook.getWorksheet('Invoice');

      if (!sheet) {
        throw new Error("시트 'Invoice'를 찾을 수 없습니다. 엑셀 파일의 시트 이름을 확인하세요.");
      }

      // Set today's date in F4
      const today = new Date();
      const formattedDate = `${today.getFullYear()}.${String(today.getMonth() + 1).padStart(2, '0')}.${String(today.getDate()).padStart(2, '0')}`;
      const dateCell = sheet.getCell('F4');
      dateCell.value = formattedDate;
      dateCell.alignment = { horizontal: 'center' };

      // Set invoice number in F5
      const invoiceNumCell = sheet.getCell('F5');
      invoiceNumCell.value = data.invoiceNum;
      invoiceNumCell.alignment = { horizontal: 'center' };

      sheet.getCell('B12').value = data.fullName;
      sheet.getCell('B15').value = data.fullAddress;
      sheet.getCell('F19').value = data.amount;
      sheet.getCell('B26').value = ` Bank name: ${data.bankName}`;
      sheet.getCell('B27').value = ` Branch name: ${data.bankBranch}`;
      sheet.getCell('B29').value = ` Account number: ${data.accountNumber}`;
      sheet.getCell('B30').value = ` Account holder: ${data.fullName}`;

      // Get the amount and tax rate
      const amount = Number(data.amount);
      
      // Calculate subtotal (F24) - sum of F19 to F23
      // Since we only have F19 with a value, subtotal will be equal to amount
      const subtotal = amount;  // =SUM(F19:F23)
      
      // Get tax rate from F25
      const taxRate = Number(sheet.getCell('F25').value) || 0;
      
      // Set tax amount (F26) to "-" if no tax, otherwise calculate it
      if (taxRate === 0) {
        const cell = sheet.getCell('F26');
        cell.value = "-";
        cell.alignment = { horizontal: 'right' };
      } else {
        const taxAmount = subtotal * taxRate;
        sheet.getCell('F26').value = taxAmount;
      }
      
      // Set Other section (F27) to "-" with right alignment
      const otherCell = sheet.getCell('F27');
      otherCell.value = "-";
      otherCell.alignment = { horizontal: 'right' };
      
      // Calculate total (F28) - when no tax, total is just subtotal
      const total = taxRate === 0 ? subtotal : subtotal + (subtotal * taxRate);

      // Set calculated values
      sheet.getCell('F24').value = subtotal;  // Subtotal
      sheet.getCell('F28').value = total;  // Total

      // Format amount cells with Japanese number format with 2 decimal places
      ['F19', 'F24'].forEach(cellRef => {
        const cell = sheet.getCell(cellRef);
        cell.numFmt = '#,##0.00';  // Japanese number format with 2 decimal places
      });

      // Format total with yen sign
      const totalCell = sheet.getCell('F28');
      totalCell.numFmt = '¥#,##0.00';  // Japanese number format with yen sign

      // Format tax amount cell if it has a value
      if (taxRate > 0) {
        sheet.getCell('F26').numFmt = '#,##0.00';
      }

      await workbook.xlsx.writeFile(tempXlsx);
      await this.convertToPDF(tempXlsx, tempPdf);
      
      // Clean up the temporary Excel file
      fs.unlinkSync(tempXlsx);
      
      return tempPdf;
    } catch (error) {
      console.error('Error generating invoice:', error);
      throw new InternalServerErrorException(error.message);
    }
  }

  private async convertToPDF(input: string, output: string): Promise<void> {
    const outDir = join(output, '..');
    const libreOfficePath = '/opt/homebrew/bin/soffice'; // Default Homebrew installation path
    
    return new Promise((resolve, reject) => {
      exec(`${libreOfficePath} --headless --convert-to pdf ${input} --outdir ${outDir}`, (err) => {
        if (err) {
          console.error('LibreOffice conversion error:', err);
          reject(new Error('PDF conversion failed. Please ensure LibreOffice is installed and accessible.'));
          return;
        }
        resolve();
      });
    });
  }
}
