import { Controller, Post, Body, Res } from '@nestjs/common';
import { InvoiceService } from './invoice.service';
import { Response } from 'express';
import * as fs from 'fs';

@Controller('invoice')
export class InvoiceController {
  constructor(private readonly invoiceService: InvoiceService) {}

  @Post('generate')
  async generateInvoice(@Body() body: any, @Res() res: Response) {
    const filePath = await this.invoiceService.fillTemplateAndExportPDF(body);
    return res.download(filePath);
  }

  @Post('generate-zip')
  async generateInvoiceZip(@Body() body: { invoices: any[] }, @Res() res: Response) {
    try {
      const zipPath = await this.invoiceService.generateInvoiceZip(body.invoices);
      const today = new Date();
      const dateStr = `${String(today.getFullYear()).slice(2)}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
      const filename = `${dateStr}.zip`;
      
      res.set({
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename="${filename}"`,
        'Content-Length': fs.statSync(zipPath).size
      });
      
      const fileStream = fs.createReadStream(zipPath);
      fileStream.pipe(res);
      
      // Clean up the file after sending
      fileStream.on('end', () => {
        fs.unlinkSync(zipPath);
      });

      // Handle errors during file streaming
      fileStream.on('error', (err) => {
        console.error('Error streaming zip file:', err);
        if (!res.headersSent) {
          res.status(500).json({ error: 'Failed to stream zip file' });
        }
        // Clean up the file in case of error
        if (fs.existsSync(zipPath)) {
          fs.unlinkSync(zipPath);
        }
      });
    } catch (error) {
      console.error('Error generating zip:', error);
      if (!res.headersSent) {
        res.status(500).json({ error: error.message || 'Failed to generate zip file' });
      }
    }
  }
}
