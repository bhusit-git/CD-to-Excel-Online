import * as pdfjs from 'pdfjs-dist/build/pdf.mjs';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { generateLawsonPaymentPdfWorkbook } from './lawsonPaymentPdf.js';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

export function generatePdfWorkbook(buffer) {
  return generateLawsonPaymentPdfWorkbook(buffer, pdfjs);
}
