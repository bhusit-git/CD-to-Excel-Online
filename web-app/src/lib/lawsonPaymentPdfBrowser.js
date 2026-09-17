// Safari needs PDF.js compatibility polyfills (including Iterator helpers).
// Keep the API and worker on the same legacy build.
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url';
import { generateLawsonPaymentPdfWorkbook } from './lawsonPaymentPdf.js';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

export function generatePdfWorkbook(buffer) {
  return generateLawsonPaymentPdfWorkbook(buffer, pdfjs);
}
