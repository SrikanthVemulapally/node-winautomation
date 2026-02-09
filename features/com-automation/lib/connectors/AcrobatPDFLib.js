/**
 * @fileoverview Acrobat PDF-lib Integration
 * Advanced PDF operations using pdf-lib and pdf-parse
 * Install: npm install pdf-lib pdf-parse
 */

const fs = require('fs');

/**
 * @class AcrobatPDFLib
 * @description Advanced PDF operations using pdf-lib
 */
class AcrobatPDFLib {
    /**
     * Extract text from PDF
     * @param {string} pdfPath - Path to PDF file
     * @returns {Promise<string>} Extracted text
     */
    static async extractText(pdfPath) {
        try {
            const pdfParse = require('pdf-parse');
            const dataBuffer = fs.readFileSync(pdfPath);
            const data = await pdfParse(dataBuffer);
            return data.text;
        } catch (e) {
            throw new Error('Failed to extract text (requires pdf-parse): ' + e.message);
        }
    }
    
    /**
     * Merge multiple PDFs into one
     * @param {Array<string>} pdfPaths - Array of PDF file paths
     * @param {string} outputPath - Output file path
     * @returns {Promise<void>}
     */
    static async mergePDFs(pdfPaths, outputPath) {
        try {
            const { PDFDocument } = require('pdf-lib');
            const mergedPdf = await PDFDocument.create();
            
            for (const pdfPath of pdfPaths) {
                const pdfBytes = fs.readFileSync(pdfPath);
                const pdf = await PDFDocument.load(pdfBytes);
                const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                copiedPages.forEach(page => mergedPdf.addPage(page));
            }
            
            const mergedPdfBytes = await mergedPdf.save();
            fs.writeFileSync(outputPath, mergedPdfBytes);
        } catch (e) {
            throw new Error('Failed to merge PDFs (requires pdf-lib): ' + e.message);
        }
    }
    
    /**
     * Split PDF into multiple files
     * @param {string} pdfPath - Source PDF path
     * @param {string} outputDir - Output directory
     * @param {Array<Array<number>>} pageRanges - Array of page ranges [[1,3], [4,6]]
     * @returns {Promise<void>}
     */
    static async splitPDF(pdfPath, outputDir, pageRanges) {
        try {
            const { PDFDocument } = require('pdf-lib');
            const pdfBytes = fs.readFileSync(pdfPath);
            const pdfDoc = await PDFDocument.load(pdfBytes);
            
            if (!fs.existsSync(outputDir)) {
                fs.mkdirSync(outputDir, { recursive: true });
            }
            
            for (const [index, range] of pageRanges.entries()) {
                const newPdf = await PDFDocument.create();
                const pages = await newPdf.copyPages(pdfDoc, range);
                pages.forEach(page => newPdf.addPage(page));
                
                const newPdfBytes = await newPdf.save();
                fs.writeFileSync(`${outputDir}/split_${index + 1}.pdf`, newPdfBytes);
            }
        } catch (e) {
            throw new Error('Failed to split PDF (requires pdf-lib): ' + e.message);
        }
    }
    
    /**
     * Protect PDF with password
     * @param {string} pdfPath - Source PDF path
     * @param {string} outputPath - Output file path
     * @param {string} userPassword - User password
     * @param {string} ownerPassword - Owner password (optional)
     * @returns {Promise<void>}
     */
    static async protectPDF(pdfPath, outputPath, userPassword, ownerPassword = null) {
        try {
            const { PDFDocument } = require('pdf-lib');
            const pdfBytes = fs.readFileSync(pdfPath);
            const pdfDoc = await PDFDocument.load(pdfBytes);
            
            const encryptedPdfBytes = await pdfDoc.save({
                userPassword: userPassword,
                ownerPassword: ownerPassword || userPassword,
            });
            
            fs.writeFileSync(outputPath, encryptedPdfBytes);
        } catch (e) {
            throw new Error('Failed to protect PDF (requires pdf-lib): ' + e.message);
        }
    }
    
    /**
     * Get PDF page count
     * @param {string} pdfPath - Path to PDF file
     * @returns {Promise<number>} Number of pages
     */
    static async getPageCount(pdfPath) {
        try {
            const { PDFDocument } = require('pdf-lib');
            const pdfBytes = fs.readFileSync(pdfPath);
            const pdfDoc = await PDFDocument.load(pdfBytes);
            return pdfDoc.getPageCount();
        } catch (e) {
            throw new Error('Failed to get page count (requires pdf-lib): ' + e.message);
        }
    }
    
    /**
     * Extract specific pages to new PDF
     * @param {string} pdfPath - Source PDF path
     * @param {string} outputPath - Output file path
     * @param {Array<number>} pageNumbers - Array of page numbers (0-based)
     * @returns {Promise<void>}
     */
    static async extractPages(pdfPath, outputPath, pageNumbers) {
        try {
            const { PDFDocument } = require('pdf-lib');
            const pdfBytes = fs.readFileSync(pdfPath);
            const pdfDoc = await PDFDocument.load(pdfBytes);
            
            const newPdf = await PDFDocument.create();
            const pages = await newPdf.copyPages(pdfDoc, pageNumbers);
            pages.forEach(page => newPdf.addPage(page));
            
            const newPdfBytes = await newPdf.save();
            fs.writeFileSync(outputPath, newPdfBytes);
        } catch (e) {
            throw new Error('Failed to extract pages (requires pdf-lib): ' + e.message);
        }
    }
}

module.exports = { AcrobatPDFLib };
