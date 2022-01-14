/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ha10id.reports.service;

import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.font.PdfFontFactory.EmbeddingStrategy;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.element.Paragraph;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.itextpdf.layout.Document;
//import static com.lowagie.text.ElementTags.FONT;

import org.springframework.stereotype.Service;

/**
 *
 * @author ha10id
 */

@Service
public class DocToPdfService {

    public Boolean convert(String src, String dst) throws FileNotFoundException, IOException {
        String k = null;
        OutputStream fileForPdf = null;
        String fileName = "./test.doc";
        String FONT = "./src/main/resources/font/SFNSText.ttf";
        PdfDocument pdfDoc = new PdfDocument(new PdfWriter("./DocToPdf.pdf"));
        //Below Code is for .doc file
        if (fileName.endsWith(".doc")) {
            HWPFDocument doc = new HWPFDocument(new FileInputStream(
                    fileName));
            WordExtractor we = new WordExtractor(doc);
            k = we.getText();
//                fileForPdf = new FileOutputStream(new File(
//                        "./DocToPdf.pdf"));
//                PdfDocument pdfDoc = new PdfDocument(new PdfWriter("./DocToPdf.pdf"));
            we.close();
        } //Below Code for .docx file
        else if (fileName.endsWith(".docx")) {
            XWPFDocument docx = new XWPFDocument(new FileInputStream(fileName));
            // using XWPFWordExtractor Class
            XWPFWordExtractor we = new XWPFWordExtractor(docx);
            k = we.getText();
//                fileForPdf = new FileOutputStream(new File(
//                        "./DocxToPdf.pdf"));
//                PdfDocument pdfDoc = new PdfDocument(new PdfWriter("./DocToPdf.pdf"));
            we.close();
        }
        PdfFont f1 = PdfFontFactory.createFont(FONT, PdfEncodings.IDENTITY_H);
        PdfFont f2 = PdfFontFactory.createFont(FONT, PdfEncodings.CP1250, EmbeddingStrategy.PREFER_EMBEDDED);
        try (Document document = new Document(pdfDoc)) {

            System.out.println(k);
            document.add(new Paragraph(k).setFont(f1));
            document.close();
//            fileForPdf.close();
        }
        return true;

    }
}
