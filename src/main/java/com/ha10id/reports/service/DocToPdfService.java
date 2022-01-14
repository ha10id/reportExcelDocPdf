/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ha10id.reports.service;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import java.util.Arrays;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.springframework.stereotype.Service;

/**
 * https://www.programmerall.com/article/2047471202/
 * @author ha10id
 */
@Service
public class DocToPdfService {
       private static final String EXCEL_XLS = "xls"; 
        private static final String EXCEL_XLSX = "xlsx";  
    public Boolean convert(String src, String dst) throws FileNotFoundException, IOException {
        byte[] pdfBytes = null;
        try {
            ByteArrayOutputStream pdfOutStream = new ByteArrayOutputStream();
            String fileName = "./test.xls";
            InputStream inputStream = new FileInputStream(fileName);
            HSSFWorkbook excelBook = new HSSFWorkbook();
            //Judge the Excel file to convert the 07+ version to the 03 version
            if(fileName.endsWith(EXCEL_XLS)){  //Excel 2003 
                excelBook = new HSSFWorkbook(inputStream); 
            }
            else if(fileName.endsWith(EXCEL_XLSX)){  // Excel 2007/2010 
                Transform xls = new Transform();   
                XSSFWorkbook workbookOld = new XSSFWorkbook(inputStream);
                xls.transformXSSF(workbookOld, excelBook);
            }  

//            //convert xls stream to html - w3 document
//
//            //convert above w3 document to html input stream
//            ByteArrayOutputStream htmlOutputStream = new ByteArrayOutputStream();
//            Source xmlSource = new DOMSource(w3Document);
//            Result outputTarget = new StreamResult(htmlOutputStream);
//            TransformerFactory.newInstance().newTransformer().transform(xmlSource, outputTarget);
//            InputStream htmlInputStream = new ByteArrayInputStream(htmlOutputStream.toByteArray());

            //convert html inputstream to pdf out stream
            ConverterProperties converterProperties = new ConverterProperties();
            HtmlConverter.convertToPdf(htmlInputStream, pdfOutStream, converterProperties);

            pdfBytes = pdfOutStream.toByteArray();
            System.out.println(Arrays.toString(pdfBytes));
            //write to physical pdf file
            OutputStream out = new FileOutputStream("./toPDF.pdf");
            out.write(pdfBytes);
            out.close();
        } catch (IOException | ParserConfigurationException | TransformerException e) {
            System.out.println(e);
            // log exception details and throw custom exception
        }

        return true;

    }
}
