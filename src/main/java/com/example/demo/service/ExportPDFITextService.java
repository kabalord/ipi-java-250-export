package com.example.demo.service;

import com.example.demo.entity.Facture;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.OutputStream;

@Service

public class ExportPDFITextService {

    public void export(OutputStream os, Facture facture) throws IOException, DocumentException {
        Document document = new Document();
        PdfWriter.getInstance(document, os);
        document.open();

        document.add(new Paragraph("Hello"));
        document.close();
    }
}
