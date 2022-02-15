package com.example.demo.service.export;

import com.example.demo.entity.Facture;
import com.example.demo.repository.FactureRepository;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfWriter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.OutputStream;

@Service
public class FactureExportPdfService {

    @Autowired
    private FactureRepository factureRepository;

    public void export(OutputStream outputStream, Long idFacture) throws DocumentException {
        Facture facture = factureRepository.findById(idFacture).get();
        Document document = new Document();
        PdfWriter.getInstance(document, outputStream);
        document.open();
        document.close();
    }
}
