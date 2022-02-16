package com.example.demo.service.export;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.repository.FactureRepository;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
@Service
public class FactureExportPdfService {

    @Autowired
    private FactureRepository factureRepository;
    private static Font catFont = new Font(Font.FontFamily.TIMES_ROMAN, 18, Font.BOLD);
    private DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
    public void export(OutputStream outputStream, Long idFacture) throws DocumentException {
        Facture facture = factureRepository.findById(idFacture).get();
        Client client = facture.getClient();
        Document document = new Document();
        PdfWriter.getInstance(document, outputStream);
        document.open();
        Paragraph paragrapheHeader1 = new Paragraph();
        paragrapheHeader1.add(new Paragraph(client.getPrenom() + " " + client.getNom(), catFont));
        paragrapheHeader1.add(new Paragraph(dateTimeFormatter.format(client.getDateNaissance()), catFont));
        document.add(paragrapheHeader1);
        Paragraph paragrapheHeader2 = new Paragraph();
        Paragraph site = new Paragraph("AMAZON.COM", catFont);
        site.setAlignment(Element.ALIGN_RIGHT);
        paragrapheHeader2.add(site);
        document.add(paragrapheHeader2);
        PdfPTable table = new PdfPTable(3);
        table.addCell("Désignation");
        table.addCell("Quantité");
        table.addCell("Prix unitaire");
        for (LigneFacture ligneFacture : facture.getLigneFactures()) {
            table.addCell(ligneFacture.getArticle().getLibelle());
            table.addCell("" + ligneFacture.getQuantite());
            table.addCell("" + ligneFacture.getArticle().getPrix());
        }
        document.add(table);
        document.close();
    }
}
