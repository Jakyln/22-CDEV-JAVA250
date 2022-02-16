package com.example.demo.controller.export;

import com.example.demo.service.export.ClientExportCVSService;
import com.example.demo.service.export.ClientExportXLSXService;
import com.example.demo.service.export.FactureExportPdfService;
import com.example.demo.service.export.FactureExportXLSXService;
import com.itextpdf.text.DocumentException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintWriter;

/**
 * Controller pour réaliser l'export des clients.
 */
@Controller
@RequestMapping("export/factures")
public class ExportFactureController {

    @Autowired
    private FactureExportPdfService factureExportPdfService;

    @Autowired
    private FactureExportXLSXService factureExportXLSXService;

    // /**
    //  * Export des clients au format CSV.
    //  */
    // @GetMapping("xlsx")
    // public void exportCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
    //     response.setContentType("text/csv");
    //     response.setHeader("Content-Disposition", "attachment; filename=\"export-factures.xslx\"");
    //     OutputStream outputStream = response.getOutputStream();
    //     factureExportXLSXService.export(outputStream);
    // }

    @GetMapping("{id}/pdf")
    public void exportCXLSX(
            @PathVariable Long id,
            HttpServletRequest request,
            HttpServletResponse response) throws IOException, DocumentException {
        response.setHeader("Content-Disposition", "attachment; filename=\"export-facture-" + id + ".pdf\"");
        OutputStream outputStream = response.getOutputStream();
        factureExportPdfService.export(outputStream, id);
    }

    @GetMapping("xlsx")
    public void exportCXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        // response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-factures.xlsx\"");
        OutputStream outputStream = response.getOutputStream();
        factureExportXLSXService.export(outputStream);
    }
}
