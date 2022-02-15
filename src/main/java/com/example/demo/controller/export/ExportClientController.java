package com.example.demo.controller.export;

import com.example.demo.service.export.ClientExportCVSService;
import com.example.demo.service.export.ClientExportXLSXService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
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
@RequestMapping("export/clients")
public class ExportClientController {

    @Autowired
    private ClientExportCVSService clientExportCVSService;
    @Autowired
    private ClientExportXLSXService clientExportXLSXService;

    /**
     * Export des clients au format CSV.
     */
    @GetMapping("csv")
    public void exportCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-clients.csv\"");
        PrintWriter writer = response.getWriter();
        clientExportCVSService.export(writer);
    }

    @GetMapping("xlsx")
    public void exportCXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        // response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-client.xlsx\"");
        OutputStream outputStream = response.getOutputStream();
        clientExportXLSXService.export(outputStream);
    }
}
