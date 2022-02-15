package com.example.demo.service.export;

import com.example.demo.entity.Client;
import com.example.demo.repository.ClientRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.util.List;

@Service
public class ClientExportXLSXService {

    @Autowired
    private ClientRepository clientRepository;

    public void export(OutputStream outputStream) throws IOException {
        // google => apache poi
        Workbook workbook = new XSSFWorkbook(); //=> Fichier

        //créer un objet style
        CellStyle cellStyleHeader = workbook.createCellStyle();
        Font fontHeader = workbook.createFont();
        fontHeader.setColor(IndexedColors.PINK.index);
        fontHeader.setBold(true);
        cellStyleHeader.setFont(fontHeader);
        applyBorder(cellStyleHeader);

        CellStyle cellStyleData = workbook.createCellStyle();
        applyBorder(cellStyleData);

        // créer une feuille = Sheet
        Sheet sheet = workbook.createSheet("clients");
        // créer des lignes = Row
        Row rowHeader = sheet.createRow(0);
        // créer des cellules = Cell

        Cell cellHeader0 = rowHeader.createCell(0);
        cellHeader0.setCellValue("Nom");
        cellHeader0.setCellStyle(cellStyleHeader);
        Cell cellHeader1 = rowHeader.createCell(1);
        cellHeader1.setCellValue("Prénom");
        cellHeader1.setCellStyle(cellStyleHeader);
        Cell cellHeader2 = rowHeader.createCell(2);
        cellHeader2.setCellValue("Age");
        cellHeader2.setCellStyle(cellStyleHeader);

        List<Client> clients = clientRepository.findAll();
        int iRow = 1;
        for (Client client : clients) {
            Row row = sheet.createRow(iRow++);
            Cell cell0 = row.createCell(0);
            cell0.setCellStyle(cellStyleData);
            cell0.setCellValue(client.getNom());
            Cell cell1 = row.createCell(1);
            cell1.setCellStyle(cellStyleData);
            cell1.setCellValue(client.getPrenom());
            Cell cell2 = row.createCell(2);
            cell2.setCellStyle(cellStyleData);
            cell2.setCellValue(LocalDate.now().getYear() - client.getDateNaissance().getYear());
        }

        workbook.write(outputStream);
        workbook.close();
    }

    private void applyBorder(CellStyle cellStyleHeader) {
        cellStyleHeader.setBorderTop(BorderStyle.THICK);
        cellStyleHeader.setBorderRight(BorderStyle.THICK);
        cellStyleHeader.setBorderBottom(BorderStyle.THICK);
        cellStyleHeader.setBorderLeft(BorderStyle.THICK);
        cellStyleHeader.setTopBorderColor(IndexedColors.BLUE.index);
        cellStyleHeader.setRightBorderColor(IndexedColors.BLUE.index);
        cellStyleHeader.setBottomBorderColor(IndexedColors.BLUE.index);
        cellStyleHeader.setLeftBorderColor(IndexedColors.BLUE.index);
    }

}
