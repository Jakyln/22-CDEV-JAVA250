package com.example.demo.service.export;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.repository.ClientRepository;
import com.example.demo.repository.FactureRepository;
import com.google.common.collect.ArrayListMultimap;
import com.google.common.collect.Multimap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.util.*;

@Service
public class FactureExportXLSXService {

    @Autowired
    private FactureRepository factureRepository;

    @Autowired
    private ClientRepository clientRepository;

    public void export(OutputStream outputStream) throws IOException {
        List<Facture> allFactures = factureRepository.findAll();
        List<Facture> facturesByCli = new ArrayList<>();
        List<Client> allClients = clientRepository.findAll();
        //On crée une hashmap pour associer les factures à un client
        HashMap<Client, List<Facture>> factuByCliMap = new HashMap<Client, List<Facture> >();


        // google => apache poi
        Workbook workbook = new XSSFWorkbook(); //=> Fichier XLSx

        //methode 1
        /*On parcourt les clients (a), puis un client parcourt toutes les factures (b).
        Si un le champ 'client' d'une facture b correspond à un client a parcouru, on ajoute cette facture dans une liste de factures.
        Après avoir recupéré toutes les factures d'un client, on met le Client en tant que clé et la liste de factures en tant que valeur dans une HashMap
        */
        for ( Client client : allClients) {
            for (Facture facture : allFactures ) {
                if(facture.getClient().equals(client)){
                    facturesByCli.add(facture);
                    System.out.println("Facture ajouté : " + facture.getId() + " du client :" + client.getPrenom());
                }
            }
            List<Facture> newArray = new ArrayList<>(facturesByCli);
            factuByCliMap.put(client,newArray);
            facturesByCli.clear();
        }
        //créer un objet style pour mettre en gras
        CellStyle cellStyleHeader = workbook.createCellStyle();
        Font fontHeader = workbook.createFont();
        fontHeader.setBold(true);
        cellStyleHeader.setFont(fontHeader);

        //On enlève le client qui n'a aucunes factures de la collection en utilisant un iterator
            for(Iterator< Map.Entry <Client, List<Facture>> > it = factuByCliMap.entrySet().iterator(); it.hasNext(); ) {
                Map.Entry<Client, List<Facture>> clientEntry = it.next();
                if(clientEntry.getValue().size() == 0) {
                    System.out.println("On enlève le client de nom : " + clientEntry.getKey().getNom());
                    it.remove();
                }
        }
        for (Client client : factuByCliMap.keySet()) {

            Sheet sheetClient = workbook.createSheet(client.getNom() + " " + client.getPrenom());
            Row rowHeaderClient1 = sheetClient.createRow(0);
            Row rowHeaderClient2 = sheetClient.createRow(1);
            Row rowHeaderClient3 = sheetClient.createRow(2);
            Row rowHeaderClient4 = sheetClient.createRow(3);

            Cell cellHeaderClient1 = rowHeaderClient1.createCell(0);
            cellHeaderClient1.setCellValue("Nom :");
            Cell cellHeaderClient2 = rowHeaderClient1.createCell(1);
            cellHeaderClient2.setCellValue(client.getNom());

            Cell cellHeaderClient3 = rowHeaderClient2.createCell(0);
            cellHeaderClient3.setCellValue("Prénom :");
            Cell cellHeaderClient4 = rowHeaderClient2.createCell(1);
            cellHeaderClient4.setCellValue(client.getPrenom());

            Cell cellHeaderClient5 = rowHeaderClient3.createCell(0);
            cellHeaderClient5.setCellValue("Année de naissance :");
            Cell cellHeaderClient6 = rowHeaderClient3.createCell(1);
            cellHeaderClient6.setCellValue(client.getDateNaissance().getYear());

            Cell cellHeaderClient7 = rowHeaderClient4.createCell(0);
            cellHeaderClient7.setCellValue(factuByCliMap.get(client).size() + " Facture(s) :");
            cellHeaderClient7.setCellStyle(cellStyleHeader);

            int iColNbFac = factuByCliMap.get(client).size();
            for (int i = 0 ,iColFacTemp = 1; i < iColNbFac; i++,iColFacTemp++) {
                Facture facture = factuByCliMap.get(client).get(i);
                rowHeaderClient4.createCell(iColFacTemp).setCellValue(facture.getId());
            }
            int iColFac = 1;
            List<Cell> lastAndFirstCellFac = new ArrayList<>();

            int iActiveCellFac = 0;
            for (Facture facture : factuByCliMap.get(client) ){

                rowHeaderClient4.createCell(iColFac).setCellValue(facture.getId());
                iColFac++;
                Sheet sheetFact = workbook.createSheet("Facture n°"+facture.getId());

                Row rowHeader = sheetFact.createRow(0);
                Cell cellHeader1 = rowHeader.createCell(0);
                cellHeader1.setCellValue("Désignation");
                cellHeader1.setCellStyle(cellStyleHeader);

                Cell cellHeader2= rowHeader.createCell(1);
                cellHeader2.setCellValue("Quantité");
                cellHeader2.setCellStyle(cellStyleHeader);

                Cell cellHeader3 = rowHeader.createCell(2);
                cellHeader3.setCellValue("Prix Unitaire");
                cellHeader3.setCellStyle(cellStyleHeader);

                double prixTotal = 0;
                int iRow = 0;
                for (LigneFacture ligneFacture : facture.getLigneFactures() ) {
                    Article article = ligneFacture.getArticle();
                    Row row = sheetFact.createRow(++iRow);
                    Cell cell0 = row.createCell(0);
                    cell0.setCellValue(article.getLibelle());
                    Cell cell1 = row.createCell(1);
                    cell1.setCellValue(article.getStock());
                    Cell cell2 = row.createCell(2);
                    cell2.setCellValue(article.getPrix());
                    prixTotal += article.getPrix();
                }
                Row row = sheetFact.createRow(++iRow);
                Cell cell = row.createCell(1);
                cell.setCellValue("Total");
                cell.setCellStyle(cellStyleHeader);

                Cell cellTotal = row.createCell(2);
                cellTotal.setCellValue(prixTotal);

                int lastRowNum = sheetFact.getLastRowNum() ; //ex : pour A ce sera 1
                Row firstRowFact = sheetFact.getRow(0);

                Row lastRowFact = sheetFact.getRow(lastRowNum);

                int lastCellNum = lastRowFact.getLastCellNum();
                Cell lastCell = lastRowFact.getCell(lastCellNum);

                int firstCellNum = firstRowFact.getFirstCellNum();
                Cell firstCell = firstRowFact.getCell(firstCellNum);

                /*lastAndFirstCellFac.add(lastRowFact.getCell(lastRowFact.getLastCellNum()));
                lastAndFirstCellFac.add(firstRowFact.getCell(0)); // on recup la derniere cellule (pour pouvoir l'avoir en premier) puis la première*/
                lastAndFirstCellFac.add(lastCell);
                lastAndFirstCellFac.add(firstCell);
                System.out.println("la list : " + lastAndFirstCellFac);
                System.out.println("la list détaillé  : " + lastAndFirstCellFac.get(0));

                //On adapte la taille des colonnes pour qu'il n'y ait pas de texte tronqué dans les feuilles factures(à utiliser quand il n'y a pas beaucoup de lignes)
                // et on selectionne les plusieurs cellules
                int nbRowsFact = sheetFact.getPhysicalNumberOfRows();
                for (int i = 0; i < nbRowsFact; i++) {
                    sheetFact.autoSizeColumn(i);
                }
                sheetFact.setActiveCell(lastAndFirstCellFac.get(1).getAddress()); //méthode pour selectionner plusieurs cellules (sous forme de string) est deprécié , on ne peut mettre qu'une adresse


                int lastCellNumFac = lastRowFact.getLastCellNum();
                String lastCellLetterFac = getCharForNumber(lastCellNumFac);// on transforme le chiffre en lettre

                ++lastRowNum; // on incremente car ca part de 0 à la base

                String lastCellFac =  lastCellLetterFac+lastRowNum;
                String cellRangeFac = "A1:"+lastCellFac;

                //code recupéré sur stackOverFlow pour selectionner plusieurs cellules
                //https://stackoverflow.com/questions/67270531/apache-poi-setactivecell-for-multiple-cells?rq=1
                XSSFSheet xssfSheet = (XSSFSheet) sheetFact;
                xssfSheet.getCTWorksheet().getSheetViews().getSheetViewArray(0).getSelectionArray(0).setSqref(
                        java.util.Arrays.asList(cellRangeFac));

                iActiveCellFac++;
            }
            int lastRowNumCli = sheetClient.getLastRowNum() ;

            Row lastRowClient = rowHeaderClient4; // la ligne du total

            int lastCellNumCli = lastRowClient.getLastCellNum();
            String lastCellLetterCli = getCharForNumber(lastCellNumCli);// on transforme le chiffre en lettre

            //On adapte la taille des colonnes pour qu'il n'y ait pas de texte tronqué dans les feuilles client (à utiliser quand il n'y a pas beaucoup de lignes)
            int nbRowsClient = sheetClient.getPhysicalNumberOfRows();
            for (int i = 0; i < nbRowsClient; i++) {
                sheetClient.autoSizeColumn(i);
            }
            sheetClient.setActiveCell(lastRowClient.getCell(--lastCellNumCli).getAddress());

            ++lastRowNumCli ; // on incremente car ca part de 0 à la base

            String lastCellCli =  lastCellLetterCli+lastRowNumCli; //ex : A5
            String cellRangeCli = "A1:"+lastCellCli;

            //code recupéré sur stackOverFlow pour selectionner plusieurs cellules
            //https://stackoverflow.com/questions/67270531/apache-poi-setactivecell-for-multiple-cells?rq=1
            XSSFSheet xssfSheetCli = (XSSFSheet) sheetClient;
            xssfSheetCli.getCTWorksheet().getSheetViews().getSheetViewArray(0).getSelectionArray(0).setSqref(
                    java.util.Arrays.asList(cellRangeCli));
        }
            workbook.write(outputStream);
            workbook.close();
    }
    //code recup sur stackoverflow pour transformer le numero du row en lettre
    private String getCharForNumber(int i) {
        return i > 0 && i < 27 ? String.valueOf((char)(i + 'A' - 1)) : null;
    }
}