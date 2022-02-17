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
        HashMap<Client, List<Facture>> factuByCliMap = new HashMap<Client, List<Facture> >();
        Multimap<Client, List<Facture>> factuByCliMap2 = ArrayListMultimap.create();
        Client c1 = new Client();
        Facture f1 = new Facture();


        // google => apache poi
        Workbook workbook = new XSSFWorkbook(); //=> Fichier XLSx

        //methode 1
        /*On parcourt les factures (a), puis une facture parcourt tout les clients (b).*
        Si un le champ 'client' d'une facture a correspond à un client b parcouru, on ajoute cette facture dans une liste de factures.
        Après avoir recupéré toutes les factures d'un client, on met le Client en tant que clé et la liste de factures en tant que valeur dans une HashMap
        commentaire a changer car j'ai inverser l'ordre
        */
        /*for (Facture facture : allFactures) {
            for ( Client client : allClients) {
                if(facture.getClient().equals(client)){
                    facturesByCli.add(facture);
                     //System.out.println("Facture ajouté : " + facture.getId() + " du client :" + client.getPrenom());
                }
                factuByCliMap.put(client,facturesByCli);
                //System.out.println("Facture ajouté2 : " + facturesByCli.get(0).getId() + " du client :" + client.getPrenom());
            }
        }*/
        //System.out.println("Valeur de retour : " + factuByCliMap);


//On crée une hashmap pour associer les factures à un client

        /*for (int i = 0; i <allClients.size() ; i++) {
            Client client = allClients.get(i);
            for (int j = 0; j < allFactures.size(); j++) {
                Facture facture = allFactures.get(j);
                if(facture.getClient().equals(client)){
                    facturesByCli.add(facture);
                }
            }
            factuByCliMap.put(client,facturesByCli);
            facturesByCli.clear();
        }*/

        for ( Client client : allClients) {
            for (Facture facture : allFactures ) {
                if(facture.getClient().equals(client)){
                    facturesByCli.add(facture);
                    System.out.println("Facture ajouté : " + facture.getId() + " du client :" + client.getPrenom());
                }
                //System.out.println("nom du client " + client.getNom());
            }
            List<Facture> newArray = new ArrayList<>(facturesByCli);
            factuByCliMap.put(client,newArray);
            facturesByCli.clear();
        }

        System.out.println("Valeur de retour : " + factuByCliMap);
        /*for (  Facture facture : allFactures) {
            for ( Client client : allClients) {
                if(facture.getClient().equals(client)){
                    facturesByCli.add(facture);
                }
                factuByCliMap.put(client,facturesByCli);
                System.out.println("nom du client " + client.getNom());
            }
        }*/

        //créer un objet style pour mettre en gras
        CellStyle cellStyleHeader = workbook.createCellStyle();
        Font fontHeader = workbook.createFont();

        fontHeader.setBold(true);
        cellStyleHeader.setFont(fontHeader);
        //CellStyle cellStyleData = workbook.createCellStyle();


        //On enlève le client qui n'a aucune factures de la collection en utilisant un iterator
            for(Iterator< Map.Entry <Client, List<Facture>> > it = factuByCliMap.entrySet().iterator(); it.hasNext(); ) {
                Map.Entry<Client, List<Facture>> clientEntry = it.next();
                if(clientEntry.getValue().size() == 0) {
                    System.out.println("On enlève le client de nom : " + clientEntry.getKey().getNom());
                    it.remove();
                }
        }

        for (Client client : factuByCliMap.keySet()) {

            //System.out.println("Je suis dans la boucle clients et mon client, " + client.getPrenom() + " a " + factuByCliMap.get(client).size() + " facture, et sa 1ere fac est : " + factuByCliMap.get(client).get(0).getId());

            Sheet sheetClient = workbook.createSheet(client.getNom() + " " + client.getPrenom());

            Row rowHeaderClient1 = sheetClient.createRow(0);
            Row rowHeaderClient2 = sheetClient.createRow(1);
            Row rowHeaderClient3 = sheetClient.createRow(2);
            Row rowHeaderClient4 = sheetClient.createRow(3);
            // créer des cellules = Cell

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
            for (Facture facture : factuByCliMap.get(client) ){

                rowHeaderClient4.createCell(iColFac).setCellValue(facture.getId());
                iColFac++;
                //metttre l'increment en dehors de la loop
                //then you can write the excel using this command:



                //System.out.println("facture num = " + facture.getId());
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

                //On adapte la taille des colonnes pour qu'il n'y ait pas de texte tronqué dans les feuilles factures(à utiliser quand il n'y a pas beaucoup de lignes)
                // et on selectionne les plusieurs cellules
                int nbRowsFact = sheetFact.getPhysicalNumberOfRows();
                for (int i = 0; i < nbRowsFact; i++) {
                    sheetFact.autoSizeColumn(i);
                    cellTotal.getAddress().formatAsString();
                    //sheetFact.setActiveCell("A1:B2");
                    sheetFact.setActiveCell(cellTotal.getAddress()); //méthode pour selectionner plusieurs cellules (sous forme de string) est deprécié , on ne peut mettre qu'une adresse
                }

                int firstRowNum = sheetFact.getTopRow(); //ex : pour A ce sera 1
                int lastRowNum = sheetFact.getLastRowNum() ;
                String firstCell = "A1";

                Row firstRowFact = sheetFact.getRow(firstRowNum);
                Row lastRowFact = sheetFact.getRow(lastRowNum);

                int lastCellNumFac = lastRowFact.getLastCellNum();
                String lastCellLetterFac = getCharForNumber(lastCellNumFac);// on transforme le chiffre en lettre

                //int firstCellNum = firstRowClient.getFirstCellNum(); // ex : 1
                ++lastRowNum; // on incremente car ca part de 0 à la base

                String lastCell =  lastCellLetterFac+lastRowNum;
                String cellRange = "A1:"+lastCell;


                //code recupéré sur stackOverFlow pour selectionner plusieurs cellules
                //https://stackoverflow.com/questions/67270531/apache-poi-setactivecell-for-multiple-cells?rq=1
                XSSFSheet xssfSheet = (XSSFSheet) sheetFact;
                xssfSheet.getCTWorksheet().getSheetViews().getSheetViewArray(0).getSelectionArray(0).setSqref(
                        java.util.Arrays.asList(cellRange));
            }

            //int firstRowNum = sheetClient.getTopRow(); //ex : pour A1 c'est 1
            int lastRowNumCli = sheetClient.getLastRowNum() ;
            //String firstCell = "A1";

            //Row firstRowClient = sheetClient.getRow(firstRowNum);
            Row lastRowClient = rowHeaderClient4; // la ligne du total

            int lastCellNumCli = lastRowClient.getLastCellNum();
            //++lastRowNum;
            String lastCellLetterCli = getCharForNumber(lastCellNumCli);// on transforme le chiffre en lettre

            //int firstCellNum = firstRowClient.getFirstCellNum(); // ex : 1


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

            //sheetClient.groupRow(0,++nbRowsClient);



            //cellBirthYear.setCellValue();


            workbook.write(outputStream);
            workbook.close();
    }
    //code recup sur stackoverflow pour transformer le numero du row en lettre
    private String getCharForNumber(int i) {
        return i > 0 && i < 27 ? String.valueOf((char)(i + 'A' - 1)) : null;
    }
}




        /*for (Facture facture : allFactures ) {
            for ( Client client : allClients) {
                if(facture.getClient().equals(client)){
                    facturesByCli.add(facture);
                }
                factuByCliMap.put(client,facturesByCli);
                System.out.println("nom du client" + client.getNom());
            }
            // créer une feuille = Sheet
            Sheet sheet = workbook.createSheet(client.getNom() + client.getPrenom());

            Sheet sheet = workbook.createSheet();
            Sheet sheet = workbook.createSheet();
            Sheet sheet = workbook.createSheet();
            Sheet sheet = workbook.createSheet();

            Row rowHeader = sheet.createRow(0);
            Row rowHeader2 = sheet.createRow(1);
            Row rowHeader3 = sheet.createRow(2);
            Row rowHeader4 = sheet.createRow(3);
            // créer des cellules = Cell

            Cell cellHeader0 = rowHeader.createCell(0);
            cellHeader0.setCellValue("Nom :");
            Cell cellName = rowHeader.createCell(1);
            cellName.setCellValue(client.getNom());


            Cell cellHeader1 = rowHeader2.createCell(0);
            cellHeader1.setCellValue("Prénom :");
            Cell cellFirstName = rowHeader.createCell(1);
            cellFirstName.setCellValue(client.getPrenom());

            Cell cellHeader2 = rowHeader3.createCell(0);
            cellHeader2.setCellValue("Année de naissance :");
            Cell cellBirthYear = rowHeader.createCell(1);
            //cellBirthYear.setCellValue();


            /*List<Client> clients = clientRepository.findAll();
            int iRow = 1;
            for (Client client : clients) {
                Row row = sheet.createRow(iRow++);
                Cell cell0 = row.createCell(0);
                cell0.setCellValue(client.getNom());
                Cell cell1 = row.createCell(1);
                cell1.setCellValue(client.getPrenom());
                Cell cell2 = row.createCell(2);
                cell2.setCellValue(LocalDate.now().getYear() - client.getDateNaissance().getYear());
            }

            workbook.write(outputStream);
            workbook.close();*/




        // créer une feuille = Sheet





// toutes les clé clients ont acces a toutes les factures