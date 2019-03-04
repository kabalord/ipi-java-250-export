package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.ClientFactureExportXlsx;
import com.example.demo.service.ClientService;
import com.example.demo.service.ExporterCSV;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;

    @Autowired
    private ClientFactureExportXlsx clientFactureExportXlsx;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        /*for(Client client : allClients)
        writer.println("Id;Nom;Prenom;Date naissance;Age");
        for (Iterator<Client> i = allClients.iterator(); i.hasNext();){
            Client client = i.next();

            writer.println(client.getId() + ";"
                    + client.getNom() + ";"
                    + client.getPrenom() + ";"
                    + client.getDateNaissance() + ";"
                    + (now.getYear() - client.getDateNaissance().getYear())
            );


        }*/

        ExporterCSV<Client> export = new ExporterCSV<>();
        export.addColumnLong("Id", c1 -> c1.getId());
        export.addColumnString("Nom", c -> c.getNom());
        export.addColumnString("Prénom", c -> c.getPrenom());
        export.addColumnString("Date de naissance", c -> c.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));
        export.addColumnInteger("Age", c -> now.getYear() - c.getDateNaissance().getYear());

        export.createCSV(response.getWriter(), allClients);

    }

    // pour obtenir une cell dans un fichier excel


    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");

        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        Workbook workbook = new XSSFWorkbook(); // crée un fichier excel
        Sheet sheet = workbook.createSheet("Clients"); // crée un onglet
        Row headerRow = sheet.createRow(0); // crée une ligne

        ArrayList<String> celTitres = new ArrayList<String>() ;
        celTitres.add("Matricule");
        celTitres.add("Prénom");
        celTitres.add("Nom");
        celTitres.add("Date de Naissance");
        celTitres.add("Age");

        int i =0;
        for(String title : celTitres){
            headerRow.createCell(i).setCellValue(title);
            i++;
        }
        //Cell cellPrenom = headerRow.createCell(0); // crée une cellule
        //cellPrenom.setCellValue("Matricule;Prénom;Nom;Date de Naissance"); // ajoute une valeur dans la cellule

        int rowNum = 1;
        for(Client client : allClients){
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(client.getId());
            row.createCell(1).setCellValue(client.getPrenom());
            row.createCell(2).setCellValue(client.getNom());
            row.createCell(3).setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")));
            row.createCell(4).setCellValue((now.getYear()-client.getDateNaissance().getYear()));
        }



        workbook.write(response.getOutputStream());
        workbook.close();

    }

    // TP wiki états imprimés

    @GetMapping("/factures/xlsx")
    public void facturesXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        List<Facture> factures = factureService.findAllFacture();
        List<Client> clients = clientService.findAllClients();

        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");

        Workbook workbook = new XSSFWorkbook();

        for (Client client: clients) {
            Sheet sheet = workbook.createSheet(client.getNom());
            Row headerRow = sheet.createRow(1);
            Cell cellPrenom = headerRow.createCell(0);
            cellPrenom.setCellValue(client.getPrenom());
            Cell cellNom = headerRow.createCell(1);
            cellNom.setCellValue(client.getNom());
            for (Facture facture: factures) {
                if(client.getId().equals(facture.getClient().getId())) {
                    createFacture(facture, workbook);
                }
            }
        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }

    private void createFacture(Facture facture, Workbook workbook){
        //Style
        Font font= workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.RED.getIndex());


        CellStyle cellStyleTopBot = workbook.createCellStyle();
        cellStyleTopBot.setBorderBottom(BorderStyle.DOUBLE);
        cellStyleTopBot.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        cellStyleTopBot.setBorderTop(BorderStyle.DOUBLE);
        cellStyleTopBot.setTopBorderColor(IndexedColors.BLUE.getIndex());

        CellStyle cellStyleTopBotLeft = workbook.createCellStyle();
        cellStyleTopBotLeft.setBorderLeft(BorderStyle.DOUBLE);
        cellStyleTopBotLeft.setLeftBorderColor(IndexedColors.BLUE.getIndex());
        cellStyleTopBotLeft.setBorderBottom(BorderStyle.DOUBLE);
        cellStyleTopBotLeft.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        cellStyleTopBotLeft.setBorderTop(BorderStyle.DOUBLE);
        cellStyleTopBotLeft.setTopBorderColor(IndexedColors.BLUE.getIndex());
        cellStyleTopBotLeft.setFont(font);

        CellStyle cellStyleTopBotRight = workbook.createCellStyle();
        cellStyleTopBotRight.setBorderRight(BorderStyle.DOUBLE);
        cellStyleTopBotRight.setRightBorderColor(IndexedColors.BLUE.getIndex());
        cellStyleTopBotRight.setBorderBottom(BorderStyle.DOUBLE);
        cellStyleTopBotRight.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        cellStyleTopBotRight.setBorderTop(BorderStyle.DOUBLE);
        cellStyleTopBotRight.setTopBorderColor(IndexedColors.BLUE.getIndex());
        cellStyleTopBotRight.setFont(font);


        Sheet sheet = workbook.createSheet("Factures"+facture.getId().toString());
        Row headerRow = sheet.createRow(2);
        Cell cellHeaderArticle = headerRow.createCell(1);
        cellHeaderArticle.setCellValue("Article");
        Cell cellHeaderQt = headerRow.createCell(2);
        cellHeaderQt.setCellValue("Quantité");
        Cell cellHeaderPrixU= headerRow.createCell(3);
        cellHeaderPrixU.setCellValue("Prix U");
        Cell cellHeaderPrixT = headerRow.createCell(4);
        cellHeaderPrixT.setCellValue("Prix Total");

        Integer r=4;
        double total=0;

        for (LigneFacture ligne:facture.getLigneFactures()) {
            Row row = sheet.createRow(r++);

            Cell cell = row.createCell(1);
            cell.setCellValue(ligne.getArticle().getLibelle());

            cell = row.createCell(2);
            double qte =ligne.getQuantite();
            cell.setCellValue(qte);

            cell = row.createCell(3);
            double pu =ligne.getArticle().getPrix();
            cell.setCellValue(pu);

            cell = row.createCell(4);
            pu =ligne.getArticle().getPrix();
            cell.setCellValue(pu*qte);

            total+=pu*qte;

        }


        Row totalRow = sheet.createRow(r++);
        Cell cell = totalRow.createCell(1);
        cell.setCellValue("Total");
        cell.setCellStyle(cellStyleTopBotLeft);
        cell = totalRow.createCell(2);
        cell.setCellStyle(cellStyleTopBot);
        cell = totalRow.createCell(3);
        cell.setCellStyle(cellStyleTopBot);
        cell = totalRow.createCell(4);
        cell.setCellValue(total);
        cell.setCellStyle(cellStyleTopBotRight);

    }

}
