package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.service.ClientService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        writer.println("Id;Nom;Prenom;Date naissance;Age");
        // for(Client client : allClients)
        LocalDate now = LocalDate.now();
        for (Iterator<Client> i = allClients.iterator(); i.hasNext();){
            Client client = i.next();

            writer.println(client.getId() + ";"
                    + client.getNom() + ";"
                    + client.getPrenom() + ";"
                    + client.getDateNaissance() + ";"
                    + (now.getYear() - client.getDateNaissance().getYear())
            );


        }
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

    @GetMapping("/factures/xlsx")
    public void facturesXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Factures");

        Row headerRow = sheet.createRow(0);

        Cell cellHeaderId = headerRow.createCell(0);
        cellHeaderId.setCellValue("Id");

        int i = 1;
        //for (Client client : allClients) {
        //   Row row = sheet.createRow(i);

        // i++;
        //}

        workbook.write(response.getOutputStream());
        workbook.close();

    }

}
