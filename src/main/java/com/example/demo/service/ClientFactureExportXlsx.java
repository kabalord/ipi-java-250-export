package com.example.demo.service;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.ServletOutputStream;
import java.io.IOException;
import java.util.List;

@Service
public class ClientFactureExportXlsx {

    public static final int IDX_PRIX_LIGNE = 3;

    @Autowired
    private ClientService clientService;

    public static final int IDX_ARTICLE = 0;

    public void facturesXlsx(ServletOutputStream outputStream) throws IOException {
        List<Client> allClients = clientService.findAllClients();
        Workbook workbook = new XSSFWorkbook();
        CellStyle styleBold = createStyleBold(workbook);
        for (Client client : allClients) {
            createSheetClient(workbook, client);
            for (Facture facture : client.getFactures()) {
                createSheetFacture(workbook, styleBold, facture);
            }
        }
        workbook.write(outputStream);
        workbook.close();
    }

    private void createSheetClient(Workbook workbook, Client client) {
        workbook.createSheet("Client " + client.getNom());
    }

    /**
     * @param workbook
     * @param styleBold
     * @param facture
     */
    private void createSheetFacture(Workbook workbook, CellStyle styleBold, Facture facture) {
        Sheet sheet = workbook.createSheet("Facture " + facture.getId());
        createFactureHeader(sheet);
        int iRow = 1;
        for (LigneFacture ligneFacture : facture.getLigneFactures()) {
            Row row = sheet.createRow(iRow);

            Cell cellArticle = row.createCell(IDX_ARTICLE);
            cellArticle.setCellValue(ligneFacture.getArticle().getLibelle());

            Cell cellQte = row.createCell(1);
            cellQte.setCellValue(ligneFacture.getQuantite());

            Cell cellPrixUnitaire = row.createCell(2);
            cellPrixUnitaire.setCellValue(ligneFacture.getArticle().getPrix());

            Cell cellPrixLigne = row.createCell(IDX_PRIX_LIGNE);
            cellPrixLigne.setCellValue(ligneFacture.getQuantite() * ligneFacture.getArticle().getPrix());

            iRow++;
        }
        Row rowTotal = sheet.createRow(iRow);
        Cell cellTotal = rowTotal.createCell(3);
        cellTotal.setCellValue(facture.getTotal());
        cellTotal.setCellStyle(styleBold);
    }

    private CellStyle createStyleBold(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        cellStyle.setFont(font);
        //je mets en gras
        font.setBold(true);
        return cellStyle;
    }

    private void createFactureHeader(Sheet sheet) {
        Row header = sheet.createRow(0);
        Cell cellHeaderArticle = header.createCell(IDX_ARTICLE);
        cellHeaderArticle.setCellValue("Article");

        Cell cellHeaderQte = header.createCell(1);
        cellHeaderQte.setCellValue("Quantité");

        Cell cellHeaderPrixUnitaire = header.createCell(2);
        cellHeaderPrixUnitaire.setCellValue("Prix Unitaire");

        Cell cellHeaderPrixLigne = header.createCell(IDX_PRIX_LIGNE);
        cellHeaderPrixLigne.setCellValue("Prix ligne");
    }

}