package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.telegram.telegrambots.bots.TelegramLongPollingBot;
import org.telegram.telegrambots.meta.api.objects.Message;
import org.telegram.telegrambots.meta.api.objects.Update;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Bot extends TelegramLongPollingBot {
    private static int userCount = 0;

    public Bot(String token) {
        super("YOUR_BOT_TOKEN");
    }

    public void saveToExcel(Update update) {
        String filePath = "users.xlsx";
        File file = new File(filePath);
        Workbook workbook;
        Sheet sheet;

        try {
            if (file.exists()) {
                FileInputStream inputStream = new FileInputStream(file);
                workbook = new XSSFWorkbook(inputStream);
                sheet = workbook.getSheetAt(0);
                inputStream.close();
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Users Data");
                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("ID");
                headerRow.createCell(1).setCellValue("Username");
                headerRow.createCell(2).setCellValue("First Name");
                headerRow.createCell(3).setCellValue("Last Name");
                headerRow.createCell(4).setCellValue("Message");
                headerRow.createCell(5).setCellValue("Phone Number");
                headerRow.createCell(8).setCellValue("Comment");
            }

            long userId = update.getMessage().getFrom().getId();
            String username = update.getMessage().getFrom().getUserName();
            String firstName = update.getMessage().getFrom().getFirstName();
            String lastName = update.getMessage().getFrom().getLastName();
            String message = update.getMessage().hasText() ? update.getMessage().getText() : "N/A";
            String phoneNumber = "N/A";
            if (update.getMessage().hasContact()) {
                phoneNumber = update.getMessage().getContact().getPhoneNumber();
            }


            int lastRow = sheet.getLastRowNum() + 1;
            Row newRow = sheet.createRow(lastRow);
            newRow.createCell(0).setCellValue(userId);
            newRow.createCell(1).setCellValue(username != null ? username : "N/A");
            newRow.createCell(2).setCellValue(firstName != null ? firstName : "N/A");
            newRow.createCell(3).setCellValue(lastName != null ? lastName : "N/A");
            newRow.createCell(4).setCellValue(message);
            newRow.createCell(5).setCellValue(phoneNumber);

            FileOutputStream outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            userCount++;
            System.out.println("Excelga saqlandi oka");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void onUpdateReceived(Update update) {
        Message message = update.getMessage();
        saveToExcel(update);

    }

    @Override
    public String getBotUsername() {
        return "YOUR_BOT_USERNAME";
    }


}
