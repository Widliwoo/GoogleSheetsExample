package com.example;

import com.example.util.SheetsServiceUtil;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class SheetsQuickstart {

    public static void main(String... args) throws IOException, GeneralSecurityException {
        // Build a new authorized API client service.
        final String spreadsheetId = "1kekoF6RyIGg75TuIQeeGnK-TpS0N7Cogf8npxwW92do";
        Sheets service = SheetsServiceUtil.getSheetsService();

        writeDataToSheet(service, spreadsheetId);
        appendDataToSheet(service, spreadsheetId);
        readAndPrintAllValues(service, spreadsheetId);
        deleteRow(service, spreadsheetId);
    }

    private static void deleteRow(Sheets service, String spreadsheetId) throws IOException {
        DimensionRange dimensionRange = new DimensionRange()
                .setSheetId(0)
                .setDimension("ROWS")
                .setStartIndex(10);
        DeleteDimensionRequest deleteDimensionRequest = new DeleteDimensionRequest()
                .setRange(dimensionRange);
        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setDeleteDimension(deleteDimensionRequest));
        BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);

        service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateSpreadsheetRequest).execute();
    }

    private static void readAndPrintAllValues(Sheets service, String spreadsheetId) throws IOException {
        System.out.println("Trying to read data from sheet");
        for (List<Object> row : service.spreadsheets().values()
                .get(spreadsheetId, "A:E")
                .execute()
                .getValues()) {
            for (Object cell : row) {
                System.out.print(cell + ", ");
            }
            System.out.println();
        }
    }

    private static void appendDataToSheet(Sheets service, String spreadsheetId) throws IOException {
        ValueRange appendBody = new ValueRange()
                .setValues(
                        List.of(
                                List.of("TEST", "test", "TEST"),
                                List.of("TEST", "test", "TEST"),
                                List.of("TEST", "test", "TEST"),
                                List.of("TEST", "test", "TEST"),
                                List.of("TEST", "test", "TEST"),
                                List.of("TEST", "test", "TEST")
                        )
                );
        service
                .spreadsheets()
                .values()
                .append(spreadsheetId, "A1", appendBody)
                .setValueInputOption("USER_ENTERED")
                .setInsertDataOption("INSERT_ROWS")
                .setIncludeValuesInResponse(true)
                .execute();
    }

    private static void writeDataToSheet(Sheets service, String spreadsheetId) throws IOException {
        List<ValueRange> data = new ArrayList<>();
        ValueRange range = new ValueRange();
        range.setValues(
                Arrays.asList(
                        Arrays.asList("col_1", "col_2", "col_3"),
                        Arrays.asList(1, 2, 3),
                        Arrays.asList(11, 22, 33),
                        Arrays.asList(111, 222, 333),
                        Arrays.asList("TEXT", "ТЕКСТ", "ФЫВРФВОЛРФЫ")
                )
        ).setRange("A1");
        data.add(range);

        BatchUpdateValuesRequest batch = new BatchUpdateValuesRequest()
                .setValueInputOption("USER_ENTERED")
                .setData(data);

        service.spreadsheets()
                .values()
                .batchUpdate(spreadsheetId, batch)
                .execute();

    }
}