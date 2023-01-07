package com.example;

import com.example.util.SheetsServiceUtil;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class SheetsQuickstart {
    private static final JsonFactory JSON_FACTORY = GsonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";

    /**
     * Global instance of the scopes required by this quickstart.
     * If modifying these scopes, delete your previously saved tokens/ folder.
     */
    private static final List<String> SCOPES =
            Collections.singletonList(SheetsScopes.SPREADSHEETS_READONLY);
    private static final String CREDENTIALS_FILE_PATH = "/secret/credentials.json";

    /**
     * Creates an authorized Credential object.
     *
     * @param HTTP_TRANSPORT The network HTTP Transport.
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found.
     */
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT)
            throws IOException {
        // Load client secrets.
        InputStream in = SheetsQuickstart.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets =
                GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    /**
     * Prints the names and majors of students in a sample spreadsheet:
     * https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
     */
    public static void main(String... args) throws IOException, GeneralSecurityException {
        // Build a new authorized API client service.
        final String spreadsheetId = "1kekoF6RyIGg75TuIQeeGnK-TpS0N7Cogf8npxwW92do";
        Sheets service = SheetsServiceUtil.getSheetsService();

        writeDataToSheet(service, spreadsheetId);
        appendDataToSheet(service, spreadsheetId);
        readAndPrintAllValues(service, spreadsheetId);
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