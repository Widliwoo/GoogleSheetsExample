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
        // Из чего состоит ссылка на документ:
        // https://docs.google.com/spreadsheets/d/{spreadsheetId}/edit#gid={sheetId}

        // Использованный пример
        // https://docs.google.com/spreadsheets/d/1kekoF6RyIGg75TuIQeeGnK-TpS0N7Cogf8npxwW92do/edit#gid=0
        final String spreadsheetId = "1kekoF6RyIGg75TuIQeeGnK-TpS0N7Cogf8npxwW92do";
        Sheets service = SheetsServiceUtil.getSheetsService();

        // Запись данных в таблицу
        writeDataToSheet(service, spreadsheetId);

        // Добавление данных в конец заполненной области таблицы
        appendDataToSheet(service, spreadsheetId);

        // Получение ячеек таблицы
        readAndPrintAllValues(service, spreadsheetId);

        // удаление строк
        deleteRow(service, spreadsheetId);
    }

    private static void deleteRow(Sheets service, String spreadsheetId) throws IOException {

        // Последовтаельность строк или столбцов, которые будут удалены.
        DimensionRange dimensionRange = new DimensionRange()

                // идентификатор листа, на котором будут удалятся строки/столбцы
                // {sheetId} из примера
                .setSheetId(0)

                // Указание измерения, которое будет подвергнуто удалению
                // ROWS - строки
                // COLUMNS - столбцы
                .setDimension("ROWS")

                // Начиная с индекса 10 (включительно) и до endIndex() (не включительно)
                .setStartIndex(10)
                .setEndIndex(100);

        // Создание запроса на удаление
        DeleteDimensionRequest deleteDimensionRequest = new DeleteDimensionRequest()
                .setRange(dimensionRange);

        // Создание списка запросов на удаление, для пакетной обработки
        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setDeleteDimension(deleteDimensionRequest));

        // создание пакетного запроса на удаление
        BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest()
                .setRequests(requests);

        // вызов процедуры пакетного удаления
        service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateSpreadsheetRequest).execute();
    }

    private static void readAndPrintAllValues(Sheets service, String spreadsheetId) throws IOException {
        System.out.println("Trying to read data from sheet");

        // Запросили у таблицы все значения в столбцах [А-Е]
        // Получили список списков. Список строк, каждая из которых является списком ячеек.
        List<List<Object>> spreadsheetValues = service.spreadsheets().values()
                .get(spreadsheetId, "A:E")
                .execute()
                .getValues();

        for (List<Object> row : spreadsheetValues) {
            for (Object cell : row) {
                System.out.print(cell.toString() + ", ");
            }
            System.out.println();
        }
    }

    private static void appendDataToSheet(Sheets service, String spreadsheetId) throws IOException {
        // Создаём последовательность с данными, которая представлена списком строк,
        // каждая из которых представлена списком ячеек
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

                // передаём последовательность для дозаписи,
                // указывая начало таблицы (А1), для которой нужно выполнить дозапись.
                .append(spreadsheetId, "A1", appendBody)

                // с этими двумя параметрами пока не успел разобраться
                .setValueInputOption("USER_ENTERED")
                .setInsertDataOption("INSERT_ROWS")

                // includeValuesInResponse. Если true, то в ответе от сервера прилетят значения,
                // которые мы дозаписали в таблицу
                .setIncludeValuesInResponse(true)
                .execute();
    }

    private static void writeDataToSheet(Sheets service, String spreadsheetId) throws IOException {
        // Список последовательностей, который будет посылаться на сервер
        List<ValueRange> data = new ArrayList<>();

        /**
         * ValueRange - объект, который представляет  последовательность ячеек в таблице.
         */
        ValueRange range = new ValueRange();
        range.setValues(

                // список строк
                Arrays.asList(
                        // список ячеек в первой строке последовательности
                        Arrays.asList("col_1", "col_2", "col_3"),

                        // список ячеек во второй строке последовательности
                        Arrays.asList(1, 2, 3),
                        Arrays.asList(11, 22, 33),
                        Arrays.asList(111, 222, 333),
                        Arrays.asList("TEXT", "ТЕКСТ", "ФЫВРФВОЛРФЫ")
                )
        ).setRange("A1");
        // SetRange(...) позволяет указать ячейку, в которой начинается левый верхний угол последовательности

        // Добавили последовательность в список для последующей отправки
        data.add(range);

        // Создание запроса на множественное обновление
        BatchUpdateValuesRequest batch = new BatchUpdateValuesRequest()
                .setValueInputOption("USER_ENTERED")
                .setData(data);


        service.spreadsheets()
                .values()

                // Указываем на какой таблице хотим применить набор обновлений
                .batchUpdate(spreadsheetId, batch)

                // Вызываем метод, который выполнит все обновления
                .execute();

    }
}