package com.groggyman;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.groggyman.service.SheetsService;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;


/**
 * Examples are borrowed from https://www.baeldung.com/google-sheets-java-client
 *
 */
public class SheetsLiveTest {

    private static Sheets sheetsService;
    private static final String SPREADSHEET_ID = "INSERT ID (needs public access)";

    @BeforeClass
    public static void setup() throws IOException, GeneralSecurityException {
        sheetsService = SheetsService.getSheetsService(null);
    }

    /**
     * Values are written as rows starting from A1
     *
     * @throws IOException If execution is unsuccessful
     */
    @Test
    public void writeValueRangeToSpreadsheet() throws IOException {
        ValueRange body = new ValueRange()
                .setValues(Arrays.asList(
                        Arrays.asList("Expenses February"),
                        Arrays.asList("books", 30),
                        Arrays.asList("pens", 10),
                        Arrays.asList("Expenses March"),
                        Arrays.asList("clothes", 20),
                        Arrays.asList("shoes", 5)
                ));

        UpdateValuesResponse result = sheetsService.spreadsheets().values()
                .update(SPREADSHEET_ID, "A1", body)
                // write concise values (not computed)
                .setValueInputOption("RAW")
                .execute();
    }

    /**
     * Two value ranges are written in batch starting from D1 and D4
     *
     * @throws IOException If execution is unsuccessful
     */
    @Test
    public void writeValueRangesToSpreadsheetInBatch() throws IOException {
        List<ValueRange> data = new ArrayList<>();
        data.add(new ValueRange()
                .setRange("D1")
                .setValues(Arrays.asList(
                        Arrays.asList("February Total", "=B2+B3")
                ))
        );
        data.add(new ValueRange()
                .setRange("D4")
                .setValues(Arrays.asList(
                        Arrays.asList("March Total", "=B5+B6")
                ))
        );

        BatchUpdateValuesRequest batchBody = new BatchUpdateValuesRequest()
                // the cell values will be computed based on formula
                .setValueInputOption("USER_ENTERED")
                .setData(data);

        BatchUpdateValuesResponse batchResult = sheetsService.spreadsheets().values()
                .batchUpdate(SPREADSHEET_ID, batchBody)
                .execute();
    }

    /**
     * Appends computed cells at bottom of table and asserts update
     *
     * @throws IOException If execution is unsuccessful
     */
    @Test
    public void appendValuesAfterTable_andAssertResult() throws IOException {
        ValueRange appendBody = new ValueRange()
                .setValues(Arrays.asList(
                        Arrays.asList("Total", "=E1+E4")
                ));
        AppendValuesResponse appendResult = sheetsService.spreadsheets().values()
                .append(SPREADSHEET_ID, "A1", appendBody)
                .setValueInputOption("USER_ENTERED")
                // INSERT_ROWS option means that we want the data to be added to a new row, and not replace any existing data after the table
                .setInsertDataOption("INSERT_ROWS")
                // we want the response object to contain the updated data
                .setIncludeValuesInResponse(true)
                .execute();

        ValueRange total = appendResult.getUpdates().getUpdatedData();
        assertEquals(total.getValues().get(0).get(1), "65");
    }

    /**
     * Reads values range in batch or single cell and asserts
     *
     * @throws IOException If execution is unsuccessful
     */
    @Test
    public void readValues_andAssert() throws IOException {
        List<String> ranges = Arrays.asList("E1", "E4");
        BatchGetValuesResponse readResult = sheetsService.spreadsheets().values()
                .batchGet(SPREADSHEET_ID)
                .setRanges(ranges)
                .execute();

        ValueRange februaryTotal = readResult.getValueRanges().get(0);
        assertEquals(februaryTotal.getValues().get(0).get(0), "40");

        ValueRange marchTotal = readResult.getValueRanges().get(1);
        assertEquals(marchTotal.getValues().get(0).get(0), "25");

        ValueRange singleResult = sheetsService.spreadsheets().values()
                .get(SPREADSHEET_ID, "E1")
                .execute();
        assertEquals(singleResult.getValues().get(0).get(0), "40");
    }

    /**
     * Creates new spreadsheet in root directory with given name
     *
     * @throws IOException If execution is unsuccessful
     */
    @Test
    public void createNewSpreadsheet() throws IOException {
        Spreadsheet spreadSheet = new Spreadsheet().setProperties(
                new SpreadsheetProperties()
                        .setTitle("This is spreadsheet created by unit test"));

        Spreadsheet result = sheetsService
                .spreadsheets()
                .create(spreadSheet).execute();

        assertNotNull(result.getSpreadsheetId());
    }


    /**
     * Manipulating the spreadsheet: creating new sheet, copying data from one sheet to another,
     * updating the spreadsheet title and asserting the actions
     *
     * @throws IOException If execution is unsuccessful
     */
    @Test
    public void manipulatingSpreadsheet() throws IOException {
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(SPREADSHEET_ID).execute();

        UpdateSpreadsheetPropertiesRequest updateSpreadsheetPropertiesRequest
                = new UpdateSpreadsheetPropertiesRequest().setFields("*")
                .setProperties(new SpreadsheetProperties().setTitle("Expenses - API"));

        AddSheetRequest addSheetRequest = new AddSheetRequest()
                .setProperties(new SheetProperties()
                        .setTitle("AutomaticSheet" + spreadsheet.getSheets().size())
                        .setSheetId(1));

        CopyPasteRequest copyRequest = new CopyPasteRequest()
                .setSource(new GridRange().setSheetId(0)
                        .setStartColumnIndex(0).setEndColumnIndex(2)
                        .setStartRowIndex(0).setEndRowIndex(1))
                .setDestination(new GridRange().setSheetId(1)
                        .setStartColumnIndex(0).setEndColumnIndex(2)
                        .setStartRowIndex(0).setEndRowIndex(1))
                .setPasteType("PASTE_VALUES");

        List<Request> requests = new ArrayList<>();
        requests.add(new Request()
                .setAddSheet(addSheetRequest));
        requests.add(new Request()
                .setCopyPaste(copyRequest));
        requests.add(new Request()
                .setUpdateSpreadsheetProperties(updateSpreadsheetPropertiesRequest));

        BatchUpdateSpreadsheetRequest body
                = new BatchUpdateSpreadsheetRequest().setRequests(requests);

        sheetsService.spreadsheets().batchUpdate(SPREADSHEET_ID, body).execute();

        spreadsheet = sheetsService.spreadsheets().get(SPREADSHEET_ID).execute();
        assertEquals("Expenses - API", spreadsheet.getProperties().getTitle());
        assertEquals(2, spreadsheet.getSheets().size());
    }
}
