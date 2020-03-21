package com.groggyman.service;

import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.sheets.v4.Sheets;

import javax.annotation.Nullable;
import java.io.IOException;
import java.security.GeneralSecurityException;

public class SheetsService {

    private static final String APPLICATION_NAME = "Groggyman Sheets Example";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

    /**
     * Creates Sheets service object used for reading and writing Google Sheets
     *
     * @return Sheets service
     * @throws GeneralSecurityException if there are HttpTransport issues
     * @throws IOException              If credentials cannot be found.
     */
    public static Sheets getSheetsService(@Nullable String applicationName) throws IOException, GeneralSecurityException {
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

        return new Sheets.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, CredentialService.getCredentials(HTTP_TRANSPORT))
                .setApplicationName(applicationName != null ? applicationName : APPLICATION_NAME)
                .build();
    }
}
