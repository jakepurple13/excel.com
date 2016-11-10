package com.excel.excel.com;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.time.DateTimeException;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.ValueRange;

public class GoogleSheets {

	/** Application name. */
	private final String APPLICATION_NAME = "Google Sheets API Java Quickstart";

	String sheetID = "1jQzehm5To1ZDM-m0B3mI6EfFChny6rFCDlP2Svn_6o0";

	/** Directory to store user credentials for this application. */
	private final static java.io.File DATA_STORE_DIR = new java.io.File(System.getProperty("user.home"),
			".credentials/sheets.googleapis.com-java-quickstart");

	/** Global instance of the {@link FileDataStoreFactory}. */
	private static FileDataStoreFactory DATA_STORE_FACTORY;

	/** Global instance of the JSON factory. */
	private final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

	/** Global instance of the HTTP transport. */
	private static HttpTransport HTTP_TRANSPORT;

	ArrayList<String> als = new ArrayList<>();

	/**
	 * Global instance of the scopes required by this quickstart.
	 *
	 * If modifying these scopes, delete your previously saved credentials at
	 * ~/.credentials/sheets.googleapis.com-java-quickstart
	 */
	private final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS_READONLY);

	static {
		try {
			HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
		} catch (Throwable t) {
			t.printStackTrace();
			System.exit(1);
		}
	}

	public GoogleSheets() throws IOException {
		// Build a new authorized API client service.
		Sheets service = getSheetsService();

		// Prints the names and majors of students in a sample spreadsheet:
		// https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
		String spreadsheetId = sheetID;

		Spreadsheet response1 = service.spreadsheets().get(spreadsheetId).setIncludeGridData(false).execute();

		List<Sheet> workSheetList = response1.getSheets();

		for (Sheet sheet : workSheetList) {
			System.err.println(sheet.getProperties().getTitle());
			String sheetName = sheet.getProperties().getTitle();
			String range = sheetName + "!A1:BG100";// "A4:E20";//BG
			ValueRange response = service.spreadsheets().values().get(spreadsheetId, range).execute();
			List<List<Object>> values = response.getValues();

			if (values == null || values.size() == 0) {
				System.out.println("No data found.");
			} else {
				// System.out.println("Size is: " + values.size());

				for (int i = 0; i < values.size(); i++) {

					List<Object> lo = values.get(i);

					for (int j = 0; j < lo.size(); j++) {
						System.out.print(lo.get(j) + "\t");
					}

					System.out.println();

					if (i >= 4) {

						switch (sheetName) {
						case "Summary":

							break;

						case "supervisors":

							break;

						default:
							// System.err.println("-------------------");
							// als.add("-------------------");
							als.add(lo.get(0) + "\t" + sheetName + "\t" + lo.get(1) + "\t" + lo.get(2));
							try {
								DateTimeFormatter parseFormat = new DateTimeFormatterBuilder().appendPattern("hh:mm a")
										.toFormatter();

								LocalTime lt = LocalTime.parse(lo.get(1).toString(), parseFormat);
								System.err.println(lt);
							} catch (DateTimeException e) {

							}
							// als.add("-------------------");
							// System.out.println(lo.get(0) + "\t" + sheetName +
							// "\t" + lo.get(1) + "\t" + lo.get(2));
							// System.err.println("-------------------");
							break;
						}

					}

				}

				/*
				 * for (List row : values) { // Print columns A and E, which
				 * correspond to indices 0 and 4.
				 * //System.out.printf("%s, %s\n", row.get(0), row.get(2));
				 * System.out.println(row.get(0)); }
				 */
			}

		}
	}

	public ArrayList<String> getList() {
		return als;
	}

	/**
	 * Creates an authorized Credential object.
	 * 
	 * @return an authorized Credential object.
	 * @throws IOException
	 */
	public Credential authorize() throws IOException {
		// Load client secrets.
		InputStream in = Quickstart.class.getResourceAsStream("/client_secret.json");
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY,
				clientSecrets, SCOPES).setDataStoreFactory(DATA_STORE_FACTORY).setAccessType("offline").build();
		Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
		System.out.println("Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
		return credential;
	}

	/**
	 * Build and return an authorized Sheets API client service.
	 * 
	 * @return an authorized Sheets API client service
	 * @throws IOException
	 */
	public Sheets getSheetsService() throws IOException {
		Credential credential = authorize();
		return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential).setApplicationName(APPLICATION_NAME)
				.build();
	}

}
