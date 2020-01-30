package br.com.google.sheets.java.basic.driver.service;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;

import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.AppendValuesResponse;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.DeleteDimensionRequest;
import com.google.api.services.sheets.v4.model.DimensionRange;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import br.com.google.sheets.java.basic.driver.model.AccessData;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Slf4j
@Getter
public class GoogleSheetsService {

	private AccessData actualAccessData;	
	
	/**
	 * Service to check if all attributes are filled in AccessData
	 * 
	 * @param accessData
	 * @return
	 */
	private static boolean accessDataIsValid(AccessData accessData) {
		
		if(accessData.getPathCredentials() == null || accessData.getPathCredentials().isEmpty()) {
			log.error("AccessData inválido, favor conferir o envio dos parâmetros.");
			return false;
		}
		
		if(accessData.getUserToDisplay() == null || accessData.getUserToDisplay().isEmpty()) {
			log.error("AccessData inválido, favor conferir o envio dos parâmetros.");
			return false;
		}
		
		if(accessData.getPathToSaveToken() == null || accessData.getPathToSaveToken().isEmpty()) {
			log.error("AccessData inválido, favor conferir o envio dos parâmetros.");
			return false;
		}
		
		if(accessData.getApplicationName() == null || accessData.getApplicationName().isEmpty()) {
			log.error("AccessData inválido, favor conferir o envio dos parâmetros.");
			return false;
		}
		
		if(accessData.getSpreadsheetId() == null || accessData.getSpreadsheetId().isEmpty()) {
			log.error("AccessData inválido, favor conferir o envio dos parâmetros.");
			return false;
		}
		
		return true;				
	}	
	
	/**
	 * Initializer Driver 
	 * 
	 * @param pathCredentials
	 * @param userToDisplay
	 * @param pathToSaveToken
	 * @param applicationName
	 * @param spreadsheetId
	 * @return
	 */
	public static GoogleSheetsService initializer(String pathCredentials, String userToDisplay, String pathToSaveToken, String applicationName, String spreadsheetId) {
		
		AccessData accessData = new AccessData(pathCredentials, userToDisplay, pathToSaveToken, applicationName, spreadsheetId);
			
		if(accessDataIsValid(accessData)) {						
			
			GoogleSheetsService googleSheetsService = new GoogleSheetsService();
			googleSheetsService.actualAccessData = accessData;			
			return googleSheetsService;		
			
		} else {
								
			return null;
		}
			
	}

	/**
	 * Service to authenticate in Google API.
	 * @return
	 */
	private Credential authorize() {

		try {

			InputStream inputStream = GoogleSheetsService.class.getResourceAsStream(actualAccessData.getPathCredentials());
			GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JacksonFactory.getDefaultInstance(),
					new InputStreamReader(inputStream));

			List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);

			GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
					GoogleNetHttpTransport.newTrustedTransport(),
					JacksonFactory.getDefaultInstance(),
					clientSecrets, scopes)
					.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(actualAccessData.getPathToSaveToken())))
					.setAccessType("offline").build();

			return new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize(actualAccessData.getUserToDisplay());			
		} catch (Exception e) {

			log.error("Ocorreu um erro ao tentar autorizar no GoogleSheets Service. Exception : ", e);
			return null;
		}
	}

	/**
	 * Service to get API Sheets. 
	 * 
	 * @return
	 */
	public Sheets getSheetsService() {

		try {

			Credential credential = authorize();
			return new Sheets.Builder(
					GoogleNetHttpTransport.newTrustedTransport(),
					JacksonFactory.getDefaultInstance(),credential)
					.setApplicationName(actualAccessData.getApplicationName()).build();

		} catch (Exception e) {

			log.error("Ocorreu um erro ao tentar obter o SheetsService no GoogleSheets Service, quando tentava executar getSheetsService(). Exception : ", e);
			return null;
		}

	}
	
	/**
	 * Service to update data in spreadsheet
	 * 
	 * @param values 
	 * @param target 
	 * @return
	 */
	public UpdateValuesResponse updateValueInSheets(List<Object> values, String target) {
		
		try {
			
			if(accessDataIsValid(actualAccessData)) {
				
				ValueRange data = new ValueRange().setValues(Arrays.asList(values));
				
				return	getSheetsService().spreadsheets().values()
						.update(getActualAccessData().getSpreadsheetId(), target, data)
						.setValueInputOption("USER_ENTERED")
						.execute();
				
			} else {

				return null;
			}
		} catch (Exception e) {
			
			log.error("Ocorreu um erro ao tentar realizar o update dos dados na planilha : " + actualAccessData.getSpreadsheetId(), e);
			return null;
		}
		
	}	
	
	/**
	 * Servce to read values in Sheets.
	 * 
	 * @param range
	 * @return
	 */
	public List<List<Object>> readValueInSheets(String range){
		
		try {
			
			ValueRange response = getSheetsService().spreadsheets()
					.values()
					.get(getActualAccessData().getSpreadsheetId(), range)
					.execute();
			
			List<List<Object>> values = response.getValues();
			
			if(values == null || values.isEmpty()) {
				
				log.info("Nenhum dado encontrado!");
				return new ArrayList<List<Object>>();			
			} else {
				
				return values;
			}			
		} catch (Exception e) {
			
			log.error("Ocorreu um erro quando driver tentou ler os dados da planilha em readValueInSheets(). Exception : " + e);
			return null;
		}
		
	}
	
	
	/**
	 * Service to insert new Rows in end of file.
	 * 
	 * @param dataRow
	 * @param nameSheet
	 * @return
	 */
	public AppendValuesResponse insertRowInSheets(List<Object> dataRow, String nameSheet) {
		
		try {
		
			ValueRange row = new ValueRange().setValues(Arrays.asList(dataRow));
			
			return getSheetsService().spreadsheets().values()
					.append(getActualAccessData().getSpreadsheetId(), nameSheet, row)
					.setValueInputOption("USER_ENTERED")
					.setInsertDataOption("INSERT_ROWS")
					.setIncludeValuesInResponse(true)
					.execute();			
			
		} catch (Exception e) {
			
			log.error("Ocorreu um erro quando driver tentou inserir a linha em insertRowInSheets(). Exception : " + e);
			return null;
		}
	}
	
	/**
	 * Service to delete Rows in SpreedSheet
	 * 
	 * @param startRow Row Start 
	 * @param end Row end, if null, will delete all lines from the start line. 
	 */
	public void deleteRowsInSheets(Integer startRow, Integer endRow, int sheetId) {

		try {
			
			DeleteDimensionRequest deleteRequest = new DeleteDimensionRequest()
					.setRange(
								new DimensionRange()
								.setSheetId(sheetId)
								.setDimension("ROWS")
								.setStartIndex(startRow)
								.setEndIndex(endRow));
	
			List<Request> requests = new ArrayList<>();
			
			requests.add(new Request().setDeleteDimension(deleteRequest));
	
			BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
			
			getSheetsService().spreadsheets().batchUpdate(getActualAccessData().getSpreadsheetId(), body).execute();
			
		} catch (Exception e) {
			
			log.error("Ocorreu um erro quando o driver tentou excluir as linhas da planilha em deleteRowsInSheets()", e);			
		}

	}
	

}
