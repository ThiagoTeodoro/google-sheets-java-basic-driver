package br.com.google.sheets.java.basic.driver.model;

import lombok.AllArgsConstructor;
import lombok.Getter;

@AllArgsConstructor
@Getter
public class AccessData {
	
	private String pathCredentials;
	private String userToDisplay;
	private String pathToSaveToken;	
	private String applicationName;
	private String spreadsheetId;
	
}
